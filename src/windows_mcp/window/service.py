"""Window enumeration, status queries, and focus management.

Stateless service (except process-name cache) providing window
discovery, overlay/browser detection, foreground management,
and bring-to-top operations.
"""

import ctypes
import logging
import threading
from contextlib import contextmanager

import win32con
import win32gui
import win32process
from psutil import Process

import windows_mcp.uia as uia
from windows_mcp.desktop.views import BoundingBox, Browser, Status, Window
from windows_mcp.vdm.core import is_window_on_current_desktop

logger = logging.getLogger(__name__)

_MAX_PARENT_DEPTH = 64
_PROCESS_CACHE_MAX = 512


class WindowService:
    """Window enumeration, status, and focus management."""

    def __init__(self):
        self._process_name_cache: dict[int, str] = {}
        self._process_cache_lock = threading.Lock()

    def get_window_status(self, control: uia.Control) -> Status:
        """Return the status (minimized/maximized/normal/hidden) of a window."""
        if uia.IsIconic(control.NativeWindowHandle):
            return Status.MINIMIZED
        elif uia.IsZoomed(control.NativeWindowHandle):
            return Status.MAXIMIZED
        elif uia.IsWindowVisible(control.NativeWindowHandle):
            return Status.NORMAL
        else:
            return Status.HIDDEN

    def is_overlay_window(self, element: uia.Control) -> bool:
        """Return True if the element is an overlay (e.g. NVIDIA, Steam)."""
        no_children = len(element.GetChildren()) == 0
        is_name = "Overlay" in (element.Name or "").strip()
        return no_children or is_name

    def is_window_browser(self, node: uia.Control) -> bool:
        """Return True if the UIA control belongs to a browser process."""
        try:
            pid = node.ProcessId
            with self._process_cache_lock:
                proc_name = self._process_name_cache.get(pid)
            if proc_name is None:
                proc_name = Process(pid).name()
                with self._process_cache_lock:
                    if len(self._process_name_cache) >= _PROCESS_CACHE_MAX:
                        self._process_name_cache.clear()
                    self._process_name_cache[pid] = proc_name
            return Browser.has_process(proc_name)
        except Exception:
            logger.debug("Failed to determine browser status for node", exc_info=True)
            return False

    def get_controls_handles(self, optimized: bool = False) -> set[int]:
        """Enumerate visible window handles on the current virtual desktop."""
        handles: set[int] = set()

        def callback(hwnd, _):
            try:
                if (
                    win32gui.IsWindow(hwnd)
                    and win32gui.IsWindowVisible(hwnd)
                    and is_window_on_current_desktop(hwnd)
                ):
                    handles.add(hwnd)
            except Exception:
                pass

        win32gui.EnumWindows(callback, None)

        if desktop_hwnd := win32gui.FindWindow("Progman", None):
            handles.add(desktop_hwnd)
        if taskbar_hwnd := win32gui.FindWindow("Shell_TrayWnd", None):
            handles.add(taskbar_hwnd)
        if secondary_taskbar_hwnd := win32gui.FindWindow("Shell_SecondaryTrayWnd", None):
            handles.add(secondary_taskbar_hwnd)
        return handles

    def get_foreground_window(self) -> uia.Control:
        """Return the UIA control for the current foreground window."""
        handle = uia.GetForegroundWindow()
        return self.get_window_from_element_handle(handle)

    def get_window_from_element_handle(self, element_handle: int) -> uia.Control:
        """Walk the UIA parent chain from a handle to find the top-level window."""
        current = uia.ControlFromHandle(element_handle)
        root_handle = uia.GetRootControl().NativeWindowHandle

        for _ in range(_MAX_PARENT_DEPTH):
            parent = current.GetParentControl()
            if parent is None or parent.NativeWindowHandle == root_handle:
                return current
            current = parent
        logger.warning(
            "get_window_from_element_handle exceeded depth limit (%d) for handle %s",
            _MAX_PARENT_DEPTH,
            element_handle,
        )
        return current

    def get_active_window(self, windows: list[Window] | None = None) -> Window | None:
        """Return the active (foreground) Window, or None if not found."""
        try:
            if windows is None:
                windows, _ = self.get_windows()
            active_window = self.get_foreground_window()
            if active_window.ClassName == "Progman":
                return None
            active_window_handle = active_window.NativeWindowHandle
            for window in windows:
                if window.handle != active_window_handle:
                    continue
                return window
            # In case active window is not present in the windows list
            return Window(
                **{
                    "name": active_window.Name,
                    "is_browser": self.is_window_browser(active_window),
                    "depth": 0,
                    "bounding_box": BoundingBox(
                        left=active_window.BoundingRectangle.left,
                        top=active_window.BoundingRectangle.top,
                        right=active_window.BoundingRectangle.right,
                        bottom=active_window.BoundingRectangle.bottom,
                        width=active_window.BoundingRectangle.width(),
                        height=active_window.BoundingRectangle.height(),
                    ),
                    "status": self.get_window_status(active_window),
                    "handle": active_window_handle,
                    "process_id": active_window.ProcessId,
                }
            )
        except Exception as ex:
            logger.error("Error in get_active_window: %s", ex)
        return None

    def get_windows(
        self, controls_handles: set[int] | None = None
    ) -> tuple[list[Window], set[int]]:
        """Enumerate all visible, non-overlay windows on the current desktop."""
        try:
            windows: list[Window] = []
            window_handles: set[int] = set()
            controls_handles = controls_handles or self.get_controls_handles()
            for idx, hwnd in enumerate(controls_handles):
                try:
                    child = uia.ControlFromHandle(hwnd)
                except Exception:
                    continue

                if self.is_overlay_window(child):
                    continue

                if isinstance(child, (uia.WindowControl, uia.PaneControl)):
                    window_pattern = child.GetPattern(uia.PatternId.WindowPattern)
                    if window_pattern is None:
                        continue

                    if window_pattern.CanMinimize and window_pattern.CanMaximize:
                        status = self.get_window_status(child)

                        bounding_rect = child.BoundingRectangle
                        if bounding_rect.isempty() and status != Status.MINIMIZED:
                            continue

                        windows.append(
                            Window(
                                **{
                                    "name": child.Name,
                                    "depth": idx,
                                    "status": status,
                                    "bounding_box": BoundingBox(
                                        left=bounding_rect.left,
                                        top=bounding_rect.top,
                                        right=bounding_rect.right,
                                        bottom=bounding_rect.bottom,
                                        width=bounding_rect.width(),
                                        height=bounding_rect.height(),
                                    ),
                                    "handle": child.NativeWindowHandle,
                                    "process_id": child.ProcessId,
                                    "is_browser": self.is_window_browser(child),
                                }
                            )
                        )
                        window_handles.add(child.NativeWindowHandle)
        except Exception as ex:
            logger.error("Error in get_windows: %s", ex)
            windows = []
        return windows, window_handles

    def get_window_from_element(self, element: uia.Control) -> Window | None:
        """Return the Window object containing the given UIA element."""
        if element is None:
            return None
        top_window = element.GetTopLevelControl()
        if top_window is None:
            return None
        handle = top_window.NativeWindowHandle
        windows, _ = self.get_windows()
        for window in windows:
            if window.handle == handle:
                return window
        return None

    def bring_window_to_top(self, target_handle: int):
        """Bring a window to the foreground using Win32 thread attachment."""
        if not win32gui.IsWindow(target_handle):
            raise ValueError("Invalid window handle")

        try:
            if win32gui.IsIconic(target_handle):
                win32gui.ShowWindow(target_handle, win32con.SW_RESTORE)

            foreground_handle = win32gui.GetForegroundWindow()

            if not win32gui.IsWindow(foreground_handle):
                win32gui.SetForegroundWindow(target_handle)
                win32gui.BringWindowToTop(target_handle)
                return

            foreground_thread, _ = win32process.GetWindowThreadProcessId(foreground_handle)
            target_thread, _ = win32process.GetWindowThreadProcessId(target_handle)

            if not foreground_thread or not target_thread or foreground_thread == target_thread:
                win32gui.SetForegroundWindow(target_handle)
                win32gui.BringWindowToTop(target_handle)
                return

            ctypes.windll.user32.AllowSetForegroundWindow(-1)

            attached = False
            try:
                win32process.AttachThreadInput(foreground_thread, target_thread, True)
                attached = True

                win32gui.SetForegroundWindow(target_handle)
                win32gui.BringWindowToTop(target_handle)

                win32gui.SetWindowPos(
                    target_handle,
                    win32con.HWND_TOP,
                    0,
                    0,
                    0,
                    0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW,
                )

            finally:
                if attached:
                    win32process.AttachThreadInput(foreground_thread, target_thread, False)

        except Exception as e:
            logger.exception("Failed to bring window to top: %s", e)

    @contextmanager
    def auto_minimize(self):
        """Context manager that minimizes the foreground window and restores on exit."""
        handle = uia.GetForegroundWindow()
        if not handle:
            yield
            return
        try:
            uia.ShowWindow(handle, win32con.SW_MINIMIZE)
            yield
        finally:
            uia.ShowWindow(handle, win32con.SW_RESTORE)
