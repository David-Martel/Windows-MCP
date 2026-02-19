import csv
import ctypes
import io
import logging
import os
import re
import threading
from contextlib import contextmanager
from locale import getpreferredencoding
from time import time
from typing import Literal

import win32con
from PIL import Image
from thefuzz import process

from windows_mcp.desktop.config import PROCESS_PER_MONITOR_DPI_AWARE
from windows_mcp.desktop.views import DesktopState, Size, Status, Window
from windows_mcp.tree.service import Tree
from windows_mcp.tree.views import TreeElementNode
from windows_mcp.vdm.core import (
    get_all_desktops,
    get_current_desktop,
)

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

import windows_mcp.uia as uia  # noqa: E402


class Desktop:
    def __init__(self):
        from windows_mcp.input import InputService
        from windows_mcp.registry import RegistryService
        from windows_mcp.scraper import ScraperService
        from windows_mcp.screen import ScreenService
        from windows_mcp.shell import ShellService
        from windows_mcp.window import WindowService

        self.encoding = getpreferredencoding()
        self._input = InputService()
        self._screen = ScreenService()
        self._window = WindowService()
        self._registry = RegistryService()
        self._shell = ShellService()
        self._scraper = ScraperService()
        self.tree = Tree(self)
        self.desktop_state = None
        self._state_lock = threading.Lock()
        # Cache for start menu app list (avoids repeated PowerShell subprocess)
        self._app_cache: dict[str, str] | None = None
        self._app_cache_time: float = 0.0
        self._app_cache_lock = threading.Lock()
        self._APP_CACHE_TTL: float = 3600.0  # 1 hour

    @staticmethod
    def _ps_quote(value: str) -> str:
        from windows_mcp.shell import ShellService

        return ShellService.ps_quote(value)

    def get_state(
        self,
        use_annotation: bool | str = True,
        use_vision: bool | str = False,
        use_dom: bool | str = False,
        as_bytes: bool | str = False,
        scale: float = 1.0,
    ) -> DesktopState:
        use_annotation = use_annotation is True or (
            isinstance(use_annotation, str) and use_annotation.lower() == "true"
        )
        use_vision = use_vision is True or (
            isinstance(use_vision, str) and use_vision.lower() == "true"
        )
        use_dom = use_dom is True or (isinstance(use_dom, str) and use_dom.lower() == "true")
        as_bytes = as_bytes is True or (isinstance(as_bytes, str) and as_bytes.lower() == "true")

        if not (0.1 <= scale <= 4.0):
            raise ValueError(f"scale must be between 0.1 and 4.0, got {scale}")

        start_time = time()

        controls_handles = self.get_controls_handles()  # Taskbar,Program Manager,Apps, Dialogs
        windows, windows_handles = self.get_windows(controls_handles=controls_handles)  # Apps
        active_window = self.get_active_window(windows=windows)  # Active Window
        active_window_handle = active_window.handle if active_window else None

        try:
            active_desktop = get_current_desktop()
            all_desktops = get_all_desktops()
        except RuntimeError:
            active_desktop = {
                "id": "00000000-0000-0000-0000-000000000000",
                "name": "Default Desktop",
            }
            all_desktops = [active_desktop]

        if active_window is not None and active_window in windows:
            windows.remove(active_window)

        logger.debug("Active window: %s", active_window or "No Active Window Found")
        logger.debug("Windows: %s", windows)

        # Preparing handles for Tree
        other_windows_handles = list(controls_handles - windows_handles)

        try:
            tree_state = self.tree.get_state(
                active_window_handle, other_windows_handles, use_dom=use_dom
            )
        except Exception:
            logger.exception("Failed to capture tree state")
            from windows_mcp.tree.views import TreeState

            tree_state = TreeState()

        if use_vision:
            if use_annotation:
                nodes = tree_state.interactive_nodes
                screenshot = self.get_annotated_screenshot(nodes=nodes)
            else:
                screenshot = self.get_screenshot()

            if scale != 1.0:
                screenshot = screenshot.resize(
                    (int(screenshot.width * scale), int(screenshot.height * scale)),
                    Image.LANCZOS,
                )

            if as_bytes:
                buffered = io.BytesIO()
                screenshot.save(buffered, format="PNG")
                screenshot = buffered.getvalue()
                buffered.close()
        else:
            screenshot = None

        desktop_state = DesktopState(
            active_window=active_window,
            windows=windows,
            active_desktop=active_desktop,
            all_desktops=all_desktops,
            screenshot=screenshot,
            tree_state=tree_state,
        )
        with self._state_lock:
            self.desktop_state = desktop_state
        # Log the time taken to capture the state
        end_time = time()
        logger.info("Desktop State capture took %.2f seconds", end_time - start_time)
        return desktop_state

    def get_window_status(self, control: uia.Control) -> Status:
        return self._window.get_window_status(control)

    def get_apps_from_start_menu(self) -> dict[str, str]:
        """Get installed apps with caching. Tries Get-StartApps first, falls back to shortcut scanning."""
        now = time()
        # Fast path: check cache without lock (atomic read on CPython)
        if self._app_cache is not None and (now - self._app_cache_time) < self._APP_CACHE_TTL:
            return self._app_cache

        with self._app_cache_lock:
            # Double-check after acquiring lock
            now = time()
            if self._app_cache is not None and (now - self._app_cache_time) < self._APP_CACHE_TTL:
                return self._app_cache

            command = "Get-StartApps | ConvertTo-Csv -NoTypeInformation"
            apps_info, status = self.execute_command(command)

            if status == 0 and apps_info and apps_info.strip():
                try:
                    reader = csv.DictReader(io.StringIO(apps_info.strip()))
                    apps = {
                        row.get("Name", "").lower(): row.get("AppID", "")
                        for row in reader
                        if row.get("Name") and row.get("AppID")
                    }
                    if apps:
                        self._app_cache = apps
                        self._app_cache_time = now
                        return apps
                except Exception as e:
                    logger.warning("Error parsing Get-StartApps output: %s", e)

            # Fallback: scan Start Menu shortcut folders (works on all Windows versions)
            logger.info("Get-StartApps unavailable, falling back to Start Menu folder scan")
            apps = self._get_apps_from_shortcuts()
            self._app_cache = apps
            self._app_cache_time = now
            return apps

    def _get_apps_from_shortcuts(self) -> dict[str, str]:
        """Scan Start Menu folders for .lnk shortcuts as a fallback for Get-StartApps."""
        import glob

        apps = {}
        start_menu_paths = [
            os.path.join(
                os.environ.get("PROGRAMDATA", r"C:\ProgramData"),
                r"Microsoft\Windows\Start Menu\Programs",
            ),
            os.path.join(
                os.environ.get("APPDATA", ""),
                r"Microsoft\Windows\Start Menu\Programs",
            ),
        ]
        for base_path in start_menu_paths:
            if not os.path.isdir(base_path):
                continue
            for lnk_path in glob.glob(os.path.join(base_path, "**", "*.lnk"), recursive=True):
                name = os.path.splitext(os.path.basename(lnk_path))[0].lower()
                if name and name not in apps:
                    apps[name] = lnk_path
        return apps

    def execute_command(self, command: str, timeout: int = 10) -> tuple[str, int]:
        return self._shell.execute(command, timeout)

    def is_window_browser(self, node: uia.Control):
        """Return True if the UIA control belongs to a browser process."""
        return self._window.is_window_browser(node)

    def resize_app(
        self,
        size: tuple[int, int] | None = None,
        loc: tuple[int, int] | None = None,
    ) -> tuple[str, int]:
        with self._state_lock:
            state = self.desktop_state
        if state is None:
            return "No desktop state available", 1
        active_window = state.active_window
        if active_window is None:
            return "No active window found", 1
        if active_window.status == Status.MINIMIZED:
            return f"{active_window.name} is minimized", 1
        elif active_window.status == Status.MAXIMIZED:
            return f"{active_window.name} is maximized", 1
        else:
            window_control = uia.ControlFromHandle(active_window.handle)
            if loc is None:
                x = window_control.BoundingRectangle.left
                y = window_control.BoundingRectangle.top
                loc = (x, y)
            if size is None:
                width = window_control.BoundingRectangle.width()
                height = window_control.BoundingRectangle.height()
                size = (width, height)
            x, y = loc
            width, height = size
            window_control.MoveWindow(x, y, width, height)
            return (f"{active_window.name} resized to {width}x{height} at {x},{y}.", 0)

    def is_app_running(self, name: str) -> bool:
        try:
            windows, _ = self._window.get_windows()
        except Exception:
            return False
        windows_dict = {window.name: window for window in windows if window.name}
        return process.extractOne(name, list(windows_dict.keys()), score_cutoff=60) is not None

    def app(
        self,
        mode: Literal["launch", "switch", "resize"],
        name: str | None = None,
        loc: tuple[int, int] | None = None,
        size: tuple[int, int] | None = None,
    ):
        if name is None and mode in ("launch", "switch"):
            return "Application name is required for launch/switch mode."

        match mode:
            case "launch":
                response, status, pid = self.launch_app(name)
                if status != 0:
                    return response

                # Smart wait using UIA Exists (avoids manual Python loops)
                launched = False
                if pid > 0:
                    if uia.WindowControl(ProcessId=pid).Exists(maxSearchSeconds=10):
                        launched = True

                if not launched:
                    # Fallback: Regex search for the window title
                    safe_name = re.escape(name)
                    if uia.WindowControl(RegexName=f"(?i).*{safe_name}.*").Exists(
                        maxSearchSeconds=10
                    ):
                        launched = True

                if launched:
                    return f"{name.title()} launched."
                return f"Launching {name.title()} sent, but window not detected yet."
            case "resize":
                response, _status = self.resize_app(size=size, loc=loc)
                return response
            case "switch":
                response, _status = self.switch_app(name)
                return response

    def launch_app(self, name: str) -> tuple[str, int, int]:
        apps_map = self.get_apps_from_start_menu()
        matched_app = process.extractOne(name, apps_map.keys(), score_cutoff=70)
        if matched_app is None:
            return (f"{name.title()} not found in start menu.", 1, 0)
        app_name, _ = matched_app
        appid = apps_map.get(app_name)
        if appid is None:
            return (f"{name.title()} not found in start menu.", 1, 0)

        pid = 0
        if os.path.exists(appid) or "\\" in appid:
            safe = self._ps_quote(appid)
            command = f"Start-Process {safe} -PassThru | Select-Object -ExpandProperty Id"
            response, status = self.execute_command(command)
            if status == 0 and response.strip().isdigit():
                pid = int(response.strip())
        else:
            if (
                not appid.replace("\\", "")
                .replace("_", "")
                .replace(".", "")
                .replace("-", "")
                .isalnum()
            ):
                return (f"Invalid app identifier: {appid}", 1, 0)
            safe = self._ps_quote(f"shell:AppsFolder\\{appid}")
            command = f"Start-Process {safe}"
            response, status = self.execute_command(command)

        return response, status, pid

    def switch_app(self, name: str):
        try:
            # Refresh state if desktop_state is None or has no windows
            with self._state_lock:
                state = self.desktop_state
            if state is None or not state.windows:
                self.get_state()
                with self._state_lock:
                    state = self.desktop_state
            if state is None:
                return ("Failed to get desktop state. Please try again.", 1)

            window_list = [w for w in [state.active_window] + state.windows if w is not None]
            if not window_list:
                return ("No windows found on the desktop.", 1)

            windows = {window.name: window for window in window_list}
            matched_window: tuple[str, float] | None = process.extractOne(
                name, list(windows.keys()), score_cutoff=70
            )
            if matched_window is None:
                return (f"Application {name.title()} not found.", 1)
            window_name, _ = matched_window
            window = windows.get(window_name)
            if window is None:
                return (f"Application {name.title()} not found.", 1)
            target_handle = window.handle

            if uia.IsIconic(target_handle):
                uia.ShowWindow(target_handle, win32con.SW_RESTORE)
                content = f"{window_name.title()} restored from Minimized state."
            else:
                self._window.bring_window_to_top(target_handle)
                content = f"Switched to {window_name.title()} window."
            return content, 0
        except Exception as e:
            return (f"Error switching app: {str(e)}", 1)

    def bring_window_to_top(self, target_handle: int):
        self._window.bring_window_to_top(target_handle)

    def get_element_handle_from_label(self, label: int) -> uia.Control:
        with self._state_lock:
            state = self.desktop_state
        if state is None or state.tree_state is None:
            raise ValueError("No desktop state available. Call get_state() first.")
        tree_state = state.tree_state
        if label < 0 or label >= len(tree_state.interactive_nodes):
            raise ValueError(
                f"Label {label} out of range (0-{len(tree_state.interactive_nodes) - 1}). "
                "The UI may have changed since last snapshot."
            )
        element_node = tree_state.interactive_nodes[label]
        xpath = element_node.xpath
        element_handle = self.get_element_from_xpath(xpath)
        return element_handle

    def get_coordinates_from_label(self, label: int) -> tuple[int, int]:
        element_handle = self.get_element_handle_from_label(label)
        bounding_rectangle = element_handle.BoundingRectangle
        return bounding_rectangle.xcenter(), bounding_rectangle.ycenter()

    def click(self, loc: tuple[int, int], button: str = "left", clicks: int = 1):
        return self._input.click(loc, button, clicks)

    def type(
        self,
        loc: tuple[int, int],
        text: str,
        caret_position: Literal["start", "idle", "end"] = "idle",
        clear: bool | str = False,
        press_enter: bool | str = False,
    ):
        return self._input.type(loc, text, caret_position, clear, press_enter)

    def scroll(
        self,
        loc: tuple[int, int] | None = None,
        type: Literal["horizontal", "vertical"] = "vertical",
        direction: Literal["up", "down", "left", "right"] = "down",
        wheel_times: int = 1,
    ) -> str | None:
        return self._input.scroll(loc, type, direction, wheel_times)

    def drag(self, loc: tuple[int, int]):
        return self._input.drag(loc)

    def move(self, loc: tuple[int, int]):
        return self._input.move(loc)

    def shortcut(self, shortcut: str):
        return self._input.shortcut(shortcut)

    def multi_select(
        self, press_ctrl: bool | str = False, locs: list[tuple[int, int]] | None = None
    ):
        return self._input.multi_select(press_ctrl, locs)

    def multi_edit(self, locs: list[tuple[int, int, str]]):
        return self._input.multi_edit(locs)

    # --- Scraper facade (delegates to ScraperService) ---

    def scrape(self, url: str) -> str:
        return self._scraper.scrape(url)

    def get_window_from_element(self, element: uia.Control) -> Window | None:
        return self._window.get_window_from_element(element)

    def is_overlay_window(self, element: uia.Control) -> bool:
        return self._window.is_overlay_window(element)

    def get_controls_handles(self, optimized: bool = False):
        return self._window.get_controls_handles(optimized)

    def get_active_window(self, windows: list[Window] | None = None) -> Window | None:
        return self._window.get_active_window(windows)

    def get_foreground_window(self) -> uia.Control:
        return self._window.get_foreground_window()

    def get_window_from_element_handle(self, element_handle: int) -> uia.Control:
        return self._window.get_window_from_element_handle(element_handle)

    def get_windows(
        self, controls_handles: set[int] | None = None
    ) -> tuple[list[Window], set[int]]:
        return self._window.get_windows(controls_handles)

    _MAX_PARENT_DEPTH = 64

    def get_xpath_from_element(self, element: uia.Control):
        current = element
        if current is None:
            return ""
        path_parts = []
        for _ in range(self._MAX_PARENT_DEPTH):
            if current is None:
                break
            parent = current.GetParentControl()
            if parent is None:
                # we are at the root node
                path_parts.append(f"{current.ControlTypeName}")
                break
            children = parent.GetChildren()
            same_type_children = [
                "-".join(map(str, child.GetRuntimeId()))
                for child in children
                if child.ControlType == current.ControlType
            ]
            current_id = "-".join(map(str, current.GetRuntimeId()))
            try:
                index = same_type_children.index(current_id)
            except ValueError:
                index = 0
            if same_type_children:
                path_parts.append(f"{current.ControlTypeName}[{index + 1}]")
            else:
                path_parts.append(f"{current.ControlTypeName}")
            current = parent
        path_parts.reverse()
        xpath = "/".join(path_parts)
        return xpath

    def get_element_from_xpath(self, xpath: str) -> uia.Control:
        pattern = re.compile(r"(\w+)(?:\[(\d+)\])?")
        parts = xpath.split("/")
        root = uia.GetRootControl()
        element = root
        for part in parts[1:]:
            match = pattern.fullmatch(part)
            if match is None:
                continue
            control_type, index = match.groups()
            index = int(index) if index else None
            children = element.GetChildren()
            same_type_children = [c for c in children if c.ControlTypeName == control_type]
            if not same_type_children:
                raise ValueError(
                    f"XPath resolution failed: no children of type '{control_type}' found. "
                    "The UI may have changed since last snapshot."
                )
            if index is not None:
                if index < 1 or index - 1 >= len(same_type_children):
                    raise ValueError(
                        f"XPath resolution failed: index {index} out of range for "
                        f"{len(same_type_children)} children of type '{control_type}'. "
                        "The UI may have changed since last snapshot."
                    )
                element = same_type_children[index - 1]
            else:
                element = same_type_children[0]
        return element

    def get_screen_size(self) -> Size:
        return self._screen.get_screen_size()

    def get_screenshot(self) -> Image.Image:
        return self._screen.get_screenshot()

    def get_annotated_screenshot(self, nodes: list[TreeElementNode]) -> Image.Image:
        return self._screen.get_annotated_screenshot(nodes)

    def send_notification(self, title: str, message: str) -> str:
        from xml.sax.saxutils import escape as xml_escape

        safe_title = xml_escape(title, {'"': "&quot;", "'": "&apos;"})
        safe_message = xml_escape(message, {'"': "&quot;", "'": "&apos;"})

        ps_script = (
            "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null\n"
            "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null\n"
            f"$notifTitle = {self._ps_quote(safe_title)}\n"
            f"$notifMessage = {self._ps_quote(safe_message)}\n"
            '$template = @"\n'
            "<toast>\n"
            "    <visual>\n"
            '        <binding template="ToastGeneric">\n'
            "            <text>$notifTitle</text>\n"
            "            <text>$notifMessage</text>\n"
            "        </binding>\n"
            "    </visual>\n"
            "</toast>\n"
            '"@\n'
            "$xml = New-Object Windows.Data.Xml.Dom.XmlDocument\n"
            "$xml.LoadXml($template)\n"
            '$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Windows MCP")\n'
            "$toast = New-Object Windows.UI.Notifications.ToastNotification $xml\n"
            "$notifier.Show($toast)"
        )
        response, status = self.execute_command(ps_script)
        if status == 0:
            return f'Notification sent: "{title}" - {message}'
        else:
            return f"Notification may have been sent. PowerShell output: {response[:200]}"

    def list_processes(
        self,
        name: str | None = None,
        sort_by: Literal["memory", "cpu", "name"] = "memory",
        limit: int = 20,
    ) -> str:
        import psutil
        from tabulate import tabulate

        procs = []
        for p in psutil.process_iter(["pid", "name", "cpu_percent", "memory_info"]):
            try:
                info = p.info
                mem_mb = info["memory_info"].rss / (1024 * 1024) if info["memory_info"] else 0
                procs.append(
                    {
                        "pid": info["pid"],
                        "name": info["name"] or "Unknown",
                        "cpu": info["cpu_percent"] or 0,
                        "mem_mb": round(mem_mb, 1),
                    }
                )
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        if name:
            from thefuzz import fuzz

            procs = [p for p in procs if fuzz.partial_ratio(name.lower(), p["name"].lower()) > 60]
        sort_key = {
            "memory": lambda x: x["mem_mb"],
            "cpu": lambda x: x["cpu"],
            "name": lambda x: x["name"].lower(),
        }
        procs.sort(key=sort_key.get(sort_by, sort_key["memory"]), reverse=(sort_by != "name"))
        procs = procs[: max(1, limit)]
        if not procs:
            return f"No processes found{f' matching {name}' if name else ''}."
        table = tabulate(
            [[p["pid"], p["name"], f"{p['cpu']:.1f}%", f"{p['mem_mb']:.1f} MB"] for p in procs],
            headers=["PID", "Name", "CPU%", "Memory"],
            tablefmt="simple",
        )
        return f"Processes ({len(procs)} shown):\n{table}"

    def kill_process(
        self, name: str | None = None, pid: int | None = None, force: bool = False
    ) -> str:
        import psutil

        if pid is None and name is None:
            return "Error: Provide either pid or name parameter for kill mode."
        killed = []
        if pid is not None:
            try:
                p = psutil.Process(pid)
                pname = p.name()
                if force:
                    p.kill()
                else:
                    p.terminate()
                killed.append(f"{pname} (PID {pid})")
            except psutil.NoSuchProcess:
                return f"No process with PID {pid} found."
            except psutil.AccessDenied:
                return f"Access denied to kill PID {pid}. Try running as administrator."
        else:
            for p in psutil.process_iter(["pid", "name"]):
                try:
                    if p.info["name"] and p.info["name"].lower() == name.lower():
                        if force:
                            p.kill()
                        else:
                            p.terminate()
                        killed.append(f"{p.info['name']} (PID {p.info['pid']})")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        if not killed:
            return f'No process matching "{name}" found or access denied.'
        return f"{'Force killed' if force else 'Terminated'}: {', '.join(killed)}"

    def lock_screen(self) -> str:
        try:
            ctypes.windll.user32.LockWorkStation()
        except Exception as e:
            return f"Failed to lock screen: {e}"
        return "Screen locked."

    def get_system_info(self) -> str:
        import platform
        from datetime import datetime, timedelta
        from textwrap import dedent

        import psutil

        from windows_mcp.native import native_system_info

        # Try Rust fast-path for CPU/memory/disk (avoids 1s blocking cpu_percent)
        native_info = native_system_info()

        if native_info is not None and "cpu_count" in native_info:
            os_str = f"{platform.system()} {platform.release()} ({platform.version()})"
            cpu_count = native_info.get("cpu_count", 0)
            cpu_usages = native_info.get("cpu_usage_percent", [])
            cpu_pct = round(sum(cpu_usages) / len(cpu_usages), 1) if cpu_usages else 0.0
            mem_total = native_info.get("total_memory_bytes", 0)
            mem_used = native_info.get("used_memory_bytes", 0)
            mem_pct = round(mem_used / mem_total * 100, 1) if mem_total else 0.0

            # Find C: disk from Rust data
            disk_pct = disk_used_gb = disk_total_gb = 0.0
            for d in native_info.get("disks", []):
                mount = d.get("mount_point", "")
                total = d.get("total_bytes", 0)
                available = d.get("available_bytes", 0)
                if mount.upper().startswith("C:"):
                    disk_total_gb = round(total / 1024**3, 1)
                    disk_used = total - available
                    disk_used_gb = round(disk_used / 1024**3, 1)
                    disk_pct = round(disk_used / total * 100, 1) if total else 0.0
                    break
        else:
            os_str = f"{platform.system()} {platform.release()} ({platform.version()})"
            cpu_pct = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count()
            mem = psutil.virtual_memory()
            mem_pct = mem.percent
            mem_used = mem.used
            mem_total = mem.total
            disk = psutil.disk_usage("C:\\")
            disk_pct = disk.percent
            disk_used_gb = round(disk.used / 1024**3, 1)
            disk_total_gb = round(disk.total / 1024**3, 1)

        # Network and uptime always from psutil (Rust doesn't have these yet)
        boot = datetime.fromtimestamp(psutil.boot_time())
        uptime = datetime.now() - boot
        uptime_str = str(timedelta(seconds=int(uptime.total_seconds())))
        net = psutil.net_io_counters()

        return dedent(f"""System Information:
  OS: {os_str}
  Machine: {platform.machine()}

  CPU: {cpu_pct}% ({cpu_count} cores)
  Memory: {mem_pct}% used ({round(mem_used / 1024**3, 1)} / {round(mem_total / 1024**3, 1)} GB)
  Disk C: {disk_pct}% used ({disk_used_gb} / {disk_total_gb} GB)

  Network: ↑ {round(net.bytes_sent / 1024**2, 1)} MB sent, ↓ {round(net.bytes_recv / 1024**2, 1)} MB received
  Uptime: {uptime_str} (booted {boot.strftime("%Y-%m-%d %H:%M")})""")

    # --- Registry facade (delegates to RegistryService) ---

    def _parse_reg_path(self, path: str) -> tuple:
        return self._registry._parse_reg_path(path)

    def registry_get(self, path: str, name: str) -> str:
        return self._registry.registry_get(path, name)

    def registry_set(self, path: str, name: str, value: str, reg_type: str = "String") -> str:
        return self._registry.registry_set(path, name, value, reg_type)

    def registry_delete(self, path: str, name: str | None = None) -> str:
        return self._registry.registry_delete(path, name)

    def registry_list(self, path: str) -> str:
        return self._registry.registry_list(path)

    @contextmanager
    def auto_minimize(self):
        with self._window.auto_minimize():
            yield
