"""Comprehensive unit tests for WindowService.

All Win32 API calls (win32gui, win32process, win32con, ctypes), UIA module
symbols, psutil.Process, and the vdm helper are mocked so no live desktop
session is required.

Coverage targets: 100% of windows_mcp/window/service.py
"""

import logging
import threading
from unittest.mock import MagicMock, call, patch

import pytest

from windows_mcp.desktop.views import BoundingBox, Browser, Status, Window
from windows_mcp.window.service import _MAX_PARENT_DEPTH, _PROCESS_CACHE_MAX, WindowService

# ---------------------------------------------------------------------------
# Patch target constants
# ---------------------------------------------------------------------------

_UIA = "windows_mcp.window.service.uia"
_WIN32GUI = "windows_mcp.window.service.win32gui"
_WIN32PROCESS = "windows_mcp.window.service.win32process"
_WIN32CON = "windows_mcp.window.service.win32con"
_CTYPES = "windows_mcp.window.service.ctypes"
_PSUTIL_PROCESS = "windows_mcp.window.service.Process"
_IS_WINDOW_ON_CURRENT_DESKTOP = "windows_mcp.window.service.is_window_on_current_desktop"
_NATIVE_LIST_WINDOWS = "windows_mcp.window.service.native_list_windows"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_control(
    name: str = "TestWindow",
    handle: int = 1001,
    class_name: str = "TestClass",
    process_id: int = 9999,
    children: list | None = None,
) -> MagicMock:
    """Build a minimal MagicMock that behaves like a uia.Control."""
    ctrl = MagicMock()
    ctrl.Name = name
    ctrl.NativeWindowHandle = handle
    ctrl.ClassName = class_name
    ctrl.ProcessId = process_id
    ctrl.GetChildren.return_value = children if children is not None else []
    return ctrl


def _make_bounding_rect(left=0, top=0, right=100, bottom=100, empty=False):
    """Return a mock bounding rectangle."""
    rect = MagicMock()
    rect.left = left
    rect.top = top
    rect.right = right
    rect.bottom = bottom
    rect.width.return_value = right - left
    rect.height.return_value = bottom - top
    rect.isempty.return_value = empty
    return rect


def _make_window(
    name: str = "TestWindow",
    handle: int = 1001,
    process_id: int = 9999,
    status: Status = Status.NORMAL,
    is_browser: bool = False,
) -> Window:
    """Build a Window dataclass with sane defaults."""
    return Window(
        name=name,
        handle=handle,
        process_id=process_id,
        status=status,
        is_browser=is_browser,
        depth=0,
        bounding_box=BoundingBox(left=0, top=0, right=800, bottom=600, width=800, height=600),
    )


@pytest.fixture()
def svc() -> WindowService:
    return WindowService()


# ===========================================================================
# get_window_status
# ===========================================================================


class TestGetWindowStatus:
    """Tests for WindowService.get_window_status()."""

    def test_returns_minimized_when_iconic(self, svc):
        ctrl = _make_control(handle=100)
        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = True
            result = svc.get_window_status(ctrl)
        assert result == Status.MINIMIZED
        mock_uia.IsIconic.assert_called_once_with(100)

    def test_returns_maximized_when_zoomed_not_iconic(self, svc):
        ctrl = _make_control(handle=200)
        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = False
            mock_uia.IsZoomed.return_value = True
            result = svc.get_window_status(ctrl)
        assert result == Status.MAXIMIZED

    def test_returns_normal_when_visible_not_iconic_not_zoomed(self, svc):
        ctrl = _make_control(handle=300)
        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = False
            mock_uia.IsZoomed.return_value = False
            mock_uia.IsWindowVisible.return_value = True
            result = svc.get_window_status(ctrl)
        assert result == Status.NORMAL

    def test_returns_hidden_when_not_visible_not_iconic_not_zoomed(self, svc):
        ctrl = _make_control(handle=400)
        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = False
            mock_uia.IsZoomed.return_value = False
            mock_uia.IsWindowVisible.return_value = False
            result = svc.get_window_status(ctrl)
        assert result == Status.HIDDEN

    def test_iconic_takes_priority_over_zoomed(self, svc):
        """If IsIconic is True, MINIMIZED is returned regardless of IsZoomed."""
        ctrl = _make_control(handle=500)
        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = True
            mock_uia.IsZoomed.return_value = True
            result = svc.get_window_status(ctrl)
        assert result == Status.MINIMIZED
        mock_uia.IsZoomed.assert_not_called()


# ===========================================================================
# is_overlay_window
# ===========================================================================


class TestIsOverlayWindow:
    """Tests for WindowService.is_overlay_window()."""

    def test_no_children_returns_true(self, svc):
        ctrl = _make_control(children=[])
        assert svc.is_overlay_window(ctrl) is True

    def test_overlay_in_name_returns_true(self, svc):
        ctrl = _make_control(name="NVIDIA Overlay", children=[MagicMock()])
        assert svc.is_overlay_window(ctrl) is True

    def test_overlay_substring_case_sensitive_true(self, svc):
        """'Overlay' must appear in Name (capitalised per implementation)."""
        ctrl = _make_control(name="Steam Overlay Tool", children=[MagicMock()])
        assert svc.is_overlay_window(ctrl) is True

    def test_not_overlay_has_children_name_no_overlay(self, svc):
        ctrl = _make_control(name="Notepad", children=[MagicMock(), MagicMock()])
        assert svc.is_overlay_window(ctrl) is False

    def test_name_none_treated_as_empty_string(self, svc):
        """None Name must not raise; fallback to '' strip keeps the overlay logic intact."""
        ctrl = MagicMock()
        ctrl.Name = None
        ctrl.GetChildren.return_value = [MagicMock()]
        assert svc.is_overlay_window(ctrl) is False

    def test_name_whitespace_only_not_overlay(self, svc):
        ctrl = _make_control(name="   ", children=[MagicMock()])
        assert svc.is_overlay_window(ctrl) is False

    def test_overlay_in_name_with_no_children_both_conditions_true(self, svc):
        """Both flags True -- still True."""
        ctrl = _make_control(name="Game Overlay", children=[])
        assert svc.is_overlay_window(ctrl) is True

    def test_lowercase_overlay_does_not_match(self, svc):
        """'overlay' (lowercase) must NOT match; the check uses 'Overlay'."""
        ctrl = _make_control(name="some overlay helper", children=[MagicMock()])
        assert svc.is_overlay_window(ctrl) is False


# ===========================================================================
# is_window_browser
# ===========================================================================


class TestIsWindowBrowser:
    """Tests for WindowService.is_window_browser()."""

    def test_chrome_process_returns_true(self, svc):
        ctrl = _make_control(process_id=100)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "chrome.exe"
            result = svc.is_window_browser(ctrl)
        assert result is True

    def test_msedge_process_returns_true(self, svc):
        ctrl = _make_control(process_id=200)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "msedge.exe"
            result = svc.is_window_browser(ctrl)
        assert result is True

    def test_firefox_process_returns_true(self, svc):
        ctrl = _make_control(process_id=300)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "firefox.exe"
            result = svc.is_window_browser(ctrl)
        assert result is True

    def test_notepad_process_returns_false(self, svc):
        ctrl = _make_control(process_id=400)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "notepad.exe"
            result = svc.is_window_browser(ctrl)
        assert result is False

    def test_cache_hit_skips_process_creation(self, svc):
        """Second call for same PID must use cache, not create Process again."""
        ctrl = _make_control(process_id=500)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "chrome.exe"
            svc.is_window_browser(ctrl)
            svc.is_window_browser(ctrl)
        assert MockProcess.call_count == 1

    def test_cache_miss_calls_process(self, svc):
        """Different PIDs must each call Process()."""
        ctrl_a = _make_control(process_id=501)
        ctrl_b = _make_control(process_id=502)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "notepad.exe"
            svc.is_window_browser(ctrl_a)
            svc.is_window_browser(ctrl_b)
        assert MockProcess.call_count == 2

    def test_cache_eviction_at_max(self, svc):
        """When cache reaches _PROCESS_CACHE_MAX, it is cleared before inserting."""
        # Pre-fill the cache to exactly _PROCESS_CACHE_MAX entries
        for pid in range(_PROCESS_CACHE_MAX):
            svc._process_name_cache[pid] = "notepad.exe"
        assert len(svc._process_name_cache) == _PROCESS_CACHE_MAX

        # The next call should clear the cache and re-add a single entry
        ctrl = _make_control(process_id=_PROCESS_CACHE_MAX + 1)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "chrome.exe"
            result = svc.is_window_browser(ctrl)

        assert result is True
        assert len(svc._process_name_cache) == 1
        assert svc._process_name_cache[_PROCESS_CACHE_MAX + 1] == "chrome.exe"

    def test_cache_eviction_boundary_below_max_no_clear(self, svc):
        """Cache size < _PROCESS_CACHE_MAX must NOT be cleared."""
        for pid in range(_PROCESS_CACHE_MAX - 1):
            svc._process_name_cache[pid] = "notepad.exe"

        ctrl = _make_control(process_id=_PROCESS_CACHE_MAX)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "chrome.exe"
            svc.is_window_browser(ctrl)

        # Old entries still present
        assert len(svc._process_name_cache) == _PROCESS_CACHE_MAX

    def test_process_raises_exception_returns_false(self, svc):
        """psutil.Process raising any exception must result in False (no crash)."""
        ctrl = _make_control(process_id=600)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.side_effect = Exception("No such process")
            result = svc.is_window_browser(ctrl)
        assert result is False

    def test_process_name_raises_exception_returns_false(self, svc):
        """Process().name() raising must result in False."""
        ctrl = _make_control(process_id=700)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.side_effect = OSError("access denied")
            result = svc.is_window_browser(ctrl)
        assert result is False

    def test_cached_value_is_used_for_browser_check(self, svc):
        """A pre-cached process name is used for browser detection."""
        svc._process_name_cache[800] = "chrome.exe"
        ctrl = _make_control(process_id=800)
        with patch(_PSUTIL_PROCESS) as MockProcess:
            result = svc.is_window_browser(ctrl)
        assert result is True
        MockProcess.assert_not_called()

    def test_thread_safety_no_race_condition(self, svc):
        """Multiple threads calling is_window_browser must not corrupt the cache."""
        results = []
        errors = []

        def worker(pid):
            ctrl = _make_control(process_id=pid)
            with patch(_PSUTIL_PROCESS) as MockProcess:
                MockProcess.return_value.name.return_value = "notepad.exe"
                try:
                    results.append(svc.is_window_browser(ctrl))
                except Exception as exc:
                    errors.append(exc)

        threads = [threading.Thread(target=worker, args=(i,)) for i in range(20)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert not errors


# ===========================================================================
# get_controls_handles
# ===========================================================================


class TestGetControlsHandles:
    """Tests for WindowService.get_controls_handles()."""

    def _make_enum_windows(self, visible_handles: list[int]):
        """Return a side_effect function that invokes callback for each handle."""

        def _enum(callback, param):
            for hwnd in visible_handles:
                callback(hwnd, param)

        return _enum

    def test_visible_handles_on_current_desktop_are_included(self, svc):
        handles = [1001, 1002, 1003]
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsWindowVisible.return_value = True
            mock_gui.EnumWindows.side_effect = self._make_enum_windows(handles)
            mock_gui.FindWindow.return_value = 0

            result = svc.get_controls_handles()

        assert handles[0] in result
        assert handles[1] in result
        assert handles[2] in result

    def test_non_visible_handles_excluded(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsWindowVisible.return_value = False
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([9001])
            mock_gui.FindWindow.return_value = 0

            result = svc.get_controls_handles()

        assert 9001 not in result

    def test_handles_not_on_current_desktop_excluded(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=False),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsWindowVisible.return_value = True
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([9002])
            mock_gui.FindWindow.return_value = 0

            result = svc.get_controls_handles()

        assert 9002 not in result

    def test_progman_handle_added_when_found(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsWindowVisible.return_value = True
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([])

            def find_window_side_effect(class_name, _):
                if class_name == "Progman":
                    return 5001
                return 0

            mock_gui.FindWindow.side_effect = find_window_side_effect

            result = svc.get_controls_handles()

        assert 5001 in result

    def test_shell_traywnd_added_when_found(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([])

            def find_window_side_effect(class_name, _):
                if class_name == "Shell_TrayWnd":
                    return 5002
                return 0

            mock_gui.FindWindow.side_effect = find_window_side_effect

            result = svc.get_controls_handles()

        assert 5002 in result

    def test_secondary_traywnd_added_when_found(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([])

            def find_window_side_effect(class_name, _):
                if class_name == "Shell_SecondaryTrayWnd":
                    return 5003
                return 0

            mock_gui.FindWindow.side_effect = find_window_side_effect

            result = svc.get_controls_handles()

        assert 5003 in result

    def test_exception_in_callback_is_swallowed(self, svc):
        """Exceptions in the EnumWindows callback must not propagate."""

        def _enum_with_error(callback, param):
            callback(9999, param)  # this will raise via IsWindow

        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, side_effect=RuntimeError("vdm error")),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsWindowVisible.return_value = True
            mock_gui.EnumWindows.side_effect = _enum_with_error
            mock_gui.FindWindow.return_value = 0

            # Must not raise
            result = svc.get_controls_handles()
        assert isinstance(result, set)

    def test_all_find_window_return_zero_not_added(self, svc):
        """When FindWindow returns 0 (falsy), handles are NOT added."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
        ):
            mock_gui.EnumWindows.side_effect = self._make_enum_windows([])
            mock_gui.FindWindow.return_value = 0

            result = svc.get_controls_handles()

        assert 0 not in result


# ===========================================================================
# get_window_from_element_handle
# ===========================================================================


class TestGetWindowFromElementHandle:
    """Tests for WindowService.get_window_from_element_handle()."""

    def test_returns_control_when_parent_is_root(self, svc):
        """Stops walking when parent.NativeWindowHandle == root_handle."""
        root = _make_control(handle=0)
        child = _make_control(handle=100)
        parent = _make_control(handle=0)  # same as root

        child.GetParentControl.return_value = parent

        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.return_value = child
            mock_uia.GetRootControl.return_value = root

            result = svc.get_window_from_element_handle(100)

        assert result is child

    def test_returns_control_when_parent_is_none(self, svc):
        """Stops walking when GetParentControl returns None."""
        root = _make_control(handle=0)
        child = _make_control(handle=200)
        child.GetParentControl.return_value = None

        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.return_value = child
            mock_uia.GetRootControl.return_value = root

            result = svc.get_window_from_element_handle(200)

        assert result is child

    def test_walks_parent_chain_until_root(self, svc):
        """Walks multiple levels before hitting root."""
        root = _make_control(handle=1)
        grandchild = _make_control(handle=10)
        child = _make_control(handle=20)
        top = _make_control(handle=30)
        # top's parent is at root level
        root_level = _make_control(handle=1)

        grandchild.GetParentControl.return_value = child
        child.GetParentControl.return_value = top
        top.GetParentControl.return_value = root_level

        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.return_value = grandchild
            mock_uia.GetRootControl.return_value = root

            result = svc.get_window_from_element_handle(10)

        assert result is top

    def test_depth_limit_exceeded_returns_current_and_logs_warning(self, svc, caplog):
        """When parent chain exceeds _MAX_PARENT_DEPTH, return current and warn."""
        root = _make_control(handle=0)

        # Build a parent that always returns a different non-root control
        # so the loop never exits early
        def make_non_root_parent(handle):
            ctrl = _make_control(handle=handle)
            return ctrl

        # Each GetParentControl returns a new non-root control
        controls = [make_non_root_parent(i + 1000) for i in range(_MAX_PARENT_DEPTH + 5)]
        for i in range(len(controls) - 1):
            controls[i].GetParentControl.return_value = controls[i + 1]
        controls[-1].GetParentControl.return_value = controls[-1]  # loop on last

        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.return_value = controls[0]
            mock_uia.GetRootControl.return_value = root

            with caplog.at_level(logging.WARNING, logger="windows_mcp.window.service"):
                result = svc.get_window_from_element_handle(999)

        # Should return something (the current after exhausting depth)
        assert result is not None
        # Warning must be logged
        assert str(_MAX_PARENT_DEPTH) in caplog.text or "depth" in caplog.text.lower()


# ===========================================================================
# get_foreground_window
# ===========================================================================


class TestGetForegroundWindow:
    """Tests for WindowService.get_foreground_window()."""

    def test_calls_get_foreground_window_and_delegates(self, svc):
        ctrl = _make_control(handle=777)
        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_from_element_handle", return_value=ctrl) as mock_gwfeh,
        ):
            mock_uia.GetForegroundWindow.return_value = 777
            result = svc.get_foreground_window()
        mock_uia.GetForegroundWindow.assert_called_once()
        mock_gwfeh.assert_called_once_with(777)
        assert result is ctrl


# ===========================================================================
# get_active_window
# ===========================================================================


class TestGetActiveWindow:
    """Tests for WindowService.get_active_window()."""

    def test_returns_matching_window_from_list(self, svc):
        win = _make_window(handle=100)
        active_ctrl = _make_control(handle=100, class_name="SomeApp")

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows", return_value=([win], {100})),
        ):
            result = svc.get_active_window()

        assert result is win

    def test_returns_none_when_foreground_is_progman(self, svc):
        active_ctrl = _make_control(handle=200, class_name="Progman")

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows", return_value=([], set())),
        ):
            result = svc.get_active_window()

        assert result is None

    def test_returns_none_when_windows_is_empty_and_no_match(self, svc):
        active_ctrl = _make_control(handle=300, class_name="Calculator")

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows", return_value=([], set())),
        ):
            result = svc.get_active_window()

        # falls back to constructing a Window from active_ctrl
        assert result is not None
        assert result.handle == 300

    def test_constructs_fallback_window_when_not_in_list(self, svc):
        active_ctrl = _make_control(handle=400, name="Fallback App", class_name="FallbackClass")
        active_ctrl.BoundingRectangle = _make_bounding_rect(0, 0, 800, 600)
        active_ctrl.ProcessId = 1234

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows", return_value=([], set())),
            patch.object(svc, "is_window_browser", return_value=False),
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
        ):
            result = svc.get_active_window()

        assert result is not None
        assert result.name == "Fallback App"
        assert result.handle == 400
        assert result.process_id == 1234
        assert result.status == Status.NORMAL
        assert result.is_browser is False

    def test_uses_provided_windows_list_without_calling_get_windows(self, svc):
        win = _make_window(handle=500)
        active_ctrl = _make_control(handle=500, class_name="SomeApp")

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows") as mock_gw,
        ):
            result = svc.get_active_window(windows=[win])

        mock_gw.assert_not_called()
        assert result is win

    def test_returns_none_on_exception(self, svc, caplog):
        with (
            patch.object(svc, "get_foreground_window", side_effect=RuntimeError("COM failure")),
            patch.object(svc, "get_windows", return_value=([], set())),
            caplog.at_level(logging.ERROR, logger="windows_mcp.window.service"),
        ):
            result = svc.get_active_window()

        assert result is None
        assert "get_active_window" in caplog.text

    def test_empty_windows_list_passed_in_returns_fallback(self, svc):
        """When windows=[] is passed, get_windows is NOT called; fallback Window created."""
        active_ctrl = _make_control(handle=600, class_name="App")
        active_ctrl.BoundingRectangle = _make_bounding_rect(10, 20, 110, 120)
        active_ctrl.ProcessId = 5678

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows") as mock_gw,
            patch.object(svc, "is_window_browser", return_value=True),
            patch.object(svc, "get_window_status", return_value=Status.MAXIMIZED),
        ):
            result = svc.get_active_window(windows=[])

        mock_gw.assert_not_called()
        assert result.is_browser is True
        assert result.status == Status.MAXIMIZED

    def test_skips_non_matching_windows_before_match(self, svc):
        """The continue branch is taken for non-matching handles before the match."""
        non_match = _make_window(handle=701)
        matching = _make_window(handle=702)
        active_ctrl = _make_control(handle=702, class_name="RealApp")

        with (
            patch.object(svc, "get_foreground_window", return_value=active_ctrl),
            patch.object(svc, "get_windows", return_value=([non_match, matching], {701, 702})),
        ):
            result = svc.get_active_window()

        assert result is matching


# ===========================================================================
# get_windows
# ===========================================================================


class TestGetWindows:
    """Tests for WindowService.get_windows() -- UIA fallback path."""

    @pytest.fixture(autouse=True)
    def _no_native(self):
        """Disable Rust fast-path so tests exercise the UIA fallback."""
        with patch(_NATIVE_LIST_WINDOWS, return_value=None):
            yield

    def _make_window_control_child(
        self,
        handle: int = 1000,
        name: str = "App",
        can_min: bool = True,
        can_max: bool = True,
        status: Status = Status.NORMAL,
        bounding_rect=None,
        process_id: int = 1234,
        is_window_control: bool = True,
    ):
        """Build a mock that isinstance checks pass for uia.WindowControl."""
        from windows_mcp import uia as uia_module

        if bounding_rect is None:
            bounding_rect = _make_bounding_rect()

        ctrl = MagicMock(spec=uia_module.WindowControl)
        ctrl.Name = name
        ctrl.NativeWindowHandle = handle
        ctrl.ProcessId = process_id
        ctrl.BoundingRectangle = bounding_rect
        ctrl.GetChildren.return_value = [MagicMock()]  # not an overlay

        window_pattern = MagicMock()
        window_pattern.CanMinimize = can_min
        window_pattern.CanMaximize = can_max
        ctrl.GetPattern.return_value = window_pattern

        return ctrl

    def test_single_valid_window_returned(self, svc):
        ctrl = self._make_window_control_child(handle=1000)

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_controls_handles", return_value={1000}),
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={1000})

        assert len(windows) == 1
        assert windows[0].handle == 1000
        assert 1000 in handles

    def test_overlay_window_filtered_out(self, svc):
        ctrl = self._make_window_control_child(handle=2000)
        # Make it look like an overlay: no children
        ctrl.GetChildren.return_value = []

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={2000})

        assert len(windows) == 0

    def test_control_from_handle_exception_continues(self, svc):
        """Exception in ControlFromHandle must be caught and iteration continues."""
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.side_effect = Exception("COM error")
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={3000})

        assert windows == []
        assert handles == set()

    def test_non_window_non_pane_control_skipped(self, svc):
        """A control that is neither WindowControl nor PaneControl is skipped."""
        from windows_mcp import uia as uia_module

        # Use a MagicMock without WindowControl/PaneControl spec
        ctrl = MagicMock()
        ctrl.GetChildren.return_value = [MagicMock()]  # not overlay
        ctrl.Name = "GenericControl"
        ctrl.NativeWindowHandle = 4000

        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromHandle.return_value = ctrl
            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={4000})

        assert windows == []

    def test_window_pattern_none_skipped(self, svc):
        """Window with no WindowPattern is skipped (cannot min/max query)."""
        ctrl = self._make_window_control_child(handle=5000)
        ctrl.GetPattern.return_value = None  # no pattern

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={5000})

        assert windows == []

    def test_cannot_minimize_or_maximize_skipped(self, svc):
        """Windows that cannot be minimized/maximized are excluded."""
        ctrl = self._make_window_control_child(handle=6000, can_min=False, can_max=False)

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={6000})

        assert windows == []

    def test_empty_bounding_rect_not_minimized_filtered(self, svc):
        """Empty bounding rect on a non-minimized window is filtered out."""
        empty_rect = _make_bounding_rect(empty=True)
        ctrl = self._make_window_control_child(handle=7000, bounding_rect=empty_rect)

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),  # NOT minimized
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={7000})

        assert windows == []

    def test_empty_bounding_rect_minimized_included(self, svc):
        """Minimized windows with empty bounding rect must still be included."""
        empty_rect = _make_bounding_rect(empty=True)
        ctrl = self._make_window_control_child(handle=8000, bounding_rect=empty_rect)

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.MINIMIZED),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.return_value = ctrl
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={8000})

        assert len(windows) == 1
        assert windows[0].handle == 8000
        assert windows[0].status == Status.MINIMIZED

    def test_get_controls_handles_called_when_none_provided(self, svc):
        """When controls_handles is None, get_controls_handles() is called."""
        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_controls_handles", return_value=set()) as mock_gch,
        ):
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles=None)

        mock_gch.assert_called_once()
        assert windows == []

    def test_outer_exception_logged_and_empty_list_returned(self, svc, caplog):
        """An unexpected outer exception returns empty list and logs error."""
        with (
            patch.object(svc, "get_controls_handles", side_effect=RuntimeError("total failure")),
            caplog.at_level(logging.ERROR, logger="windows_mcp.window.service"),
        ):
            windows, handles = svc.get_windows()

        assert windows == []
        assert "get_windows" in caplog.text

    def test_multiple_windows_all_added(self, svc):
        """All valid windows in handles set are returned."""
        ctrl_a = self._make_window_control_child(handle=9001, name="App A")
        ctrl_b = self._make_window_control_child(handle=9002, name="App B")

        handle_to_ctrl = {9001: ctrl_a, 9002: ctrl_b}

        with (
            patch(_UIA) as mock_uia,
            patch.object(svc, "get_window_status", return_value=Status.NORMAL),
            patch.object(svc, "is_window_browser", return_value=False),
        ):
            mock_uia.ControlFromHandle.side_effect = lambda h: handle_to_ctrl[h]
            from windows_mcp import uia as uia_module

            mock_uia.WindowControl = uia_module.WindowControl
            mock_uia.PaneControl = uia_module.PaneControl
            mock_uia.PatternId = uia_module.PatternId

            windows, handles = svc.get_windows(controls_handles={9001, 9002})

        assert len(windows) == 2
        window_handles = {w.handle for w in windows}
        assert window_handles == {9001, 9002}


# ===========================================================================
# get_window_from_element
# ===========================================================================


class TestGetWindowFromElement:
    """Tests for WindowService.get_window_from_element()."""

    def test_returns_none_when_element_is_none(self, svc):
        result = svc.get_window_from_element(None)
        assert result is None

    def test_returns_none_when_top_level_control_is_none(self, svc):
        ctrl = _make_control()
        ctrl.GetTopLevelControl.return_value = None

        with patch.object(svc, "get_windows", return_value=([], set())):
            result = svc.get_window_from_element(ctrl)

        assert result is None

    def test_returns_matching_window(self, svc):
        ctrl = _make_control(handle=100)
        top_level = _make_control(handle=100)
        ctrl.GetTopLevelControl.return_value = top_level

        win = _make_window(handle=100)

        with patch.object(svc, "get_windows", return_value=([win], {100})):
            result = svc.get_window_from_element(ctrl)

        assert result is win

    def test_returns_none_when_no_matching_window(self, svc):
        ctrl = _make_control(handle=200)
        top_level = _make_control(handle=200)
        ctrl.GetTopLevelControl.return_value = top_level

        other_win = _make_window(handle=999)

        with patch.object(svc, "get_windows", return_value=([other_win], {999})):
            result = svc.get_window_from_element(ctrl)

        assert result is None

    def test_handle_mismatch_returns_none(self, svc):
        ctrl = _make_control(handle=300)
        top_level = _make_control(handle=300)
        ctrl.GetTopLevelControl.return_value = top_level

        # window list has different handles
        wins = [_make_window(handle=301), _make_window(handle=302)]

        with patch.object(svc, "get_windows", return_value=(wins, {301, 302})):
            result = svc.get_window_from_element(ctrl)

        assert result is None


# ===========================================================================
# bring_window_to_top
# ===========================================================================


class TestBringWindowToTop:
    """Tests for WindowService.bring_window_to_top()."""

    def test_raises_value_error_for_invalid_handle(self, svc):
        with patch(_WIN32GUI) as mock_gui:
            mock_gui.IsWindow.return_value = False

            with pytest.raises(ValueError, match="Invalid window handle"):
                svc.bring_window_to_top(9999)

    def test_iconic_window_is_restored_before_focus(self, svc):
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS),
            patch(_CTYPES),
        ):
            import win32con as _wc

            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.return_value = True
            mock_gui.GetForegroundWindow.return_value = 0  # triggers simple path
            mock_gui.IsWindow.side_effect = lambda h: True  # always valid

            svc.bring_window_to_top(100)

            mock_gui.ShowWindow.assert_called_once_with(100, _wc.SW_RESTORE)

    def test_no_valid_foreground_calls_set_foreground_directly(self, svc):
        """When foreground handle is not a valid window, SetForegroundWindow is called."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS),
            patch(_CTYPES),
        ):
            # First call IsWindow(target) -> True, second IsWindow(foreground) -> False
            mock_gui.IsWindow.side_effect = [True, False]
            mock_gui.IsIconic.return_value = False
            mock_gui.GetForegroundWindow.return_value = 200

            svc.bring_window_to_top(100)

            mock_gui.SetForegroundWindow.assert_called_once_with(100)
            mock_gui.BringWindowToTop.assert_called_once_with(100)

    def test_same_thread_no_attach_calls_set_foreground(self, svc):
        """When foreground and target share the same thread, no attach is needed."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS) as mock_process,
            patch(_CTYPES),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.return_value = False
            mock_gui.GetForegroundWindow.return_value = 500

            # Same thread ID for both windows
            mock_process.GetWindowThreadProcessId.return_value = (42, 1000)

            svc.bring_window_to_top(100)

            mock_gui.SetForegroundWindow.assert_called_once_with(100)
            mock_gui.BringWindowToTop.assert_called_once_with(100)
            mock_process.AttachThreadInput.assert_not_called()

    def test_different_threads_attaches_and_detaches(self, svc):
        """Different threads trigger AllowSetForegroundWindow + AttachThreadInput."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS) as mock_process,
            patch(_CTYPES) as mock_ctypes,
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.return_value = False
            mock_gui.GetForegroundWindow.return_value = 600

            # Different thread IDs: foreground=10, target=20
            mock_process.GetWindowThreadProcessId.side_effect = [(10, 1000), (20, 2000)]

            svc.bring_window_to_top(100)

            mock_ctypes.windll.user32.AllowSetForegroundWindow.assert_called_once_with(-1)
            mock_process.AttachThreadInput.assert_any_call(10, 20, True)
            mock_process.AttachThreadInput.assert_any_call(10, 20, False)

            mock_gui.SetForegroundWindow.assert_called_once_with(100)
            mock_gui.BringWindowToTop.assert_called_once_with(100)
            mock_gui.SetWindowPos.assert_called_once()

    def test_attach_input_detached_even_when_set_foreground_raises(self, svc):
        """AttachThreadInput(detach) must be called in finally even on error."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS) as mock_process,
            patch(_CTYPES),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.return_value = False
            mock_gui.GetForegroundWindow.return_value = 700
            mock_gui.SetForegroundWindow.side_effect = Exception("access denied")

            mock_process.GetWindowThreadProcessId.side_effect = [(30, 1000), (40, 2000)]

            # Should NOT raise -- outer try/except catches and logs
            svc.bring_window_to_top(100)

            # Even though SetForegroundWindow raised, detach should be called
            mock_process.AttachThreadInput.assert_any_call(30, 40, False)

    def test_zero_thread_id_falls_back_to_simple_path(self, svc):
        """When foreground_thread or target_thread is 0, use simple path."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS) as mock_process,
            patch(_CTYPES),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.return_value = False
            mock_gui.GetForegroundWindow.return_value = 800

            # foreground_thread = 0 -> falsy
            mock_process.GetWindowThreadProcessId.side_effect = [(0, 999), (50, 2000)]

            svc.bring_window_to_top(100)

            mock_gui.SetForegroundWindow.assert_called_once_with(100)
            mock_gui.BringWindowToTop.assert_called_once_with(100)
            mock_process.AttachThreadInput.assert_not_called()

    def test_exception_logged_not_raised(self, svc, caplog):
        """Exceptions in the inner try block are logged, not propagated."""
        with (
            patch(_WIN32GUI) as mock_gui,
            patch(_WIN32PROCESS),
            patch(_CTYPES),
            caplog.at_level(logging.ERROR, logger="windows_mcp.window.service"),
        ):
            mock_gui.IsWindow.return_value = True
            mock_gui.IsIconic.side_effect = Exception("unexpected")

            # Should not raise
            svc.bring_window_to_top(100)

        assert "bring window to top" in caplog.text.lower()


# ===========================================================================
# auto_minimize
# ===========================================================================


class TestAutoMinimize:
    """Tests for WindowService.auto_minimize() context manager."""

    def test_zero_handle_yields_without_minimize(self, svc):
        """handle=0 (falsy) must yield immediately without calling ShowWindow."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 0

            entered = []
            with svc.auto_minimize():
                entered.append(True)

            assert entered == [True]
            mock_uia.ShowWindow.assert_not_called()

    def test_nonzero_handle_minimizes_on_enter(self, svc):
        """Non-zero handle must call ShowWindow(SW_MINIMIZE) on enter."""
        import win32con as _wc

        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 1234

            with svc.auto_minimize():
                pass

            mock_uia.ShowWindow.assert_any_call(1234, _wc.SW_MINIMIZE)

    def test_nonzero_handle_restores_on_exit(self, svc):
        """Non-zero handle must call ShowWindow(SW_RESTORE) on exit."""
        import win32con as _wc

        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 5678

            with svc.auto_minimize():
                pass

            mock_uia.ShowWindow.assert_any_call(5678, _wc.SW_RESTORE)

    def test_restore_called_even_when_body_raises(self, svc):
        """ShowWindow(SW_RESTORE) must be called in finally even if body raises."""
        import win32con as _wc

        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 9999

            with pytest.raises(ValueError, match="body error"):
                with svc.auto_minimize():
                    raise ValueError("body error")

            mock_uia.ShowWindow.assert_any_call(9999, _wc.SW_RESTORE)

    def test_minimize_then_restore_order(self, svc):
        """SW_MINIMIZE must be called before SW_RESTORE."""
        import win32con as _wc

        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 1111

            with svc.auto_minimize():
                pass

            calls = mock_uia.ShowWindow.call_args_list
            assert len(calls) == 2
            assert calls[0] == call(1111, _wc.SW_MINIMIZE)
            assert calls[1] == call(1111, _wc.SW_RESTORE)

    def test_zero_handle_no_exception_on_empty_body(self, svc):
        """zero handle with empty body is a no-op (no exception, no ShowWindow)."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 0

            with svc.auto_minimize():
                pass  # empty body

            mock_uia.ShowWindow.assert_not_called()


# ===========================================================================
# Browser.has_process (sanity checks -- used by is_window_browser)
# ===========================================================================


class TestBrowserHasProcess:
    """Sanity tests for Browser.has_process used internally."""

    def test_chrome_exe_recognised(self):
        assert Browser.has_process("chrome.exe") is True

    def test_msedge_exe_recognised(self):
        assert Browser.has_process("msedge.exe") is True

    def test_firefox_exe_recognised(self):
        assert Browser.has_process("firefox.exe") is True

    def test_uppercase_chrome_recognised(self):
        """Browser.has_process lower-cases before matching."""
        assert Browser.has_process("CHROME.EXE") is True

    def test_notepad_not_recognised(self):
        assert Browser.has_process("notepad.exe") is False

    def test_empty_string_not_recognised(self):
        assert Browser.has_process("") is False


# ===========================================================================
# _get_windows_native (Rust fast-path)
# ===========================================================================


class TestGetWindowsNative:
    """Tests for WindowService._get_windows_native() and the Rust fast-path in get_windows()."""

    def _make_native_window(
        self,
        hwnd=1000,
        title="Test App",
        class_name="TestClass",
        pid=1234,
        left=0,
        top=0,
        right=800,
        bottom=600,
        is_minimized=False,
        is_maximized=False,
        is_visible=True,
    ):
        return {
            "hwnd": hwnd,
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "rect": {"left": left, "top": top, "right": right, "bottom": bottom},
            "is_minimized": is_minimized,
            "is_maximized": is_maximized,
            "is_visible": is_visible,
        }

    def test_returns_none_when_native_unavailable(self, svc):
        with patch(_NATIVE_LIST_WINDOWS, return_value=None):
            result = svc._get_windows_native()
        assert result is None

    def test_returns_windows_from_native(self, svc):
        windows_data = [self._make_native_window(hwnd=100, title="App A")]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            result = svc._get_windows_native()
        assert result is not None
        windows, handles = result
        assert len(windows) == 1
        assert windows[0].handle == 100
        assert windows[0].name == "App A"
        assert 100 in handles

    def test_overlay_title_filtered_out(self, svc):
        windows_data = [self._make_native_window(title="NVIDIA Overlay")]
        with patch(_NATIVE_LIST_WINDOWS, return_value=windows_data):
            windows, handles = svc._get_windows_native()
        assert len(windows) == 0

    def test_empty_title_filtered_out(self, svc):
        windows_data = [self._make_native_window(title="")]
        with patch(_NATIVE_LIST_WINDOWS, return_value=windows_data):
            windows, handles = svc._get_windows_native()
        assert len(windows) == 0

    def test_minimized_window_included(self, svc):
        windows_data = [
            self._make_native_window(
                hwnd=200,
                title="Minimized App",
                is_minimized=True,
                left=0,
                top=0,
                right=0,
                bottom=0,
            )
        ]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc._get_windows_native()
        assert len(windows) == 1
        assert windows[0].status == Status.MINIMIZED

    def test_maximized_window_status(self, svc):
        windows_data = [self._make_native_window(hwnd=300, is_maximized=True)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc._get_windows_native()
        assert len(windows) == 1
        assert windows[0].status == Status.MAXIMIZED

    def test_normal_window_status(self, svc):
        windows_data = [self._make_native_window(hwnd=400)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc._get_windows_native()
        assert len(windows) == 1
        assert windows[0].status == Status.NORMAL

    def test_zero_size_non_minimized_filtered(self, svc):
        """Zero-size windows that aren't minimized should be filtered out."""
        windows_data = [
            self._make_native_window(left=0, top=0, right=0, bottom=0, is_minimized=False)
        ]
        with patch(_NATIVE_LIST_WINDOWS, return_value=windows_data):
            windows, handles = svc._get_windows_native()
        assert len(windows) == 0

    def test_browser_detection_via_pid(self, svc):
        windows_data = [self._make_native_window(hwnd=500, pid=9999)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "chrome.exe"
            windows, handles = svc._get_windows_native()
        assert len(windows) == 1
        assert windows[0].is_browser is True

    def test_bounding_box_computed_correctly(self, svc):
        windows_data = [self._make_native_window(left=10, top=20, right=810, bottom=620)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, _ = svc._get_windows_native()
        bb = windows[0].bounding_box
        assert bb.left == 10
        assert bb.top == 20
        assert bb.right == 810
        assert bb.bottom == 620
        assert bb.width == 800
        assert bb.height == 600

    def test_vdm_filter_excludes_other_desktop(self, svc):
        """When controls_handles provided, windows not in set are VDM-filtered."""
        windows_data = [self._make_native_window(hwnd=600)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=False),
        ):
            windows, handles = svc._get_windows_native(controls_handles={999})
        assert len(windows) == 0

    def test_vdm_filter_includes_current_desktop(self, svc):
        windows_data = [self._make_native_window(hwnd=700)]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_IS_WINDOW_ON_CURRENT_DESKTOP, return_value=True),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc._get_windows_native(controls_handles={999})
        assert len(windows) == 1

    def test_get_windows_prefers_native_over_uia(self, svc):
        """get_windows() should use native path when available."""
        windows_data = [self._make_native_window(hwnd=800, title="Native App")]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
            patch(_UIA) as mock_uia,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc.get_windows(controls_handles={800})
        # UIA ControlFromHandle should NOT be called (native path was used)
        mock_uia.ControlFromHandle.assert_not_called()
        assert len(windows) == 1
        assert windows[0].name == "Native App"

    def test_multiple_native_windows(self, svc):
        windows_data = [
            self._make_native_window(hwnd=901, title="App A", pid=1001),
            self._make_native_window(hwnd=902, title="App B", pid=1002),
            self._make_native_window(hwnd=903, title="App C", pid=1003),
        ]
        with (
            patch(_NATIVE_LIST_WINDOWS, return_value=windows_data),
            patch(_PSUTIL_PROCESS) as MockProcess,
        ):
            MockProcess.return_value.name.return_value = "notepad.exe"
            windows, handles = svc._get_windows_native()
        assert len(windows) == 3
        assert handles == {901, 902, 903}


# ===========================================================================
# _is_browser_pid
# ===========================================================================


class TestIsBrowserPid:
    """Tests for WindowService._is_browser_pid()."""

    def test_chrome_pid(self, svc):
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "chrome.exe"
            assert svc._is_browser_pid(100) is True

    def test_non_browser_pid(self, svc):
        with patch(_PSUTIL_PROCESS) as MockProcess:
            MockProcess.return_value.name.return_value = "notepad.exe"
            assert svc._is_browser_pid(200) is False

    def test_zero_pid_returns_false(self, svc):
        assert svc._is_browser_pid(0) is False

    def test_negative_pid_returns_false(self, svc):
        assert svc._is_browser_pid(-1) is False

    def test_no_such_process_returns_false(self, svc):
        from psutil import NoSuchProcess

        with patch(_PSUTIL_PROCESS, side_effect=NoSuchProcess(999)):
            assert svc._is_browser_pid(999) is False

    def test_access_denied_returns_false(self, svc):
        from psutil import AccessDenied

        with patch(_PSUTIL_PROCESS, side_effect=AccessDenied(999)):
            assert svc._is_browser_pid(999) is False

    def test_uses_cache(self, svc):
        svc._process_name_cache[500] = "chrome.exe"
        with patch(_PSUTIL_PROCESS) as MockProcess:
            result = svc._is_browser_pid(500)
        MockProcess.assert_not_called()
        assert result is True
