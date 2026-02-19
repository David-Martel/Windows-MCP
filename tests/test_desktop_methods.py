"""Unit tests for uncovered Desktop service methods.

Tests for: send_notification, list_processes, kill_process, get_system_info,
lock_screen, and get_state tree_state exception handling.

All UIA/COM/win32/psutil interactions are mocked so the suite runs headless
with no live desktop required.
"""

import sys
from contextlib import contextmanager
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from tests.desktop_helpers import make_bare_desktop

# Alias for backwards compat with existing test code
_make_bare_desktop = make_bare_desktop


# ---------------------------------------------------------------------------
# Helpers for psutil injection
# ---------------------------------------------------------------------------


def _make_proc_info(pid: int, name: str, cpu: float = 0.0, mem_rss: int = 1024 * 1024):
    mem = MagicMock()
    mem.rss = mem_rss
    return {"pid": pid, "name": name, "cpu_percent": cpu, "memory_info": mem}


def _make_mock_psutil(proc_infos: list) -> MagicMock:
    """Build a psutil mock whose process_iter yields the given info dicts."""
    mock_psutil = MagicMock()
    proc_mocks = []
    for info in proc_infos:
        p = MagicMock()
        p.info = info
        proc_mocks.append(p)
    mock_psutil.process_iter.return_value = proc_mocks
    mock_psutil.NoSuchProcess = type("NoSuchProcess", (Exception,), {})
    mock_psutil.AccessDenied = type("AccessDenied", (Exception,), {})
    return mock_psutil


@contextmanager
def _inject_psutil(mock_psutil):
    """Context manager that injects a mock psutil into sys.modules."""
    orig = sys.modules.get("psutil")
    sys.modules["psutil"] = mock_psutil
    try:
        yield
    finally:
        if orig is None:
            sys.modules.pop("psutil", None)
        else:
            sys.modules["psutil"] = orig


@contextmanager
def _inject_tabulate(mock_tabulate_fn):
    """Context manager that injects a mock tabulate module into sys.modules."""
    orig = sys.modules.get("tabulate")
    mock_mod = MagicMock()
    mock_mod.tabulate = mock_tabulate_fn
    sys.modules["tabulate"] = mock_mod
    try:
        yield
    finally:
        if orig is None:
            sys.modules.pop("tabulate", None)
        else:
            sys.modules["tabulate"] = orig


@contextmanager
def _inject_thefuzz_fuzz(mock_fuzz_obj):
    """Inject a mock fuzz object by patching the thefuzz package in sys.modules.

    list_processes uses ``from thefuzz import fuzz`` (local import).  Python
    resolves this by looking up sys.modules['thefuzz'] and reading its .fuzz
    attribute, so we inject a mock thefuzz module whose .fuzz attribute points
    to our mock.
    """
    orig_thefuzz = sys.modules.get("thefuzz")
    orig_thefuzz_fuzz = sys.modules.get("thefuzz.fuzz")

    mock_thefuzz_mod = MagicMock()
    mock_thefuzz_mod.fuzz = mock_fuzz_obj
    sys.modules["thefuzz"] = mock_thefuzz_mod
    sys.modules["thefuzz.fuzz"] = mock_fuzz_obj
    try:
        yield
    finally:
        if orig_thefuzz is None:
            sys.modules.pop("thefuzz", None)
        else:
            sys.modules["thefuzz"] = orig_thefuzz
        if orig_thefuzz_fuzz is None:
            sys.modules.pop("thefuzz.fuzz", None)
        else:
            sys.modules["thefuzz.fuzz"] = orig_thefuzz_fuzz


# ===========================================================================
# 1. send_notification
# ===========================================================================


class TestSendNotification:
    """send_notification builds a PowerShell script, calls execute_command,
    and returns a formatted string based on the exit status."""

    def test_success_returns_formatted_string(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("", 0))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            result = d.send_notification("Hello", "World")
        assert result == 'Notification sent: "Hello" - World'

    def test_failure_returns_fallback_string(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("some error output", 1))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            result = d.send_notification("Title", "Body")
        assert "Notification may have been sent" in result
        assert "PowerShell output" in result

    def test_failure_truncates_response_at_200_chars(self):
        d = _make_bare_desktop()
        long_output = "x" * 500
        d.execute_command = MagicMock(return_value=(long_output, 2))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            result = d.send_notification("T", "M")
        # The truncated portion should not exceed 200 characters embedded in result
        assert "x" * 201 not in result
        assert "x" * 200 in result

    def test_xml_escapes_ampersand_in_title(self):
        d = _make_bare_desktop()
        captured_scripts = []

        def capture_cmd(script, timeout=10):
            captured_scripts.append(script)
            return ("", 0)

        d.execute_command = capture_cmd
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification("Tom & Jerry", "Message")
        assert "&amp;" in captured_scripts[0]
        assert "&" not in captured_scripts[0].replace("&amp;", "").replace("&quot;", "").replace(
            "&apos;", ""
        )

    def test_xml_escapes_less_than_in_message(self):
        d = _make_bare_desktop()
        captured_scripts = []

        def capture_cmd(script, timeout=10):
            captured_scripts.append(script)
            return ("", 0)

        d.execute_command = capture_cmd
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification("Title", "a < b")
        assert "&lt;" in captured_scripts[0]

    def test_xml_escapes_greater_than_in_message(self):
        d = _make_bare_desktop()
        captured_scripts = []

        def capture_cmd(script, timeout=10):
            captured_scripts.append(script)
            return ("", 0)

        d.execute_command = capture_cmd
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification("Title", "a > b")
        assert "&gt;" in captured_scripts[0]

    def test_xml_escapes_double_quote_in_title(self):
        d = _make_bare_desktop()
        captured_scripts = []

        def capture_cmd(script, timeout=10):
            captured_scripts.append(script)
            return ("", 0)

        d.execute_command = capture_cmd
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification('Say "hello"', "body")
        assert "&quot;" in captured_scripts[0]

    def test_xml_escapes_single_quote_in_title(self):
        d = _make_bare_desktop()
        captured_scripts = []

        def capture_cmd(script, timeout=10):
            captured_scripts.append(script)
            return ("", 0)

        d.execute_command = capture_cmd
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification("It's here", "body")
        assert "&apos;" in captured_scripts[0]

    def test_execute_command_called_once(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("", 0))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            d.send_notification("Title", "Body")
        d.execute_command.assert_called_once()

    def test_ps_quote_called_for_title_and_message(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("", 0))
        ps_quote_calls = []

        def recording_ps_quote(v):
            ps_quote_calls.append(v)
            return f"'{v}'"

        with patch.object(type(d), "_ps_quote", staticmethod(recording_ps_quote)):
            d.send_notification("My Title", "My Body")
        # ps_quote must have been called with the xml-escaped title and message
        assert any("My Title" in c for c in ps_quote_calls)
        assert any("My Body" in c for c in ps_quote_calls)

    def test_success_title_appears_in_result(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("", 0))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            result = d.send_notification("Alert", "Done")
        assert "Alert" in result
        assert "Done" in result

    def test_status_zero_is_success_not_fallback(self):
        d = _make_bare_desktop()
        d.execute_command = MagicMock(return_value=("any output", 0))
        with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
            result = d.send_notification("T", "M")
        assert "may have been sent" not in result

    def test_non_zero_status_triggers_fallback(self):
        for status in (1, 2, 127, -1):
            d = _make_bare_desktop()
            d.execute_command = MagicMock(return_value=("err", status))
            with patch.object(type(d), "_ps_quote", staticmethod(lambda v: f"'{v}'")):
                result = d.send_notification("T", "M")
            assert "may have been sent" in result, f"Expected fallback for status={status}"


# ===========================================================================
# 2. list_processes
# ===========================================================================


class TestListProcesses:
    """Desktop.list_processes delegates to ProcessService."""

    def test_delegates_with_all_args(self):
        d = _make_bare_desktop()
        d._process.list_processes.return_value = "Processes (3 shown):\ntable"
        result = d.list_processes(name="chrome", sort_by="cpu", limit=10)
        d._process.list_processes.assert_called_once_with(name="chrome", sort_by="cpu", limit=10)
        assert "Processes" in result

    def test_delegates_with_defaults(self):
        d = _make_bare_desktop()
        d._process.list_processes.return_value = "ok"
        d.list_processes()
        d._process.list_processes.assert_called_once_with(name=None, sort_by="memory", limit=20)


# ===========================================================================
# 3. kill_process
# ===========================================================================


class TestKillProcess:
    """Desktop.kill_process delegates to ProcessService."""

    def test_delegates_by_pid(self):
        d = _make_bare_desktop()
        d._process.kill_process.return_value = "Terminated: notepad.exe (PID 123)"
        result = d.kill_process(pid=123, force=False)
        d._process.kill_process.assert_called_once_with(name=None, pid=123, force=False)
        assert "Terminated" in result

    def test_delegates_by_name(self):
        d = _make_bare_desktop()
        d._process.kill_process.return_value = "Terminated: ghost.exe (PID 5)"
        result = d.kill_process(name="ghost.exe")
        d._process.kill_process.assert_called_once_with(name="ghost.exe", pid=None, force=False)
        assert "ghost.exe" in result

    def test_delegates_force(self):
        d = _make_bare_desktop()
        d._process.kill_process.return_value = "Force killed: app.exe (PID 1)"
        result = d.kill_process(pid=1, force=True)
        d._process.kill_process.assert_called_once_with(name=None, pid=1, force=True)
        assert "Force killed" in result


# ===========================================================================
# 4. get_system_info
# ===========================================================================


class TestGetSystemInfo:
    """get_system_info uses Rust native_system_info when available,
    falls back to psutil otherwise. Network and uptime always come from psutil."""

    def _make_psutil_for_system_info(self):
        mock_psutil = MagicMock()
        mock_psutil.cpu_percent.return_value = 25.0
        mock_psutil.cpu_count.return_value = 8
        mock_psutil.virtual_memory.return_value = SimpleNamespace(
            percent=60.0, used=8 * 1024**3, total=16 * 1024**3
        )
        mock_psutil.disk_usage.return_value = SimpleNamespace(
            percent=45.0, used=200 * 1024**3, total=500 * 1024**3
        )
        mock_psutil.boot_time.return_value = 1700000000.0
        mock_psutil.net_io_counters.return_value = SimpleNamespace(
            bytes_sent=100 * 1024**2, bytes_recv=500 * 1024**2
        )
        return mock_psutil

    def test_fallback_path_contains_key_sections(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert "CPU" in result
        assert "Memory" in result
        assert "Disk" in result
        assert "Network" in result
        assert "Uptime" in result

    def test_fallback_path_uses_psutil_cpu_percent(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert "25.0%" in result

    def test_fallback_path_uses_psutil_memory(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert "60.0%" in result

    def test_fallback_path_uses_psutil_disk(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert "45.0%" in result

    def test_fallback_network_bytes(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert "100.0 MB sent" in result
        assert "500.0 MB received" in result

    def test_native_path_cpu_from_rust_data(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        native_data = {
            "cpu_count": 16,
            "cpu_usage_percent": [10.0, 20.0, 30.0, 40.0],
            "total_memory_bytes": 32 * 1024**3,
            "used_memory_bytes": 16 * 1024**3,
            "disks": [
                {
                    "mount_point": "C:\\",
                    "total_bytes": 500 * 1024**3,
                    "available_bytes": 300 * 1024**3,
                }
            ],
        }
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=native_data),
        ):
            result = d.get_system_info()
        # avg of [10,20,30,40] = 25.0
        assert "25.0%" in result
        assert "16 cores" in result

    def test_native_path_c_drive_disk_used(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        native_data = {
            "cpu_count": 4,
            "cpu_usage_percent": [50.0],
            "total_memory_bytes": 8 * 1024**3,
            "used_memory_bytes": 4 * 1024**3,
            "disks": [
                {
                    "mount_point": "C:\\",
                    "total_bytes": 100 * 1024**3,
                    "available_bytes": 60 * 1024**3,
                }
            ],
        }
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=native_data),
        ):
            result = d.get_system_info()
        # used = 100 - 60 = 40 GB; pct = 40%
        assert "40.0%" in result

    def test_native_path_skips_psutil_cpu_percent(self):
        """When native data is present, psutil.cpu_percent should NOT be called."""
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        native_data = {
            "cpu_count": 4,
            "cpu_usage_percent": [20.0],
            "total_memory_bytes": 8 * 1024**3,
            "used_memory_bytes": 2 * 1024**3,
            "disks": [],
        }
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=native_data),
        ):
            d.get_system_info()
        mock_psutil.cpu_percent.assert_not_called()

    def test_native_path_still_uses_psutil_boot_time(self):
        """Network/uptime must always come from psutil, even on the native path."""
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        native_data = {
            "cpu_count": 4,
            "cpu_usage_percent": [10.0],
            "total_memory_bytes": 8 * 1024**3,
            "used_memory_bytes": 2 * 1024**3,
            "disks": [],
        }
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=native_data),
        ):
            d.get_system_info()
        mock_psutil.boot_time.assert_called_once()
        mock_psutil.net_io_counters.assert_called_once()

    def test_native_none_triggers_fallback(self):
        """native_system_info() returning None must activate the psutil fallback."""
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            d.get_system_info()
        mock_psutil.cpu_percent.assert_called_once()

    def test_native_missing_cpu_count_triggers_fallback(self):
        """native data without 'cpu_count' key falls back to psutil."""
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        # 'cpu_count' key missing -> falls back
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value={"other": "data"}),
        ):
            d.get_system_info()
        mock_psutil.cpu_percent.assert_called_once()

    def test_result_is_string(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
        ):
            result = d.get_system_info()
        assert isinstance(result, str)

    def test_result_contains_os_info(self):
        d = _make_bare_desktop()
        mock_psutil = self._make_psutil_for_system_info()
        with (
            _inject_psutil(mock_psutil),
            patch("windows_mcp.native.native_system_info", return_value=None),
            patch("platform.system", return_value="Windows"),
            patch("platform.release", return_value="11"),
        ):
            result = d.get_system_info()
        assert "OS" in result


# ===========================================================================
# 5. lock_screen
# ===========================================================================


class TestLockScreen:
    """lock_screen calls LockWorkStation and returns 'Screen locked.'"""

    def test_returns_screen_locked(self):
        d = _make_bare_desktop()
        with patch("ctypes.windll") as mock_windll:
            mock_windll.user32.LockWorkStation.return_value = 1
            result = d.lock_screen()
        assert result == "Screen locked."

    def test_calls_lock_workstation(self):
        d = _make_bare_desktop()
        with patch("ctypes.windll") as mock_windll:
            d.lock_screen()
        mock_windll.user32.LockWorkStation.assert_called_once()

    def test_return_value_is_exact_string(self):
        d = _make_bare_desktop()
        with patch("ctypes.windll"):
            result = d.lock_screen()
        assert result == "Screen locked."
        assert result.endswith(".")

    def test_lock_workstation_called_regardless_of_return_value(self):
        """LockWorkStation return value is ignored; method always returns the string."""
        d = _make_bare_desktop()
        with patch("ctypes.windll") as mock_windll:
            mock_windll.user32.LockWorkStation.return_value = 0  # simulated failure
            result = d.lock_screen()
        assert result == "Screen locked."


# ===========================================================================
# 6. get_state -- tree_state exception handling
# ===========================================================================


class TestGetStateTreeException:
    """When tree.get_state() raises, get_state must catch and use empty TreeState."""

    _UIA = "windows_mcp.desktop.service.uia"

    def _setup_get_state_minimal(self, d):
        """Patch all collaborators so get_state() can reach the tree.get_state() call."""
        d.get_controls_handles = MagicMock(return_value=set())
        d.get_windows = MagicMock(return_value=([], set()))
        d.get_active_window = MagicMock(return_value=None)

    def test_tree_exception_gives_empty_tree_state(self):
        d = _make_bare_desktop()
        self._setup_get_state_minimal(d)
        d.tree.get_state.side_effect = RuntimeError("COM exploded")

        from windows_mcp.tree.views import TreeState

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                return_value=({"id": "0", "name": "Default"}, [{"id": "0", "name": "Default"}]),
            ),
        ):
            state = d.get_state(use_vision=False)

        assert isinstance(state.tree_state, TreeState)
        assert state.tree_state.interactive_nodes == []
        assert state.tree_state.dom_informative_nodes == []

    def test_tree_exception_does_not_propagate(self):
        """get_state must not re-raise the tree exception."""
        d = _make_bare_desktop()
        self._setup_get_state_minimal(d)
        d.tree.get_state.side_effect = Exception("any tree error")

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                return_value=({"id": "0", "name": "Default"}, [{"id": "0", "name": "Default"}]),
            ),
        ):
            try:
                d.get_state(use_vision=False)
            except Exception as exc:
                pytest.fail(f"get_state propagated tree exception: {exc}")

    def test_tree_value_error_also_caught(self):
        d = _make_bare_desktop()
        self._setup_get_state_minimal(d)
        d.tree.get_state.side_effect = ValueError("bad tree state")

        from windows_mcp.tree.views import TreeState

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                return_value=({"id": "0", "name": "Default"}, [{"id": "0", "name": "Default"}]),
            ),
        ):
            state = d.get_state(use_vision=False)

        assert isinstance(state.tree_state, TreeState)

    def test_desktop_state_still_populated_on_tree_exception(self):
        """Even on tree failure, the returned DesktopState must be a valid object."""
        d = _make_bare_desktop()
        self._setup_get_state_minimal(d)
        d.tree.get_state.side_effect = RuntimeError("tree gone")

        from windows_mcp.desktop.views import DesktopState

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                return_value=({"id": "0", "name": "Default"}, [{"id": "0", "name": "Default"}]),
            ),
        ):
            state = d.get_state(use_vision=False)

        assert isinstance(state, DesktopState)
        assert state.active_desktop == {"id": "0", "name": "Default"}

    def test_empty_tree_state_has_no_scrollable_nodes(self):
        """Default TreeState has empty scrollable_nodes list."""
        d = _make_bare_desktop()
        self._setup_get_state_minimal(d)
        d.tree.get_state.side_effect = RuntimeError("fail")

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                return_value=({"id": "0", "name": "Default"}, [{"id": "0", "name": "Default"}]),
            ),
        ):
            state = d.get_state(use_vision=False)

        assert state.tree_state.scrollable_nodes == []


# ---------------------------------------------------------------------------
# launch_app -- coverage lines 313-341
# ---------------------------------------------------------------------------


class TestLaunchApp:
    """Unit tests for Desktop.launch_app() covering all branches."""

    def test_app_not_found_in_start_menu(self):
        """No match in start menu returns error tuple."""
        d = _make_bare_desktop()
        d.get_apps_from_start_menu = MagicMock(return_value={"notepad": "notepad.exe"})
        with patch("windows_mcp.desktop.service.process") as mock_fuzz:
            mock_fuzz.extractOne.return_value = None
            msg, status, pid = d.launch_app("nonexistent_app_xyz")
        assert status == 1
        assert pid == 0
        assert "not found" in msg.lower()

    def test_app_found_with_path_launch(self):
        """App ID is a filesystem path -- uses Start-Process with path."""
        d = _make_bare_desktop()
        d.get_apps_from_start_menu = MagicMock(return_value={"notepad": r"C:\Windows\notepad.exe"})
        with (
            patch("windows_mcp.desktop.service.process") as mock_fuzz,
            patch("windows_mcp.desktop.service.os.path.exists", return_value=True),
        ):
            mock_fuzz.extractOne.return_value = ("notepad", 90)
            d.execute_command = MagicMock(return_value=("1234", 0))
            msg, status, pid = d.launch_app("notepad")
        assert pid == 1234
        assert status == 0

    def test_app_found_with_shell_folder_launch(self):
        """App ID is a UWP AUMID -- uses shell:AppsFolder launch."""
        d = _make_bare_desktop()
        # UWP AUMIDs contain '!' delimiter between package name and entry point
        d.get_apps_from_start_menu = MagicMock(
            return_value={"calculator": "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"}
        )
        with (
            patch("windows_mcp.desktop.service.process") as mock_fuzz,
            patch("windows_mcp.desktop.service.os.path.exists", return_value=False),
        ):
            mock_fuzz.extractOne.return_value = ("calculator", 85)
            d.execute_command = MagicMock(return_value=("", 0))
            msg, status, pid = d.launch_app("calculator")
        assert status == 0

    def test_app_invalid_identifier_returns_error(self):
        """App ID with invalid chars (not alphanumeric/underscore/dot/dash/backslash) rejects."""
        d = _make_bare_desktop()
        d.get_apps_from_start_menu = MagicMock(return_value={"badapp": "$(malicious;command)"})
        with (
            patch("windows_mcp.desktop.service.process") as mock_fuzz,
            patch("windows_mcp.desktop.service.os.path.exists", return_value=False),
        ):
            mock_fuzz.extractOne.return_value = ("badapp", 95)
            msg, status, pid = d.launch_app("badapp")
        assert status == 1
        assert "Invalid app identifier" in msg

    def test_matched_app_name_not_in_map_returns_error(self):
        """extractOne returns a match but app_map.get returns None (race condition)."""
        d = _make_bare_desktop()
        d.get_apps_from_start_menu = MagicMock(return_value={})
        with patch("windows_mcp.desktop.service.process") as mock_fuzz:
            mock_fuzz.extractOne.return_value = ("ghost", 80)
            msg, status, pid = d.launch_app("ghost")
        assert status == 1


# ---------------------------------------------------------------------------
# get_apps_from_start_menu + _get_apps_from_shortcuts -- coverage lines 164-224
# ---------------------------------------------------------------------------


class TestAppCache:
    """Unit tests for app cache and Start Menu scanning."""

    def test_cache_hit_returns_cached(self):
        """Warm cache returns immediately without shell execution."""
        from time import time

        d = _make_bare_desktop()
        cached = {"notepad": "notepad.exe"}
        d._app_cache = cached
        d._app_cache_time = time()
        result = d.get_apps_from_start_menu()
        assert result is cached

    def test_cache_miss_calls_shell(self):
        """Cold cache triggers Get-StartApps PowerShell command."""
        d = _make_bare_desktop()
        d._app_cache = None
        csv_output = '"Name","AppID"\n"Notepad","notepad.exe"\n"Calc","calc.exe"'
        d.execute_command = MagicMock(return_value=(csv_output, 0))
        result = d.get_apps_from_start_menu()
        assert "notepad" in result
        assert "calc" in result

    def test_cache_miss_bad_csv_falls_back_to_shortcuts(self):
        """Malformed CSV falls back to Start Menu shortcut scanning."""
        d = _make_bare_desktop()
        d._app_cache = None
        d.execute_command = MagicMock(return_value=("not-csv-at-all", 0))
        d._get_apps_from_shortcuts = MagicMock(return_value={"notepad": "notepad.lnk"})
        result = d.get_apps_from_start_menu()
        d._get_apps_from_shortcuts.assert_called_once()
        assert "notepad" in result

    def test_cache_miss_command_fails_falls_back_to_shortcuts(self):
        """Non-zero exit from Get-StartApps falls back to shortcuts."""
        d = _make_bare_desktop()
        d._app_cache = None
        d.execute_command = MagicMock(return_value=("error", 1))
        d._get_apps_from_shortcuts = MagicMock(return_value={"calc": "calc.lnk"})
        result = d.get_apps_from_start_menu()
        assert "calc" in result

    def test_shortcut_scanning(self, tmp_path):
        """_get_apps_from_shortcuts finds .lnk files in Start Menu folders."""
        import os

        d = _make_bare_desktop()
        # Create fake Start Menu structure
        programs = tmp_path / "Microsoft" / "Windows" / "Start Menu" / "Programs"
        programs.mkdir(parents=True)
        (programs / "Notepad.lnk").write_bytes(b"fake")
        sub = programs / "Accessories"
        sub.mkdir()
        (sub / "WordPad.lnk").write_bytes(b"fake")

        with (
            patch.dict(os.environ, {"PROGRAMDATA": str(tmp_path), "APPDATA": ""}),
        ):
            result = d._get_apps_from_shortcuts()
        assert "notepad" in result
        assert "wordpad" in result


# ---------------------------------------------------------------------------
# is_app_running -- coverage line 263-269
# ---------------------------------------------------------------------------


class TestIsAppRunning:
    """Unit tests for Desktop.is_app_running()."""

    def test_app_running_found(self):
        d = _make_bare_desktop()
        from tests.desktop_helpers import make_window

        win = make_window(name="Notepad")
        d._window.get_windows.return_value = ([win], {win.handle})
        with patch("windows_mcp.desktop.service.process") as mock_fuzz:
            mock_fuzz.extractOne.return_value = ("Notepad", 90)
            assert d.is_app_running("notepad") is True

    def test_app_running_not_found(self):
        d = _make_bare_desktop()
        from tests.desktop_helpers import make_window

        win = make_window(name="Chrome")
        d._window.get_windows.return_value = ([win], {win.handle})
        with patch("windows_mcp.desktop.service.process") as mock_fuzz:
            mock_fuzz.extractOne.return_value = None
            assert d.is_app_running("notepad") is False

    def test_app_running_exception_returns_false(self):
        d = _make_bare_desktop()
        d._window.get_windows.side_effect = RuntimeError("COM failure")
        assert d.is_app_running("notepad") is False


# ---------------------------------------------------------------------------
# VDM fallback path -- coverage lines 98-103
# ---------------------------------------------------------------------------


class TestGetStateVdmFallback:
    """Desktop.get_state() VDM RuntimeError fallback."""

    _UIA = "windows_mcp.desktop.service.uia"

    def test_vdm_runtime_error_uses_default_desktop(self):
        """When VDM raises RuntimeError, state uses a default desktop placeholder."""
        d = _make_bare_desktop()
        d._window.get_controls_handles.return_value = set()
        d._window.get_windows.return_value = ([], set())
        d._window.get_active_window.return_value = None
        d.tree.get_state.return_value = MagicMock()

        with (
            patch(self._UIA),
            patch(
                "windows_mcp.desktop.service.get_desktop_info",
                side_effect=RuntimeError("VDM not available"),
            ),
        ):
            state = d.get_state(use_vision=False)

        assert state.active_desktop["name"] == "Default Desktop"
        assert len(state.all_desktops) == 1
