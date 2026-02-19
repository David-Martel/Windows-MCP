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
    """list_processes filters, sorts, and limits the process list."""

    def _run(
        self,
        d,
        proc_infos: list,
        *,
        name=None,
        sort_by="memory",
        limit=20,
    ) -> str:
        mock_psutil = _make_mock_psutil(proc_infos)
        tab_fn = MagicMock(return_value="table_output")
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            result = d.list_processes(name=name, sort_by=sort_by, limit=limit)
        return result

    def test_no_filter_returns_table(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(1, "alpha"), _make_proc_info(2, "beta")]
        result = self._run(d, infos)
        assert "Processes" in result

    def test_empty_process_list_no_filter(self):
        d = _make_bare_desktop()
        result = self._run(d, [])
        assert "No processes found" in result
        assert "matching" not in result

    def test_empty_after_name_filter(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(1, "notepad")]
        mock_fuzz = MagicMock()
        mock_fuzz.partial_ratio.return_value = 10  # below threshold -- all filtered out
        mock_psutil = _make_mock_psutil(infos)
        tab_fn = MagicMock(return_value="table_output")
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn), _inject_thefuzz_fuzz(mock_fuzz):
            result = d.list_processes(name="chrome")
        assert "No processes found" in result
        assert "matching chrome" in result

    def test_name_filter_includes_high_ratio_process(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(1, "notepad"), _make_proc_info(2, "chrome")]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        mock_fuzz = MagicMock()
        # notepad scores low, chrome scores high
        mock_fuzz.partial_ratio.side_effect = lambda a, b: 90 if "chrome" in b else 10
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn), _inject_thefuzz_fuzz(mock_fuzz):
            d.list_processes(name="chrome")
        rows = tab_fn.call_args[0][0]
        pids = [r[0] for r in rows]
        assert 2 in pids
        assert 1 not in pids

    def test_sort_by_memory_default_descending(self):
        d = _make_bare_desktop()
        infos = [
            _make_proc_info(1, "low", mem_rss=10 * 1024 * 1024),
            _make_proc_info(2, "high", mem_rss=100 * 1024 * 1024),
            _make_proc_info(3, "mid", mem_rss=50 * 1024 * 1024),
        ]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(sort_by="memory")
        rows = tab_fn.call_args[0][0]
        pids = [r[0] for r in rows]
        assert pids == [2, 3, 1], f"Expected [2,3,1] (desc memory), got {pids}"

    def test_sort_by_cpu_descending(self):
        d = _make_bare_desktop()
        infos = [
            _make_proc_info(1, "low", cpu=5.0),
            _make_proc_info(2, "high", cpu=80.0),
            _make_proc_info(3, "mid", cpu=30.0),
        ]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(sort_by="cpu")
        rows = tab_fn.call_args[0][0]
        pids = [r[0] for r in rows]
        assert pids == [2, 3, 1], f"Expected [2,3,1] (desc cpu), got {pids}"

    def test_sort_by_name_ascending(self):
        d = _make_bare_desktop()
        infos = [
            _make_proc_info(1, "zebra"),
            _make_proc_info(2, "alpha"),
            _make_proc_info(3, "monkey"),
        ]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(sort_by="name")
        rows = tab_fn.call_args[0][0]
        pids = [r[0] for r in rows]
        assert pids == [2, 3, 1], f"Expected [2,3,1] (asc name), got {pids}"

    def test_limit_positive_respected(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(i, f"proc{i}") for i in range(10)]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(limit=3)
        rows = tab_fn.call_args[0][0]
        assert len(rows) == 3

    def test_limit_negative_clamped_to_one(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(i, f"proc{i}") for i in range(5)]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(limit=-5)
        rows = tab_fn.call_args[0][0]
        assert len(rows) == 1

    def test_limit_zero_clamped_to_one(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(i, f"proc{i}") for i in range(5)]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes(limit=0)
        rows = tab_fn.call_args[0][0]
        assert len(rows) == 1

    def test_result_contains_count(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(i, f"p{i}") for i in range(3)]
        tab_fn = MagicMock(return_value="faked_table")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            result = d.list_processes()
        assert "3 shown" in result

    def test_none_name_skips_fuzzy_filter(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(1, "notepad"), _make_proc_info(2, "chrome")]
        tab_fn = MagicMock(return_value="table_output")
        mock_fuzz = MagicMock()
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn), _inject_thefuzz_fuzz(mock_fuzz):
            d.list_processes(name=None)
        # fuzz.partial_ratio must NOT be called when name is None
        mock_fuzz.partial_ratio.assert_not_called()

    def test_nosuchprocess_and_accessdenied_are_skipped(self):
        """Processes that raise NoSuchProcess or AccessDenied are silently skipped."""
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        good_proc = MagicMock()
        good_mem = MagicMock()
        good_mem.rss = 1024 * 1024
        good_proc.info = {"pid": 42, "name": "good", "cpu_percent": 0.0, "memory_info": good_mem}

        bad_proc = MagicMock()
        bad_proc.info = {"pid": 99, "name": "bad", "cpu_percent": 0.0, "memory_info": None}
        # Accessing bad_proc.info raises the exception during iteration
        # Simulate: the for loop body raises for this proc
        # We achieve this by making info a property that raises
        type(bad_proc).info = property(lambda self: (_ for _ in ()).throw(NoSuch("gone")))

        mock_psutil.process_iter.return_value = [good_proc, bad_proc]

        d = _make_bare_desktop()
        tab_fn = MagicMock(return_value="table_output")
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            result = d.list_processes()
        assert "Processes" in result
        rows = tab_fn.call_args[0][0]
        pids = [r[0] for r in rows]
        assert 42 in pids
        assert 99 not in pids

    def test_table_headers_correct(self):
        d = _make_bare_desktop()
        infos = [_make_proc_info(1, "notepad")]
        tab_fn = MagicMock(return_value="table_output")
        mock_psutil = _make_mock_psutil(infos)
        with _inject_psutil(mock_psutil), _inject_tabulate(tab_fn):
            d.list_processes()
        _, kwargs = tab_fn.call_args
        headers = kwargs.get("headers") or tab_fn.call_args[0][1]
        assert "PID" in headers
        assert "Name" in headers
        assert "CPU%" in headers
        assert "Memory" in headers


# ===========================================================================
# 3. kill_process
# ===========================================================================


class TestKillProcess:
    """kill_process kills by PID or by name, with optional force mode."""

    def _make_pid_psutil(self, pid, name, *, force_raises=None, term_raises=None):
        """Build a psutil mock for PID-based kill tests."""
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        proc = MagicMock()
        proc.name.return_value = name
        if force_raises:
            proc.kill.side_effect = force_raises
        if term_raises:
            proc.terminate.side_effect = term_raises
        mock_psutil.Process.return_value = proc
        return mock_psutil, NoSuch, AccessDenied, proc

    def test_neither_pid_nor_name_returns_error(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        mock_psutil.NoSuchProcess = Exception
        mock_psutil.AccessDenied = Exception
        with _inject_psutil(mock_psutil):
            result = d.kill_process(name=None, pid=None)
        assert "Error" in result
        assert "pid" in result.lower() or "name" in result.lower()

    def test_kill_by_pid_terminate_called(self):
        d = _make_bare_desktop()
        mock_psutil, _, _, proc = self._make_pid_psutil(123, "notepad.exe")
        with _inject_psutil(mock_psutil):
            result = d.kill_process(pid=123, force=False)
        proc.terminate.assert_called_once()
        proc.kill.assert_not_called()
        assert "123" in result
        assert "notepad.exe" in result

    def test_kill_by_pid_force_uses_kill(self):
        d = _make_bare_desktop()
        mock_psutil, _, _, proc = self._make_pid_psutil(456, "chrome.exe")
        with _inject_psutil(mock_psutil):
            result = d.kill_process(pid=456, force=True)
        proc.kill.assert_called_once()
        proc.terminate.assert_not_called()
        assert "Force killed" in result

    def test_kill_by_pid_no_such_process(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied
        mock_psutil.Process.side_effect = NoSuch("gone")
        with _inject_psutil(mock_psutil):
            result = d.kill_process(pid=9999)
        assert "9999" in result
        assert "No process" in result

    def test_kill_by_pid_access_denied(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied
        mock_psutil.Process.side_effect = AccessDenied("denied")
        with _inject_psutil(mock_psutil):
            result = d.kill_process(pid=1234)
        assert "Access denied" in result
        assert "1234" in result

    def test_kill_by_name_terminates_matching(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        p1 = MagicMock()
        p1.info = {"pid": 10, "name": "notepad.exe"}
        p2 = MagicMock()
        p2.info = {"pid": 20, "name": "chrome.exe"}
        mock_psutil.process_iter.return_value = [p1, p2]

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="notepad.exe", force=False)

        p1.terminate.assert_called_once()
        p2.terminate.assert_not_called()
        assert "notepad.exe" in result
        assert "Terminated" in result

    def test_kill_by_name_force_uses_kill(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        p1 = MagicMock()
        p1.info = {"pid": 11, "name": "target.exe"}
        mock_psutil.process_iter.return_value = [p1]

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="target.exe", force=True)

        p1.kill.assert_called_once()
        assert "Force killed" in result

    def test_kill_by_name_case_insensitive(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        p1 = MagicMock()
        p1.info = {"pid": 77, "name": "Notepad.EXE"}
        mock_psutil.process_iter.return_value = [p1]

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="notepad.exe")

        p1.terminate.assert_called_once()
        assert "Notepad.EXE" in result

    def test_kill_by_name_no_match_returns_no_process_message(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        p1 = MagicMock()
        p1.info = {"pid": 5, "name": "explorer.exe"}
        mock_psutil.process_iter.return_value = [p1]

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="ghost.exe")

        assert "No process matching" in result
        assert "ghost.exe" in result

    def test_kill_by_name_multiple_processes_all_terminated(self):
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        procs = []
        for i in range(3):
            p = MagicMock()
            p.info = {"pid": 100 + i, "name": "worker.exe"}
            procs.append(p)
        mock_psutil.process_iter.return_value = procs

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="worker.exe")

        for p in procs:
            p.terminate.assert_called_once()
        assert "Terminated" in result
        assert "worker.exe" in result

    def test_kill_by_name_nosuchprocess_skipped(self):
        """Processes that disappear mid-iteration are silently skipped."""
        d = _make_bare_desktop()
        mock_psutil = MagicMock()
        NoSuch = type("NoSuchProcess", (Exception,), {})
        AccessDenied = type("AccessDenied", (Exception,), {})
        mock_psutil.NoSuchProcess = NoSuch
        mock_psutil.AccessDenied = AccessDenied

        vanished = MagicMock()
        vanished.info = {"pid": 55, "name": "target.exe"}
        vanished.terminate.side_effect = NoSuch("gone")

        surviving = MagicMock()
        surviving.info = {"pid": 56, "name": "target.exe"}

        mock_psutil.process_iter.return_value = [vanished, surviving]

        with _inject_psutil(mock_psutil):
            result = d.kill_process(name="target.exe")

        surviving.terminate.assert_called_once()
        # Only the surviving process appears in the killed list
        assert "56" in result

    def test_terminate_label_when_not_force(self):
        d = _make_bare_desktop()
        mock_psutil, _, _, _ = self._make_pid_psutil(1, "app.exe")
        with _inject_psutil(mock_psutil):
            result = d.kill_process(pid=1, force=False)
        assert "Terminated" in result
        assert "Force killed" not in result


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
                "windows_mcp.desktop.service.get_current_desktop",
                return_value={"id": "0", "name": "Default"},
            ),
            patch(
                "windows_mcp.desktop.service.get_all_desktops",
                return_value=[{"id": "0", "name": "Default"}],
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
                "windows_mcp.desktop.service.get_current_desktop",
                return_value={"id": "0", "name": "Default"},
            ),
            patch(
                "windows_mcp.desktop.service.get_all_desktops",
                return_value=[{"id": "0", "name": "Default"}],
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
                "windows_mcp.desktop.service.get_current_desktop",
                return_value={"id": "0", "name": "Default"},
            ),
            patch(
                "windows_mcp.desktop.service.get_all_desktops",
                return_value=[{"id": "0", "name": "Default"}],
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
                "windows_mcp.desktop.service.get_current_desktop",
                return_value={"id": "0", "name": "Default"},
            ),
            patch(
                "windows_mcp.desktop.service.get_all_desktops",
                return_value=[{"id": "0", "name": "Default"}],
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
                "windows_mcp.desktop.service.get_current_desktop",
                return_value={"id": "0", "name": "Default"},
            ),
            patch(
                "windows_mcp.desktop.service.get_all_desktops",
                return_value=[{"id": "0", "name": "Default"}],
            ),
        ):
            state = d.get_state(use_vision=False)

        assert state.tree_state.scrollable_nodes == []
