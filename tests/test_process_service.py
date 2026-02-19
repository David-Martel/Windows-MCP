"""Unit tests for ProcessService.

Covers:
- is_protected() -- set membership and regex pattern matching
- list_processes() -- enumeration, fuzzy name filtering, sorting, limit, error skipping
- kill_process() -- by PID, by name, force flag, protected-process guard, error paths
"""

from unittest.mock import MagicMock, patch

import psutil
import pytest

from windows_mcp.process.service import ProcessService

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_proc(pid: int, name: str, cpu: float = 0.0, mem_rss: int = 0) -> MagicMock:
    """Return a mock psutil.Process as returned by process_iter."""
    mem_info = MagicMock()
    mem_info.rss = mem_rss
    proc = MagicMock()
    proc.info = {
        "pid": pid,
        "name": name,
        "cpu_percent": cpu,
        "memory_info": mem_info,
    }
    return proc


# ---------------------------------------------------------------------------
# 1. TestIsProtected
# ---------------------------------------------------------------------------


class TestIsProtected:
    """ProcessService.is_protected() -- set membership and regex pattern checks."""

    def test_csrss_protected(self):
        assert ProcessService.is_protected("csrss.exe") is True

    def test_lsass_protected(self):
        assert ProcessService.is_protected("lsass.exe") is True

    def test_svchost_protected(self):
        assert ProcessService.is_protected("svchost.exe") is True

    def test_services_protected(self):
        assert ProcessService.is_protected("services.exe") is True

    def test_smss_protected(self):
        assert ProcessService.is_protected("smss.exe") is True

    def test_wininit_protected(self):
        assert ProcessService.is_protected("wininit.exe") is True

    def test_winlogon_protected(self):
        assert ProcessService.is_protected("winlogon.exe") is True

    def test_msmpeng_protected(self):
        assert ProcessService.is_protected("msmpeng.exe") is True

    def test_system_protected(self):
        assert ProcessService.is_protected("system") is True

    def test_registry_protected(self):
        assert ProcessService.is_protected("registry") is True

    def test_memory_compression_protected(self):
        assert ProcessService.is_protected("memory compression") is True

    def test_case_insensitive_uppercase(self):
        """CSRSS.EXE (all caps) is matched case-insensitively."""
        assert ProcessService.is_protected("CSRSS.EXE") is True

    def test_case_insensitive_mixed(self):
        """Lsass.Exe with mixed casing is still protected."""
        assert ProcessService.is_protected("Lsass.Exe") is True

    def test_case_insensitive_svchost_upper(self):
        assert ProcessService.is_protected("SVCHOST.EXE") is True

    def test_normal_process_not_protected(self):
        """notepad.exe is not in the protected set."""
        assert ProcessService.is_protected("notepad.exe") is False

    def test_chrome_not_protected(self):
        assert ProcessService.is_protected("chrome.exe") is False

    def test_explorer_not_protected(self):
        assert ProcessService.is_protected("explorer.exe") is False

    def test_partial_name_without_extension_not_in_set(self):
        """'svchost' without .exe is not an exact set member.

        The regex also requires the .exe suffix, so it returns False.
        """
        assert ProcessService.is_protected("svchost") is False

    def test_system_idle_process_protected_by_regex(self):
        """'system idle process.exe' is covered by the regex pattern."""
        assert ProcessService.is_protected("system idle process.exe") is True

    def test_empty_string_not_protected(self):
        assert ProcessService.is_protected("") is False

    def test_substring_embedded_not_matched(self):
        """'not_csrss.exe' shares a substring but is not protected."""
        assert ProcessService.is_protected("not_csrss.exe") is False


# ---------------------------------------------------------------------------
# 2. TestListProcesses
# ---------------------------------------------------------------------------


class TestListProcesses:
    """ProcessService.list_processes() -- enumeration, filtering, sorting, limit."""

    @pytest.fixture()
    def svc(self) -> ProcessService:
        return ProcessService()

    def test_basic_listing_returns_table(self, svc: ProcessService):
        procs = [
            _make_proc(1, "alpha.exe", cpu=1.0, mem_rss=10 * 1024 * 1024),
            _make_proc(2, "beta.exe", cpu=2.0, mem_rss=20 * 1024 * 1024),
        ]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes()

        assert "Processes" in result
        assert "alpha.exe" in result
        assert "beta.exe" in result

    def test_filter_by_name_uses_fuzzy_match(self, svc: ProcessService):
        procs = [
            _make_proc(1, "notepad.exe"),
            _make_proc(2, "chrome.exe"),
        ]

        def _fake_partial_ratio(a: str, b: str) -> int:
            # Only notepad passes the > 60 threshold
            return 80 if "notepad" in b else 30

        with patch("psutil.process_iter", return_value=procs):
            with patch("thefuzz.fuzz.partial_ratio", side_effect=_fake_partial_ratio):
                result = svc.list_processes(name="notepad")

        assert "notepad.exe" in result
        assert "chrome.exe" not in result

    def test_filter_by_name_no_match_returns_message(self, svc: ProcessService):
        procs = [_make_proc(1, "notepad.exe")]

        with patch("psutil.process_iter", return_value=procs):
            with patch("thefuzz.fuzz.partial_ratio", return_value=10):
                result = svc.list_processes(name="xyzzy")

        assert "No processes found" in result
        assert "xyzzy" in result

    def test_sort_by_memory_descending(self, svc: ProcessService):
        procs = [
            _make_proc(1, "low.exe", mem_rss=5 * 1024 * 1024),
            _make_proc(2, "high.exe", mem_rss=50 * 1024 * 1024),
        ]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes(sort_by="memory")

        # high.exe should appear before low.exe in the table
        assert result.index("high.exe") < result.index("low.exe")

    def test_sort_by_cpu_descending(self, svc: ProcessService):
        procs = [
            _make_proc(1, "idle.exe", cpu=0.1),
            _make_proc(2, "busy.exe", cpu=99.0),
        ]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes(sort_by="cpu")

        assert result.index("busy.exe") < result.index("idle.exe")

    def test_sort_by_name_ascending(self, svc: ProcessService):
        procs = [
            _make_proc(1, "zebra.exe"),
            _make_proc(2, "apple.exe"),
        ]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes(sort_by="name")

        assert result.index("apple.exe") < result.index("zebra.exe")

    def test_limit_truncates_list(self, svc: ProcessService):
        procs = [_make_proc(i, f"proc{i:02d}.exe") for i in range(10)]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes(limit=3)

        assert "3 shown" in result

    def test_limit_one_always_shows_at_least_one(self, svc: ProcessService):
        procs = [_make_proc(1, "only.exe")]
        with patch("psutil.process_iter", return_value=procs):
            result = svc.list_processes(limit=0)

        # max(1, 0) means at least 1 process is shown
        assert "only.exe" in result

    def test_no_processes_returns_message(self, svc: ProcessService):
        with patch("psutil.process_iter", return_value=[]):
            result = svc.list_processes()

        assert "No processes found" in result

    def test_access_denied_process_is_skipped(self, svc: ProcessService):
        good_proc = _make_proc(1, "good.exe")
        bad_proc = MagicMock()
        bad_proc.info  # trigger property access
        bad_proc.__iter__ = MagicMock(side_effect=psutil.AccessDenied(pid=2))

        # Simulate AccessDenied raised when accessing .info
        denied_proc = MagicMock()
        type(denied_proc).info = property(fget=MagicMock(side_effect=psutil.AccessDenied(pid=2)))

        with patch("psutil.process_iter", return_value=[denied_proc, good_proc]):
            result = svc.list_processes()

        assert "good.exe" in result

    def test_no_such_process_is_skipped(self, svc: ProcessService):
        gone_proc = MagicMock()
        type(gone_proc).info = property(fget=MagicMock(side_effect=psutil.NoSuchProcess(pid=99)))
        good_proc = _make_proc(1, "alive.exe")

        with patch("psutil.process_iter", return_value=[gone_proc, good_proc]):
            result = svc.list_processes()

        assert "alive.exe" in result

    def test_memory_info_none_treated_as_zero(self, svc: ProcessService):
        """A process whose memory_info is None contributes 0 MB without crashing."""
        proc = MagicMock()
        proc.info = {"pid": 5, "name": "nomem.exe", "cpu_percent": 0.0, "memory_info": None}
        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.list_processes()

        assert "nomem.exe" in result

    def test_name_none_replaced_with_unknown(self, svc: ProcessService):
        """A process whose name is None is listed as 'Unknown'."""
        proc = MagicMock()
        mem = MagicMock()
        mem.rss = 0
        proc.info = {"pid": 6, "name": None, "cpu_percent": 0.0, "memory_info": mem}
        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.list_processes()

        assert "Unknown" in result


# ---------------------------------------------------------------------------
# 3. TestKillProcess
# ---------------------------------------------------------------------------


class TestKillProcess:
    """ProcessService.kill_process() -- termination paths and safety checks."""

    @pytest.fixture()
    def svc(self) -> ProcessService:
        return ProcessService()

    # ------------------------------------------------------------------
    # Error: neither pid nor name supplied
    # ------------------------------------------------------------------

    def test_no_pid_no_name_returns_error(self, svc: ProcessService):
        result = svc.kill_process()

        assert "Error" in result
        assert "pid" in result.lower() or "name" in result.lower()

    # ------------------------------------------------------------------
    # Kill by PID -- happy path
    # ------------------------------------------------------------------

    def test_kill_by_pid_terminate(self, svc: ProcessService):
        mock_proc = MagicMock()
        mock_proc.name.return_value = "notepad.exe"

        with patch("psutil.Process", return_value=mock_proc):
            result = svc.kill_process(pid=1234)

        mock_proc.terminate.assert_called_once()
        mock_proc.kill.assert_not_called()
        assert "1234" in result

    def test_kill_by_pid_force_uses_kill(self, svc: ProcessService):
        mock_proc = MagicMock()
        mock_proc.name.return_value = "notepad.exe"

        with patch("psutil.Process", return_value=mock_proc):
            result = svc.kill_process(pid=1234, force=True)

        mock_proc.kill.assert_called_once()
        mock_proc.terminate.assert_not_called()
        assert "Force killed" in result

    def test_kill_by_pid_returns_terminated_message(self, svc: ProcessService):
        mock_proc = MagicMock()
        mock_proc.name.return_value = "calc.exe"

        with patch("psutil.Process", return_value=mock_proc):
            result = svc.kill_process(pid=42)

        assert "Terminated" in result
        assert "calc.exe" in result

    # ------------------------------------------------------------------
    # Kill by PID -- protected process guard
    # ------------------------------------------------------------------

    def test_protected_process_blocked_by_pid(self, svc: ProcessService):
        """When the PID resolves to csrss.exe the kill is refused."""
        mock_proc = MagicMock()
        mock_proc.name.return_value = "csrss.exe"

        with patch("psutil.Process", return_value=mock_proc):
            result = svc.kill_process(pid=4)

        assert "Refused" in result
        assert "csrss.exe" in result
        mock_proc.terminate.assert_not_called()
        mock_proc.kill.assert_not_called()

    def test_protected_process_blocked_by_pid_lsass(self, svc: ProcessService):
        mock_proc = MagicMock()
        mock_proc.name.return_value = "lsass.exe"

        with patch("psutil.Process", return_value=mock_proc):
            result = svc.kill_process(pid=700)

        assert "Refused" in result

    # ------------------------------------------------------------------
    # Kill by PID -- error paths
    # ------------------------------------------------------------------

    def test_no_such_process_by_pid(self, svc: ProcessService):
        with patch("psutil.Process", side_effect=psutil.NoSuchProcess(pid=9999)):
            result = svc.kill_process(pid=9999)

        assert "No process with PID 9999" in result

    def test_access_denied_by_pid(self, svc: ProcessService):
        with patch("psutil.Process", side_effect=psutil.AccessDenied(pid=8888)):
            result = svc.kill_process(pid=8888)

        assert "Access denied" in result
        assert "8888" in result

    # ------------------------------------------------------------------
    # Kill by name -- happy path
    # ------------------------------------------------------------------

    def test_kill_by_name_terminate(self, svc: ProcessService):
        procs = [_make_proc(100, "notepad.exe"), _make_proc(200, "chrome.exe")]
        for p in procs:
            p.terminate = MagicMock()
            p.kill = MagicMock()

        with patch("psutil.process_iter", return_value=procs):
            result = svc.kill_process(name="notepad.exe")

        procs[0].terminate.assert_called_once()
        procs[1].terminate.assert_not_called()
        assert "notepad.exe" in result

    def test_kill_by_name_force_uses_kill(self, svc: ProcessService):
        proc = _make_proc(300, "notepad.exe")
        proc.kill = MagicMock()
        proc.terminate = MagicMock()

        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.kill_process(name="notepad.exe", force=True)

        proc.kill.assert_called_once()
        proc.terminate.assert_not_called()
        assert "Force killed" in result

    def test_kill_by_name_case_insensitive(self, svc: ProcessService):
        """Name matching is case-insensitive (Notepad.EXE matches notepad.exe)."""
        proc = _make_proc(400, "notepad.exe")
        proc.kill = MagicMock()
        proc.terminate = MagicMock()

        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.kill_process(name="Notepad.EXE")

        proc.terminate.assert_called_once()
        assert "notepad.exe" in result

    # ------------------------------------------------------------------
    # Kill by name -- protected process guard
    # ------------------------------------------------------------------

    def test_protected_process_blocked_by_name(self, svc: ProcessService):
        """When the matching process is svchost.exe the kill is refused."""
        proc = _make_proc(500, "svchost.exe")
        proc.kill = MagicMock()
        proc.terminate = MagicMock()

        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.kill_process(name="svchost.exe")

        assert "Refused" in result
        assert "svchost.exe" in result
        proc.terminate.assert_not_called()
        proc.kill.assert_not_called()

    def test_protected_process_blocked_by_name_services(self, svc: ProcessService):
        proc = _make_proc(600, "services.exe")
        proc.kill = MagicMock()
        proc.terminate = MagicMock()

        with patch("psutil.process_iter", return_value=[proc]):
            result = svc.kill_process(name="services.exe")

        assert "Refused" in result

    # ------------------------------------------------------------------
    # Kill by name -- error paths
    # ------------------------------------------------------------------

    def test_no_matches_by_name_returns_message(self, svc: ProcessService):
        with patch("psutil.process_iter", return_value=[]):
            result = svc.kill_process(name="ghostapp.exe")

        assert "No process matching" in result
        assert "ghostapp.exe" in result

    def test_access_denied_during_name_iteration_is_skipped(self, svc: ProcessService):
        """AccessDenied on a process during name-based kill does not abort iteration."""
        good_proc = _make_proc(700, "target.exe")
        good_proc.terminate = MagicMock()
        good_proc.kill = MagicMock()

        denied_proc = MagicMock()
        type(denied_proc).info = property(fget=MagicMock(side_effect=psutil.AccessDenied(pid=701)))

        with patch("psutil.process_iter", return_value=[denied_proc, good_proc]):
            result = svc.kill_process(name="target.exe")

        good_proc.terminate.assert_called_once()
        assert "target.exe" in result

    def test_no_such_process_during_name_iteration_is_skipped(self, svc: ProcessService):
        """NoSuchProcess on one process during iteration is silently skipped."""
        good_proc = _make_proc(800, "alive.exe")
        good_proc.terminate = MagicMock()
        good_proc.kill = MagicMock()

        gone_proc = MagicMock()
        type(gone_proc).info = property(fget=MagicMock(side_effect=psutil.NoSuchProcess(pid=801)))

        with patch("psutil.process_iter", return_value=[gone_proc, good_proc]):
            result = svc.kill_process(name="alive.exe")

        good_proc.terminate.assert_called_once()
        assert "alive.exe" in result

    def test_kill_multiple_processes_with_same_name(self, svc: ProcessService):
        """All matching processes are killed when multiple share the same name."""
        proc1 = _make_proc(900, "notepad.exe")
        proc2 = _make_proc(901, "notepad.exe")
        for p in (proc1, proc2):
            p.terminate = MagicMock()
            p.kill = MagicMock()

        with patch("psutil.process_iter", return_value=[proc1, proc2]):
            result = svc.kill_process(name="notepad.exe")

        proc1.terminate.assert_called_once()
        proc2.terminate.assert_called_once()
        assert "Terminated" in result
