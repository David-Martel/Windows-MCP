"""Tests for ShellService.execute() subprocess invocation paths.

Covers lines 94-128 of src/windows_mcp/shell/service.py:
- Blocked command early-return path
- Successful subprocess run returning stdout
- Fallback to stderr when stdout is empty
- subprocess.TimeoutExpired handling
- Generic OSError (and Exception subclass) handling
- Custom timeout forwarded to subprocess.run
- Output decoding using self.encoding
"""

import base64
import subprocess
from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.shell.service import ShellService

# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------


@pytest.fixture()
def svc() -> ShellService:
    """Return a ShellService instance with encoding pinned to utf-8 for determinism."""
    service = ShellService()
    service.encoding = "utf-8"
    return service


def _make_completed_process(
    stdout: bytes = b"",
    stderr: bytes = b"",
    returncode: int = 0,
) -> subprocess.CompletedProcess:
    """Build a subprocess.CompletedProcess with the given fields."""
    result = MagicMock(spec=subprocess.CompletedProcess)
    result.stdout = stdout
    result.stderr = stderr
    result.returncode = returncode
    return result


def _encoded_command(command: str) -> str:
    """Return the base64-encoded UTF-16LE form that execute() builds internally."""
    return base64.b64encode(command.encode("utf-16le")).decode("ascii")


# ---------------------------------------------------------------------------
# TestExecuteBlockedCommand
# ---------------------------------------------------------------------------


class TestExecuteBlockedCommand:
    """execute() returns an error tuple immediately when the command is blocked."""

    def test_blocked_command_returns_error_message(self, svc: ShellService):
        # "format C:" matches the first default blocklist pattern.
        output, code = svc.execute("format C:")

        assert "blocked by safety filter" in output
        assert code == 1

    def test_blocked_command_includes_matched_pattern(self, svc: ShellService):
        output, code = svc.execute("diskpart")

        assert "diskpart" in output.lower() or "matched pattern" in output
        assert code == 1

    def test_blocked_command_does_not_invoke_subprocess(self, svc: ShellService):
        with patch("subprocess.run") as mock_run:
            svc.execute("format C:")
            mock_run.assert_not_called()


# ---------------------------------------------------------------------------
# TestExecuteSuccess
# ---------------------------------------------------------------------------


class TestExecuteSuccess:
    """execute() returns decoded stdout and returncode=0 on success."""

    def test_returns_stdout_on_success(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"hello\n", returncode=0)

        with patch("subprocess.run", return_value=completed):
            output, code = svc.execute("Get-Process")

        assert output == "hello\n"
        assert code == 0

    def test_returncode_forwarded(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"ok", returncode=42)

        with patch("subprocess.run", return_value=completed):
            _, code = svc.execute("Get-Process")

        assert code == 42

    def test_command_is_base64_encoded_in_subprocess_call(self, svc: ShellService):
        command = "Get-Date"
        completed = _make_completed_process(stdout=b"Monday", returncode=0)

        with patch("subprocess.run", return_value=completed) as mock_run:
            svc.execute(command)

        args, _ = mock_run.call_args
        argv = args[0]
        assert "-EncodedCommand" in argv
        encoded_index = argv.index("-EncodedCommand") + 1
        assert argv[encoded_index] == _encoded_command(command)

    def test_powershell_flags_present(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"result", returncode=0)

        with patch("subprocess.run", return_value=completed) as mock_run:
            svc.execute("Get-Date")

        argv = mock_run.call_args[0][0]
        assert argv[0] == "powershell"
        assert "-NoProfile" in argv
        assert "-OutputFormat" in argv
        assert "Text" in argv


# ---------------------------------------------------------------------------
# TestExecuteStderrFallback
# ---------------------------------------------------------------------------


class TestExecuteStderrFallback:
    """execute() falls back to stderr output when stdout is empty."""

    def test_stderr_returned_when_stdout_empty(self, svc: ShellService):
        completed = _make_completed_process(
            stdout=b"",
            stderr=b"error text",
            returncode=1,
        )

        with patch("subprocess.run", return_value=completed):
            output, code = svc.execute("Get-Process")

        assert output == "error text"
        assert code == 1

    def test_stdout_preferred_over_stderr_when_both_present(self, svc: ShellService):
        completed = _make_completed_process(
            stdout=b"primary output",
            stderr=b"warning line",
            returncode=0,
        )

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert output == "primary output"

    def test_empty_string_returned_when_both_streams_empty(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"", stderr=b"", returncode=0)

        with patch("subprocess.run", return_value=completed):
            output, code = svc.execute("Get-Process")

        # stdout or stderr -- both empty, result is falsy empty string
        assert output == ""
        assert code == 0


# ---------------------------------------------------------------------------
# TestExecuteTimeout
# ---------------------------------------------------------------------------


class TestExecuteTimeout:
    """execute() handles subprocess.TimeoutExpired gracefully."""

    def test_timeout_returns_error_tuple(self, svc: ShellService):
        with patch(
            "subprocess.run", side_effect=subprocess.TimeoutExpired(cmd="powershell", timeout=10)
        ):
            output, code = svc.execute("Start-Sleep -Seconds 60")

        assert "timed out" in output.lower()
        assert code == 1

    def test_timeout_does_not_propagate_exception(self, svc: ShellService):
        with patch(
            "subprocess.run", side_effect=subprocess.TimeoutExpired(cmd="powershell", timeout=10)
        ):
            # Must not raise -- exception must be swallowed and converted to tuple.
            result = svc.execute("Start-Sleep -Seconds 60")

        assert isinstance(result, tuple)
        assert len(result) == 2


# ---------------------------------------------------------------------------
# TestExecuteGenericError
# ---------------------------------------------------------------------------


class TestExecuteGenericError:
    """execute() handles OSError and other unexpected exceptions."""

    def test_oserror_returns_error_tuple(self, svc: ShellService):
        with patch("subprocess.run", side_effect=OSError("powershell not found")):
            output, code = svc.execute("Get-Process")

        assert "OSError" in output
        assert "powershell not found" in output
        assert code == 1

    def test_oserror_does_not_propagate(self, svc: ShellService):
        with patch("subprocess.run", side_effect=OSError("exec failed")):
            result = svc.execute("Get-Process")

        assert isinstance(result, tuple)

    def test_generic_exception_returns_error_tuple(self, svc: ShellService):
        with patch("subprocess.run", side_effect=RuntimeError("unexpected")):
            output, code = svc.execute("Get-Process")

        assert "RuntimeError" in output
        assert "unexpected" in output
        assert code == 1

    def test_generic_exception_includes_type_name(self, svc: ShellService):
        with patch("subprocess.run", side_effect=ValueError("bad value")):
            output, _ = svc.execute("Get-Process")

        assert "ValueError" in output


# ---------------------------------------------------------------------------
# TestExecuteTimeout_Parameter
# ---------------------------------------------------------------------------


class TestExecuteTimeoutParameter:
    """execute() passes the timeout argument through to subprocess.run."""

    def test_default_timeout_is_ten_seconds(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"ok", returncode=0)

        with patch("subprocess.run", return_value=completed) as mock_run:
            svc.execute("Get-Date")

        _, kwargs = mock_run.call_args
        assert kwargs.get("timeout") == 10

    def test_custom_timeout_forwarded(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"ok", returncode=0)

        with patch("subprocess.run", return_value=completed) as mock_run:
            svc.execute("Get-Date", timeout=30)

        _, kwargs = mock_run.call_args
        assert kwargs.get("timeout") == 30

    def test_timeout_one_second(self, svc: ShellService):
        completed = _make_completed_process(stdout=b"fast", returncode=0)

        with patch("subprocess.run", return_value=completed) as mock_run:
            svc.execute("dir", timeout=1)

        _, kwargs = mock_run.call_args
        assert kwargs.get("timeout") == 1


# ---------------------------------------------------------------------------
# TestExecuteDecoding
# ---------------------------------------------------------------------------


class TestExecuteDecoding:
    """execute() decodes stdout/stderr bytes using self.encoding."""

    def test_utf8_output_decoded_correctly(self, svc: ShellService):
        svc.encoding = "utf-8"
        completed = _make_completed_process(stdout="caf\u00e9".encode("utf-8"), returncode=0)

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert output == "caf\u00e9"

    def test_latin1_encoding_used(self, svc: ShellService):
        svc.encoding = "latin-1"
        # b"\xe9" is 'é' in latin-1
        completed = _make_completed_process(stdout=b"caf\xe9", returncode=0)

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert output == "caf\xe9".encode("latin-1").decode("latin-1")

    def test_invalid_bytes_replaced_with_ignore(self, svc: ShellService):
        svc.encoding = "utf-8"
        # b"\xff\xfe" is invalid in strict utf-8 but should be silently dropped (errors="ignore")
        completed = _make_completed_process(stdout=b"ok\xff\xfedone", returncode=0)

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert "ok" in output
        assert "done" in output

    def test_stderr_also_decoded_using_encoding(self, svc: ShellService):
        svc.encoding = "utf-8"
        completed = _make_completed_process(
            stdout=b"",
            stderr="err\u00e9".encode("utf-8"),
            returncode=1,
        )

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert output == "err\u00e9"

    def test_string_stdout_not_double_decoded(self, svc: ShellService):
        # If subprocess somehow returns a str (not bytes), the isinstance check
        # must skip decoding and leave it untouched.
        completed = _make_completed_process(returncode=0)
        completed.stdout = "already a string"
        completed.stderr = b""

        with patch("subprocess.run", return_value=completed):
            output, _ = svc.execute("Get-Process")

        assert output == "already a string"
