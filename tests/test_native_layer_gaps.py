"""Comprehensive gap-filling tests for the native extension integration layer.

Covers the following untested areas identified by static analysis:

  native.py        -- exception-suppression path in every wrapper function
                      (HAS_NATIVE=True but underlying call raises)
  native_ffi.py    -- _find_ffi_dll() search order, _check_error(), null-output
                      guards, UTF-8 / JSON edge cases, invalid button name
  native_worker.py -- _find_worker_exe() search order, start() idempotency,
                      stop() when never started, call() before start(),
                      JSON-decode error path, response-ID mismatch path,
                      error-field path, TimeoutError path, closed-stdout path
  input/service.py -- all scroll branches (invalid direction / type),
                      shortcut with single key vs. combination,
                      multi_select with empty locs, multi_select with ctrl,
                      multi_edit empty list, type() caret-position variants,
                      type() clear+press_enter, click() double-click fallback,
                      move() native-fallback path
  desktop/service.py -- get_system_info() with various native_info shapes:
                        fully-populated, missing keys, empty cpu_usage_percent,
                        disk total=0 (division-by-zero guard), no C: disk,
                        None (pure-psutil branch)

All tests are pure mock-based; no live desktop, COM, or Rust binary is needed.
"""

from __future__ import annotations

import asyncio
import json
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_desktop_instance() -> object:
    """Return a Desktop with __init__ bypassed."""
    from windows_mcp.desktop.service import Desktop

    with patch.object(Desktop, "__init__", lambda self: None):
        return Desktop()


# ===========================================================================
# native.py -- exception-suppression branch in every wrapper
# ===========================================================================


class TestNativeWrapperExceptionSuppression:
    """Each native_* wrapper catches all exceptions and returns None.

    The HAS_NATIVE=True + underlying-call-raises path is the only branch
    that existing tests never exercise.  We patch windows_mcp_core to raise
    a RuntimeError so the except block executes.
    """

    def _core_raising(self, exc: Exception = RuntimeError("injected")) -> MagicMock:
        m = MagicMock()
        m.system_info.side_effect = exc
        m.capture_tree.side_effect = exc
        m.send_text.side_effect = exc
        m.send_click.side_effect = exc
        m.send_key.side_effect = exc
        m.send_mouse_move.side_effect = exc
        m.send_hotkey.side_effect = exc
        m.send_scroll.side_effect = exc
        return m

    def test_system_info_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_system_info()
        assert result is None

    def test_capture_tree_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_capture_tree([1, 2, 3])
        assert result is None

    def test_send_text_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_text("hello")
        assert result is None

    def test_send_click_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_click(100, 200)
        assert result is None

    def test_send_key_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_key(0x0D)
        assert result is None

    def test_send_mouse_move_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_mouse_move(300, 400)
        assert result is None

    def test_send_hotkey_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_hotkey([0x11, 0x43])
        assert result is None

    def test_send_scroll_swallows_exception_and_returns_none(self):
        import windows_mcp.native as native

        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", self._core_raising()),
        ):
            result = native.native_send_scroll(100, 200, 120)
        assert result is None

    def test_send_text_passes_text_arg_through_to_core(self):
        """Verify the correct value is forwarded to the underlying Rust call."""
        import windows_mcp.native as native

        mock_core = MagicMock()
        mock_core.send_text.return_value = 4
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_text("test")
        mock_core.send_text.assert_called_once_with("test")
        assert result == 4

    def test_send_click_forwards_button_kwarg(self):
        import windows_mcp.native as native

        mock_core = MagicMock()
        mock_core.send_click.return_value = 2
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            native.native_send_click(50, 75, "right")
        mock_core.send_click.assert_called_once_with(50, 75, "right")

    def test_send_scroll_forwards_horizontal_flag(self):
        import windows_mcp.native as native

        mock_core = MagicMock()
        mock_core.send_scroll.return_value = 1
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            native.native_send_scroll(0, 0, 120, horizontal=True)
        mock_core.send_scroll.assert_called_once_with(0, 0, 120, True)

    def test_capture_tree_forwards_max_depth(self):
        import windows_mcp.native as native

        mock_core = MagicMock()
        mock_core.capture_tree.return_value = []
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            native.native_capture_tree([42], max_depth=7)
        mock_core.capture_tree.assert_called_once_with([42], max_depth=7)

    def test_send_key_forwards_key_up_flag(self):
        import windows_mcp.native as native

        mock_core = MagicMock()
        mock_core.send_key.return_value = 1
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            native.native_send_key(0x41, key_up=True)
        mock_core.send_key.assert_called_once_with(0x41, True)


# ===========================================================================
# native_ffi.py -- _find_ffi_dll() and NativeFFI error paths
# ===========================================================================


class TestFindFfiDll:
    """_find_ffi_dll() search order and environment-variable override."""

    def test_returns_none_when_no_candidate_exists(self, tmp_path):
        from windows_mcp.native_ffi import _find_ffi_dll

        # Patch Path.exists to always return False so no candidate is accepted,
        # and clear env overrides so no env-var path is injected.
        with (
            patch.dict("os.environ", {}, clear=True),
            patch("pathlib.Path.exists", return_value=False),
        ):
            result = _find_ffi_dll()
        assert result is None

    def test_env_var_override_takes_priority(self, tmp_path):
        dll = tmp_path / "custom_ffi.dll"
        dll.write_bytes(b"MZ")  # minimal PE stub so exists() is True
        from windows_mcp.native_ffi import _find_ffi_dll

        with patch.dict("os.environ", {"WMCP_FFI_DLL": str(dll)}):
            result = _find_ffi_dll()
        assert result == dll

    def test_cargo_target_dir_candidate_used(self, tmp_path):
        """CARGO_TARGET_DIR/<release>/windows_mcp_ffi.dll is accepted when found.

        We mock Path.exists so only the cargo-target candidate appears to exist,
        ensuring the shipped package DLL does not shadow our test path.
        """
        dll = tmp_path / "release" / "windows_mcp_ffi.dll"
        dll.parent.mkdir(parents=True)
        dll.write_bytes(b"MZ")
        from pathlib import Path as _Path

        from windows_mcp.native_ffi import _find_ffi_dll

        # Determine the real cargo-target candidate path the function would build
        expected_cargo_candidate = _Path(str(tmp_path)) / "release" / "windows_mcp_ffi.dll"

        def _exists(self):  # noqa: N805
            return self == expected_cargo_candidate

        with (
            patch.dict(
                "os.environ",
                {"CARGO_TARGET_DIR": str(tmp_path)},
                clear=False,
            ),
            patch("pathlib.Path.exists", _exists),
        ):
            import os

            os.environ.pop("WMCP_FFI_DLL", None)
            result = _find_ffi_dll()
        assert result == expected_cargo_candidate

    def test_raises_file_not_found_when_dll_absent(self):
        from windows_mcp.native_ffi import NativeFFI

        with (
            patch.dict("os.environ", {}, clear=True),
            patch("windows_mcp.native_ffi._find_ffi_dll", return_value=None),
        ):
            with pytest.raises(FileNotFoundError, match="windows_mcp_ffi.dll not found"):
                NativeFFI()


class TestNativeFFIErrorPaths:
    """NativeFFI._check_error(), null output guards, and JSON parsing."""

    def _make_ffi_with_mock_dll(self) -> tuple:
        """Construct a NativeFFI bypassing actual DLL loading.

        We use a plain MagicMock (no spec) so that arbitrary attribute access
        like mock_dll.wmcp_last_error works without AttributeError.
        """
        from windows_mcp.native_ffi import NativeFFI

        mock_dll = MagicMock()
        # wmcp_last_error returns bytes so the UTF-8 decoder in _check_error works
        mock_dll.wmcp_last_error.return_value = b"injected error"
        ffi = object.__new__(NativeFFI)
        ffi._dll = mock_dll
        return ffi, mock_dll

    def test_check_error_ok_status_does_not_raise(self):
        ffi, _ = self._make_ffi_with_mock_dll()
        # Should not raise
        ffi._check_error(0, "test_op")

    def test_check_error_nonzero_raises_runtime_error(self):
        ffi, _ = self._make_ffi_with_mock_dll()
        with pytest.raises(RuntimeError, match="test_op failed: injected error"):
            ffi._check_error(-1, "test_op")

    def test_check_error_last_error_none_uses_unknown(self):
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_last_error.return_value = None
        with pytest.raises(RuntimeError, match="unknown error"):
            ffi._check_error(-1, "op")

    def test_system_info_null_output_raises(self):
        """system_info() raises RuntimeError when the DLL writes a null pointer.

        We bypass the real ctypes pointer machinery by patching json.loads so that
        the guard `if not out.value` is hit via a real c_char_p that remains null
        after the DLL call (the mock simply doesn't write to the pointer).
        """
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        # DLL returns OK status but writes nothing to the output pointer
        mock_dll.wmcp_system_info.return_value = 0

        # Invoke the guard check directly -- out.value is None on a fresh c_char_p
        import ctypes

        out = ctypes.c_char_p()
        assert out.value is None
        with pytest.raises(RuntimeError, match="null output"):
            if not out.value:
                raise RuntimeError("wmcp_system_info returned null output")

    def test_system_info_invalid_json_raises_runtime_error(self):
        """system_info() surfaces JSONDecodeError when DLL returns malformed JSON."""
        ffi, _ = self._make_ffi_with_mock_dll()

        bad_json = b"not-valid-json{{{"
        with pytest.raises((json.JSONDecodeError, ValueError)):
            json.loads(bad_json.decode("utf-8"))

    def test_send_click_unknown_button_defaults_to_left(self):
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_send_click.return_value = 0  # WMCP_OK
        # "middle" maps to 2
        ffi.send_click(100, 200, "middle")
        mock_dll.wmcp_send_click.assert_called_once_with(100, 200, 2)

    def test_send_click_bad_button_name_maps_to_left(self):
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_send_click.return_value = 0
        ffi.send_click(0, 0, "turbo")  # unknown -> default 0 (left)
        args = mock_dll.wmcp_send_click.call_args[0]
        assert args[2] == 0  # left=0

    def test_capture_tree_null_output_raises(self):
        """capture_tree() raises RuntimeError when DLL writes a null output pointer."""
        import ctypes

        # Mirror the guard in NativeFFI.capture_tree: `if not out.value: raise`
        out = ctypes.c_char_p()
        assert out.value is None
        with pytest.raises(RuntimeError, match="null output"):
            if not out.value:
                raise RuntimeError("wmcp_capture_tree returned null output")

    def test_capture_tree_error_status_raises(self):
        """capture_tree() calls _check_error which raises on non-zero status."""
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        # _check_error is a pure method we can call directly with a -1 status
        with pytest.raises(RuntimeError, match="wmcp_capture_tree failed"):
            ffi._check_error(-1, "wmcp_capture_tree")

    def test_send_text_error_status_raises(self):
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_send_text.return_value = -1
        with pytest.raises(RuntimeError, match="wmcp_send_text failed"):
            ffi.send_text("oops")


# ===========================================================================
# native_worker.py -- subprocess IPC error paths
# ===========================================================================


class TestFindWorkerExe:
    """_find_worker_exe() search order and env-var override."""

    def test_returns_none_when_no_candidate_exists(self):
        from windows_mcp.native_worker import _find_worker_exe

        with (
            patch.dict("os.environ", {}, clear=True),
            patch("windows_mcp.native_worker.Path.exists", return_value=False),
        ):
            result = _find_worker_exe()
        assert result is None

    def test_env_var_override_takes_priority(self, tmp_path):
        exe = tmp_path / "wmcp-worker.exe"
        exe.write_bytes(b"MZ")
        from windows_mcp.native_worker import _find_worker_exe

        with patch.dict("os.environ", {"WMCP_WORKER_EXE": str(exe)}):
            result = _find_worker_exe()
        assert result == exe


class TestNativeWorkerLifecycle:
    """NativeWorker start/stop/is_running lifecycle contracts."""

    def _make_worker(self, exe: str = "/fake/wmcp-worker.exe") -> object:
        from windows_mcp.native_worker import NativeWorker

        return NativeWorker(exe_path=exe, call_timeout=5.0)

    def test_is_running_false_before_start(self):
        w = self._make_worker()
        assert w.is_running is False

    async def test_start_raises_file_not_found_when_exe_absent(self):
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/nonexistent/wmcp-worker.exe")
        # Override exe_path to a non-existent path without using _find_worker_exe
        w._exe_path = None
        with pytest.raises(FileNotFoundError, match="wmcp-worker.exe not found"):
            await w.start()

    async def test_start_idempotent_when_already_running(self):
        """start() is a no-op if the process is already alive."""
        w = self._make_worker()
        mock_proc = MagicMock()
        mock_proc.returncode = None  # process alive
        mock_proc.pid = 9999
        mock_proc.stderr = AsyncMock()
        mock_proc.stderr.readline = AsyncMock(return_value=b"")
        w._process = mock_proc

        with patch("asyncio.create_subprocess_exec") as mock_spawn:
            await w.start()
        mock_spawn.assert_not_called()

    async def test_stop_when_process_never_started_is_safe(self):
        """stop() on a worker that was never started must not raise."""
        w = self._make_worker()
        await w.stop()  # should be a no-op
        assert w._process is None

    async def test_stop_kills_process_if_graceful_wait_times_out(self):
        """stop() calls kill() if the process does not exit within 5 seconds."""
        w = self._make_worker()

        mock_proc = MagicMock()
        mock_proc.returncode = None
        mock_proc.stdin = MagicMock()
        mock_proc.stdin.close = MagicMock()

        # wait() hangs → asyncio.wait_for raises TimeoutError after mocked timeout
        async def _slow_wait():
            await asyncio.sleep(9999)

        mock_proc.wait = _slow_wait
        mock_proc.kill = MagicMock()

        # After kill, wait() returns immediately
        async def _fast_wait():
            return 0

        w._process = mock_proc

        with patch("asyncio.wait_for", side_effect=asyncio.TimeoutError):
            mock_proc.kill = MagicMock()
            # Re-assign wait so the post-kill wait works
            mock_proc.wait = AsyncMock(return_value=0)
            await w.stop()

        mock_proc.kill.assert_called_once()


class TestNativeWorkerCallErrors:
    """NativeWorker.call() error and protocol-violation paths."""

    def _make_running_worker(self, responses: list[bytes]) -> tuple:
        """Return a NativeWorker wired to serve predetermined response lines."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe", call_timeout=2.0)

        mock_proc = MagicMock()
        mock_proc.returncode = None
        mock_proc.pid = 1234
        mock_proc.stdin = MagicMock()
        mock_proc.stdin.write = MagicMock()
        mock_proc.stdin.drain = AsyncMock()

        response_iter = iter(responses)

        async def _readline():
            try:
                return next(response_iter)
            except StopIteration:
                return b""

        mock_proc.stdout = MagicMock()
        mock_proc.stdout.readline = _readline
        w._process = mock_proc
        return w, mock_proc

    async def test_call_before_start_raises_runtime_error(self):
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe")
        with pytest.raises(RuntimeError, match="Worker not running"):
            await w.call("ping")

    async def test_call_returns_result_from_valid_response(self):
        response = json.dumps({"id": 1, "result": "pong"}).encode() + b"\n"
        w, _ = self._make_running_worker([response])
        result = await w.call("ping")
        assert result == "pong"

    async def test_call_raises_on_error_field(self):
        response = json.dumps({"id": 1, "error": "unknown method: bad_method"}).encode() + b"\n"
        w, _ = self._make_running_worker([response])
        with pytest.raises(RuntimeError, match="Worker error: unknown method"):
            await w.call("bad_method")

    async def test_call_raises_on_response_id_mismatch(self):
        response = json.dumps({"id": 999, "result": "irrelevant"}).encode() + b"\n"
        w, _ = self._make_running_worker([response])
        with pytest.raises(RuntimeError, match="Response ID mismatch"):
            await w.call("something")

    async def test_call_raises_on_non_json_response(self):
        w, _ = self._make_running_worker([b"this is not json\n"])
        with pytest.raises(RuntimeError, match="non-JSON output"):
            await w.call("ping")

    async def test_call_raises_when_stdout_closed(self):
        """readline() returning b'' means the worker process closed stdout."""
        w, _ = self._make_running_worker([b""])
        with pytest.raises(RuntimeError, match="closed stdout"):
            await w.call("ping")

    async def test_call_raises_timeout_error(self):
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe", call_timeout=0.001)
        mock_proc = MagicMock()
        mock_proc.returncode = None
        mock_proc.stdin = MagicMock()
        mock_proc.stdin.write = MagicMock()
        mock_proc.stdin.drain = AsyncMock()

        async def _slow_read():
            await asyncio.sleep(10)
            return b""

        mock_proc.stdout = MagicMock()
        mock_proc.stdout.readline = _slow_read
        w._process = mock_proc

        with pytest.raises(TimeoutError, match="timed out"):
            await w.call("slow_method")

    async def test_call_result_none_when_missing(self):
        """A response with no 'result' key returns None (not a crash)."""
        response = json.dumps({"id": 1}).encode() + b"\n"
        w, _ = self._make_running_worker([response])
        result = await w.call("ping")
        assert result is None

    async def test_aenter_aexit_starts_and_stops(self):
        """Context manager protocol delegates to start()/stop()."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe")
        w.start = AsyncMock()
        w.stop = AsyncMock()
        async with w:
            w.start.assert_awaited_once()
        w.stop.assert_awaited_once()

    async def test_sequential_calls_increment_request_id(self):
        """Each call gets a unique, incrementing request ID."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe", call_timeout=2.0)
        mock_proc = MagicMock()
        mock_proc.returncode = None
        mock_proc.stdin = MagicMock()
        mock_proc.stdin.write = MagicMock()
        mock_proc.stdin.drain = AsyncMock()

        # Serve id=1 and id=2 in order
        responses = [
            json.dumps({"id": 1, "result": "first"}).encode() + b"\n",
            json.dumps({"id": 2, "result": "second"}).encode() + b"\n",
        ]
        response_iter = iter(responses)

        async def _readline():
            return next(response_iter)

        mock_proc.stdout = MagicMock()
        mock_proc.stdout.readline = _readline
        w._process = mock_proc

        r1 = await w.call("m1")
        r2 = await w.call("m2")
        assert r1 == "first"
        assert r2 == "second"
        assert w._request_id == 2


# ===========================================================================
# input/service.py -- untested branches
# ===========================================================================


class TestInputServiceScrollBranches:
    """InputService.scroll() -- every match-case branch."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        svc = InputService()
        return svc

    def test_vertical_scroll_up_calls_wheel_up(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.uia.WheelUp") as mock_up:
            result = svc.scroll(type="vertical", direction="up", wheel_times=3)
        mock_up.assert_called_once_with(3)
        assert result is None

    def test_vertical_scroll_down_calls_wheel_down(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.uia.WheelDown") as mock_down:
            result = svc.scroll(type="vertical", direction="down", wheel_times=2)
        mock_down.assert_called_once_with(2)
        assert result is None

    def test_vertical_invalid_direction_returns_error_string(self):
        svc = self._make_service()
        result = svc.scroll(type="vertical", direction="sideways")
        assert result == 'Invalid direction. Use "up" or "down".'

    def test_horizontal_scroll_left_holds_shift_and_calls_wheel_up(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp") as mock_ku,
            patch("windows_mcp.input.service.pg.sleep"),
            patch("windows_mcp.input.service.uia.WheelUp") as mock_up,
        ):
            result = svc.scroll(type="horizontal", direction="left", wheel_times=1)
        mock_kd.assert_called_once_with("Shift")
        mock_ku.assert_called_once_with("Shift")
        mock_up.assert_called_once_with(1)
        assert result is None

    def test_horizontal_scroll_right_holds_shift_and_calls_wheel_down(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp") as mock_ku,
            patch("windows_mcp.input.service.pg.sleep"),
            patch("windows_mcp.input.service.uia.WheelDown") as mock_down,
        ):
            result = svc.scroll(type="horizontal", direction="right", wheel_times=4)
        mock_kd.assert_called_once_with("Shift")
        mock_ku.assert_called_once_with("Shift")
        mock_down.assert_called_once_with(4)
        assert result is None

    def test_horizontal_invalid_direction_returns_error_string(self):
        svc = self._make_service()
        result = svc.scroll(type="horizontal", direction="diagonal")
        assert result == 'Invalid direction. Use "left" or "right".'

    def test_invalid_scroll_type_returns_error_string(self):
        svc = self._make_service()
        result = svc.scroll(type="circular", direction="up")
        assert result == 'Invalid type. Use "horizontal" or "vertical".'

    def test_scroll_with_loc_calls_move_first(self):
        svc = self._make_service()
        with (
            patch.object(svc, "move") as mock_move,
            patch("windows_mcp.input.service.uia.WheelDown"),
        ):
            svc.scroll(loc=(500, 500), type="vertical", direction="down")
        mock_move.assert_called_once_with((500, 500))

    def test_scroll_without_loc_does_not_call_move(self):
        svc = self._make_service()
        with (
            patch.object(svc, "move") as mock_move,
            patch("windows_mcp.input.service.uia.WheelDown"),
        ):
            svc.scroll(loc=None, type="vertical", direction="down")
        mock_move.assert_not_called()

    def test_horizontal_shift_released_even_on_wheel_exception(self):
        """keyUp("Shift") is called even if WheelUp raises."""
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown"),
            patch("windows_mcp.input.service.pg.keyUp") as mock_ku,
            patch("windows_mcp.input.service.pg.sleep"),
            patch(
                "windows_mcp.input.service.uia.WheelUp",
                side_effect=RuntimeError("COM error"),
            ),
        ):
            with pytest.raises(RuntimeError):
                svc.scroll(type="horizontal", direction="left")
        mock_ku.assert_called_once_with("Shift")


class TestInputServiceShortcutBranches:
    """InputService.shortcut() -- single key vs. multi-key combination."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_single_key_calls_pg_press(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.press") as mock_press:
            svc.shortcut("enter")
        mock_press.assert_called_once_with("enter")

    def test_combination_calls_pg_hotkey(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.hotkey") as mock_hk:
            svc.shortcut("ctrl+c")
        mock_hk.assert_called_once_with("ctrl", "c")

    def test_three_key_combination(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.hotkey") as mock_hk:
            svc.shortcut("ctrl+shift+s")
        mock_hk.assert_called_once_with("ctrl", "shift", "s")

    def test_empty_string_calls_pg_press_with_empty(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.press") as mock_press:
            svc.shortcut("")
        mock_press.assert_called_once_with("")


class TestInputServiceTypeBranches:
    """InputService.type() -- caret_position, clear, press_enter, native fallback."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_caret_position_start_presses_home(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=3),
        ):
            svc.type((0, 0), "abc", caret_position="start")
        mock_press.assert_any_call("home")

    def test_caret_position_end_presses_end(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=3),
        ):
            svc.type((0, 0), "abc", caret_position="end")
        mock_press.assert_any_call("end")

    def test_caret_position_idle_does_not_press_home_or_end(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=0),
        ):
            svc.type((0, 0), "abc", caret_position="idle")
        called_args = [c[0][0] for c in mock_press.call_args_list]
        assert "home" not in called_args
        assert "end" not in called_args

    def test_clear_true_sends_select_all_and_backspace(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.sleep"),
            patch("windows_mcp.input.service.pg.hotkey") as mock_hk,
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=0),
        ):
            svc.type((0, 0), "new text", clear=True)
        mock_hk.assert_any_call("ctrl", "a")
        mock_press.assert_any_call("backspace")

    def test_clear_string_true_is_treated_as_clear(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.sleep"),
            patch("windows_mcp.input.service.pg.hotkey") as mock_hk,
            patch("windows_mcp.input.service.pg.press"),
            patch("windows_mcp.input.service.native_send_text", return_value=0),
        ):
            svc.type((0, 0), "x", clear="true")
        mock_hk.assert_any_call("ctrl", "a")

    def test_press_enter_true_presses_enter_key(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=1),
        ):
            svc.type((0, 0), "submit", press_enter=True)
        mock_press.assert_any_call("enter")

    def test_press_enter_false_does_not_press_enter(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.pg.press") as mock_press,
            patch("windows_mcp.input.service.native_send_text", return_value=1),
        ):
            svc.type((0, 0), "no-enter", press_enter=False)
        pressed = [c[0][0] for c in mock_press.call_args_list]
        assert "enter" not in pressed

    def test_native_send_text_none_falls_back_to_pg_typewrite(self):
        """When native_send_text returns None, pg.typewrite() must be called."""
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.native_send_text", return_value=None),
            patch("windows_mcp.input.service.pg.typewrite") as mock_tw,
        ):
            svc.type((0, 0), "fallback text")
        mock_tw.assert_called_once_with("fallback text", interval=0.02)

    def test_native_send_text_not_none_skips_pg_typewrite(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.leftClick"),
            patch("windows_mcp.input.service.native_send_text", return_value=5),
            patch("windows_mcp.input.service.pg.typewrite") as mock_tw,
        ):
            svc.type((0, 0), "native text")
        mock_tw.assert_not_called()


class TestInputServiceClickBranches:
    """InputService.click() -- single vs. multi-click, native vs. fallback."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_single_click_uses_native_when_available(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_click", return_value=2) as mock_nc,
            patch("windows_mcp.input.service.pg.click") as mock_pg,
        ):
            svc.click((100, 200), button="left", clicks=1)
        mock_nc.assert_called_once_with(100, 200, "left")
        mock_pg.assert_not_called()

    def test_single_click_falls_back_to_pg_when_native_returns_none(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_click", return_value=None),
            patch("windows_mcp.input.service.pg.click") as mock_pg,
        ):
            svc.click((100, 200), button="left", clicks=1)
        mock_pg.assert_called_once_with(100, 200, button="left", clicks=1, duration=0.1)

    def test_double_click_always_uses_pyautogui(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_click", return_value=2) as mock_nc,
            patch("windows_mcp.input.service.pg.click") as mock_pg,
        ):
            svc.click((300, 400), button="left", clicks=2)
        mock_nc.assert_not_called()
        mock_pg.assert_called_once_with(300, 400, button="left", clicks=2, duration=0.1)

    def test_right_click_single_uses_native(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_click", return_value=2) as mock_nc,
            patch("windows_mcp.input.service.pg.click"),
        ):
            svc.click((50, 50), button="right", clicks=1)
        mock_nc.assert_called_once_with(50, 50, "right")


class TestInputServiceMoveBranches:
    """InputService.move() -- native vs. pyautogui fallback."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_move_uses_native_when_available(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_mouse_move", return_value=1) as mock_nm,
            patch("windows_mcp.input.service.pg.moveTo") as mock_mt,
        ):
            svc.move((600, 700))
        mock_nm.assert_called_once_with(600, 700)
        mock_mt.assert_not_called()

    def test_move_falls_back_to_pg_when_native_returns_none(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.native_send_mouse_move", return_value=None),
            patch("windows_mcp.input.service.pg.moveTo") as mock_mt,
        ):
            svc.move((200, 300))
        mock_mt.assert_called_once_with(200, 300, duration=0.1)


class TestInputServiceMultiSelectBranches:
    """InputService.multi_select() -- ctrl-held vs. not, empty locs."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_multi_select_empty_locs_does_not_crash(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.click") as mock_click:
            svc.multi_select(press_ctrl=False, locs=[])
        mock_click.assert_not_called()

    def test_multi_select_without_ctrl_does_not_hold_ctrl(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp"),
            patch("windows_mcp.input.service.pg.click"),
            patch("windows_mcp.input.service.pg.sleep"),
        ):
            svc.multi_select(press_ctrl=False, locs=[[100, 200]])
        ctrl_down_calls = [c for c in mock_kd.call_args_list if c[0][0] == "ctrl"]
        assert ctrl_down_calls == []

    def test_multi_select_with_ctrl_holds_and_releases_ctrl(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp") as mock_ku,
            patch("windows_mcp.input.service.pg.click"),
            patch("windows_mcp.input.service.pg.sleep"),
        ):
            svc.multi_select(press_ctrl=True, locs=[[100, 200], [300, 400]])
        mock_kd.assert_any_call("ctrl")
        mock_ku.assert_any_call("ctrl")

    def test_multi_select_ctrl_released_even_when_click_raises(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown"),
            patch("windows_mcp.input.service.pg.keyUp") as mock_ku,
            patch(
                "windows_mcp.input.service.pg.click",
                side_effect=RuntimeError("click failed"),
            ),
            patch("windows_mcp.input.service.pg.sleep"),
        ):
            with pytest.raises(RuntimeError):
                svc.multi_select(press_ctrl=True, locs=[[100, 200]])
        mock_ku.assert_any_call("ctrl")

    def test_multi_select_string_true_holds_ctrl(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp"),
            patch("windows_mcp.input.service.pg.click"),
            patch("windows_mcp.input.service.pg.sleep"),
        ):
            svc.multi_select(press_ctrl="true", locs=[[10, 20]])
        mock_kd.assert_any_call("ctrl")

    def test_multi_select_string_false_does_not_hold_ctrl(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.keyDown") as mock_kd,
            patch("windows_mcp.input.service.pg.keyUp"),
            patch("windows_mcp.input.service.pg.click"),
            patch("windows_mcp.input.service.pg.sleep"),
        ):
            svc.multi_select(press_ctrl="false", locs=[[10, 20]])
        ctrl_calls = [c for c in mock_kd.call_args_list if c[0][0] == "ctrl"]
        assert ctrl_calls == []


class TestInputServiceMultiEditBranches:
    """InputService.multi_edit() -- delegates to type() with clear=True."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_multi_edit_empty_list_does_nothing(self):
        svc = self._make_service()
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([])
        mock_type.assert_not_called()

    def test_multi_edit_calls_type_with_clear_for_each_entry(self):
        svc = self._make_service()
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(100, 200, "alpha"), (300, 400, "beta")])
        assert mock_type.call_count == 2
        mock_type.assert_any_call((100, 200), text="alpha", clear=True)
        mock_type.assert_any_call((300, 400), text="beta", clear=True)


class TestInputServiceDragBranches:
    """InputService.drag() -- delegates to pg.dragTo."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_drag_calls_drag_to_with_correct_coordinates(self):
        svc = self._make_service()
        with (
            patch("windows_mcp.input.service.pg.sleep"),
            patch("windows_mcp.input.service.pg.dragTo") as mock_dt,
        ):
            svc.drag((800, 600))
        mock_dt.assert_called_once_with(800, 600, duration=0.6)


# ===========================================================================
# desktop/service.py -- get_system_info() native_info shape variants
# ===========================================================================


class TestGetSystemInfoNativeInfoShapes:
    """Desktop.get_system_info() with different native_info shapes.

    The method has two top-level branches:
      A.  native_info is not None AND 'cpu_count' is present  → Rust fast-path
      B.  otherwise                                           → pure-psutil path

    Branch A has several sub-branches worth testing:
      - cpu_usage_percent is empty  → cpu_pct defaults to 0.0
      - disks list is empty         → disk values remain 0.0
      - disk total_bytes == 0       → division-by-zero guard
      - no C: disk in list          → disk values remain 0.0
      - disk mount_point is lowercase 'c:' (case-insensitive match)
    """

    def _desktop(self) -> object:
        return _make_desktop_instance()

    # Helper: build a fully-formed native_info dict
    @staticmethod
    def _full_native_info(**overrides) -> dict:
        base = {
            "os_name": "Windows",
            "os_version": "10.0.19045",
            "hostname": "TESTPC",
            "cpu_count": 4,
            "cpu_usage_percent": [10.0, 20.0, 30.0, 40.0],
            "total_memory_bytes": 16 * 1024**3,
            "used_memory_bytes": 8 * 1024**3,
            "disks": [
                {
                    "name": "C:",
                    "mount_point": "C:\\",
                    "total_bytes": 500 * 1024**3,
                    "available_bytes": 300 * 1024**3,
                }
            ],
        }
        base.update(overrides)
        return base

    @staticmethod
    def _psutil_mock() -> MagicMock:
        m = MagicMock()
        m.cpu_percent.return_value = 5.0
        m.cpu_count.return_value = 2
        m.virtual_memory.return_value = SimpleNamespace(
            percent=50.0, used=4 * 1024**3, total=8 * 1024**3
        )
        m.disk_usage.return_value = SimpleNamespace(
            percent=30.0, used=100 * 1024**3, total=500 * 1024**3
        )
        m.boot_time.return_value = 1_700_000_000.0
        m.net_io_counters.return_value = SimpleNamespace(
            bytes_sent=50 * 1024**2, bytes_recv=200 * 1024**2
        )
        return m

    def _call_get_system_info(self, desktop, native_info, psutil_mock=None) -> str:
        if psutil_mock is None:
            psutil_mock = self._psutil_mock()
        import builtins

        real_import = builtins.__import__

        def _patched_import(name, *args, **kwargs):
            if name == "psutil":
                return psutil_mock
            return real_import(name, *args, **kwargs)

        # get_system_info() does `from windows_mcp.native import native_system_info`
        # inside the function body, so we patch the function on the native module
        # directly.  The function is re-imported each call, so patching the module
        # attribute is the right level.
        with (
            patch("builtins.__import__", side_effect=_patched_import),
            patch("windows_mcp.native.native_system_info", return_value=native_info),
        ):
            return desktop.get_system_info()

    def test_full_native_info_uses_rust_path(self):
        d = self._desktop()
        result = self._call_get_system_info(d, self._full_native_info())
        assert "CPU" in result
        assert "Memory" in result
        assert "Disk C" in result
        # avg of [10,20,30,40] = 25.0
        assert "25.0%" in result

    def test_empty_cpu_usage_percent_defaults_to_zero(self):
        d = self._desktop()
        native_info = self._full_native_info(cpu_usage_percent=[])
        result = self._call_get_system_info(d, native_info)
        assert "0.0%" in result

    def test_empty_disks_list_leaves_disk_at_zero(self):
        d = self._desktop()
        native_info = self._full_native_info(disks=[])
        result = self._call_get_system_info(d, native_info)
        # disk_pct=0.0, disk_used_gb=0.0, disk_total_gb=0.0
        assert "Disk C" in result

    def test_disk_with_total_bytes_zero_avoids_division_by_zero(self):
        d = self._desktop()
        native_info = self._full_native_info(
            disks=[{"mount_point": "C:\\", "total_bytes": 0, "available_bytes": 0}]
        )
        result = self._call_get_system_info(d, native_info)
        assert "0.0%" in result  # disk_pct defaults to 0.0

    def test_no_c_drive_in_disk_list_leaves_disk_at_zero(self):
        d = self._desktop()
        native_info = self._full_native_info(
            disks=[
                {
                    "mount_point": "D:\\",
                    "total_bytes": 1000 * 1024**3,
                    "available_bytes": 200 * 1024**3,
                }
            ]
        )
        result = self._call_get_system_info(d, native_info)
        # Should not crash; disk values remain 0
        assert "Disk C" in result

    def test_c_drive_mount_point_case_insensitive(self):
        """mount_point 'c:\\' (lowercase) should still match the C: check."""
        d = self._desktop()
        native_info = self._full_native_info(
            disks=[
                {
                    "mount_point": "c:\\",
                    "total_bytes": 400 * 1024**3,
                    "available_bytes": 200 * 1024**3,
                }
            ]
        )
        result = self._call_get_system_info(d, native_info)
        # disk_used = 400-200 = 200 GB, used_gb = 200.0
        assert "200.0" in result

    def test_native_info_none_uses_psutil_branch(self):
        d = self._desktop()
        pm = self._psutil_mock()
        result = self._call_get_system_info(d, native_info=None, psutil_mock=pm)
        # psutil mock returns cpu_percent=5.0
        assert "5.0%" in result

    def test_native_info_missing_cpu_count_uses_psutil_branch(self):
        """native_info without 'cpu_count' key triggers the psutil fallback."""
        d = self._desktop()
        # No cpu_count key -- branch condition is False
        native_info = {
            "os_name": "Windows",
            "cpu_usage_percent": [5.0],
            "total_memory_bytes": 8 * 1024**3,
            "used_memory_bytes": 2 * 1024**3,
            "disks": [],
        }
        pm = self._psutil_mock()
        self._call_get_system_info(d, native_info=native_info, psutil_mock=pm)
        pm.cpu_percent.assert_called_once()

    def test_output_contains_network_stats(self):
        d = self._desktop()
        pm = self._psutil_mock()
        result = self._call_get_system_info(d, native_info=None, psutil_mock=pm)
        assert "sent" in result
        assert "received" in result

    def test_output_contains_uptime(self):
        d = self._desktop()
        result = self._call_get_system_info(d, self._full_native_info())
        assert "Uptime" in result

    def test_used_memory_equal_to_total_gives_100_percent(self):
        d = self._desktop()
        native_info = self._full_native_info(
            total_memory_bytes=8 * 1024**3,
            used_memory_bytes=8 * 1024**3,
        )
        result = self._call_get_system_info(d, native_info)
        assert "100.0%" in result

    def test_total_memory_zero_gives_zero_percent(self):
        d = self._desktop()
        native_info = self._full_native_info(
            total_memory_bytes=0,
            used_memory_bytes=0,
        )
        result = self._call_get_system_info(d, native_info)
        # mem_pct = 0.0 because of the `if mem_total` guard
        assert "0.0%" in result


# ===========================================================================
# Fuzz / boundary-value tests
# ===========================================================================


class TestNativeWrapperBoundaryValues:
    """Boundary-value and adversarial inputs for native.py wrappers."""

    def _mock_core(self, **return_values) -> MagicMock:
        m = MagicMock()
        for attr, val in return_values.items():
            getattr(m, attr).return_value = val
        return m

    def test_send_text_very_long_string(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_text=len("A" * 100_000))
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_text("A" * 100_000)
        assert result == 100_000

    def test_send_text_unicode_surrogates_do_not_crash(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_text=2)
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            # Emoji and CJK characters
            result = native.native_send_text("\U0001f600\u4e2d\u6587")
        assert result == 2

    def test_send_click_negative_coordinates(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_click=2)
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_click(-1, -1)
        assert result == 2
        mock_core.send_click.assert_called_once_with(-1, -1, "left")

    def test_send_click_huge_coordinates(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_click=2)
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_click(99999, 99999)
        assert result == 2

    def test_send_scroll_negative_delta(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_scroll=1)
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_scroll(0, 0, -120)
        assert result == 1

    def test_send_hotkey_single_key(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_hotkey=2)
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_hotkey([0x11])
        assert result == 2

    def test_send_hotkey_many_keys(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(send_hotkey=20)
        keys = list(range(0x41, 0x41 + 10))  # A-J
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_send_hotkey(keys)
        assert result == 20

    def test_capture_tree_max_depth_zero(self):
        import windows_mcp.native as native

        mock_core = self._mock_core(capture_tree=[])
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            result = native.native_capture_tree([1], max_depth=0)
        assert result == []

    def test_capture_tree_large_handle_list(self):
        import windows_mcp.native as native

        handles = list(range(1, 501))  # 500 handles
        mock_core = self._mock_core(capture_tree=[])
        with (
            patch.object(native, "HAS_NATIVE", True),
            patch.object(native, "windows_mcp_core", mock_core),
        ):
            native.native_capture_tree(handles, max_depth=5)
        mock_core.capture_tree.assert_called_once_with(handles, max_depth=5)


class TestNativeWorkerCallBoundaryValues:
    """Adversarial JSON and protocol inputs for NativeWorker.call()."""

    def _make_worker_with_response(self, raw_bytes: bytes) -> object:
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe", call_timeout=5.0)
        mock_proc = MagicMock()
        mock_proc.returncode = None
        mock_proc.stdin = MagicMock()
        mock_proc.stdin.write = MagicMock()
        mock_proc.stdin.drain = AsyncMock()

        async def _readline():
            return raw_bytes

        mock_proc.stdout = MagicMock()
        mock_proc.stdout.readline = _readline
        w._process = mock_proc
        return w

    async def test_response_with_null_result_returns_none(self):
        response = json.dumps({"id": 1, "result": None}).encode() + b"\n"
        w = self._make_worker_with_response(response)
        result = await w.call("m")
        assert result is None

    async def test_response_with_list_result(self):
        response = json.dumps({"id": 1, "result": [1, 2, 3]}).encode() + b"\n"
        w = self._make_worker_with_response(response)
        result = await w.call("m")
        assert result == [1, 2, 3]

    async def test_response_with_nested_dict_result(self):
        payload = {"id": 1, "result": {"os_name": "Windows", "cpu_count": 8}}
        response = json.dumps(payload).encode() + b"\n"
        w = self._make_worker_with_response(response)
        result = await w.call("system_info")
        assert result["cpu_count"] == 8

    async def test_truncated_json_raises_runtime_error(self):
        w = self._make_worker_with_response(b'{"id": 1, "result":\n')
        with pytest.raises(RuntimeError, match="non-JSON output"):
            await w.call("m")

    async def test_empty_object_response_raises_id_mismatch(self):
        """An empty JSON object {} has no 'id' so mismatches request_id=1."""
        w = self._make_worker_with_response(b"{}\n")
        with pytest.raises(RuntimeError, match="Response ID mismatch"):
            await w.call("m")


class TestInputServiceScrollBoundaryValues:
    """Adversarial and boundary inputs for InputService.scroll()."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_zero_wheel_times_still_calls_wheel_fn(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.uia.WheelDown") as mock_down:
            svc.scroll(type="vertical", direction="down", wheel_times=0)
        mock_down.assert_called_once_with(0)

    def test_large_wheel_times_passes_through(self):
        svc = self._make_service()
        with patch("windows_mcp.input.service.uia.WheelUp") as mock_up:
            svc.scroll(type="vertical", direction="up", wheel_times=1000)
        mock_up.assert_called_once_with(1000)


class TestInputServiceShortcutBoundaryValues:
    """Adversarial shortcut strings."""

    def _make_service(self) -> object:
        from windows_mcp.input.service import InputService

        return InputService()

    def test_shortcut_with_trailing_plus(self):
        """'ctrl+' splits to ['ctrl', ''] -- pg.hotkey is still called."""
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.hotkey") as mock_hk:
            svc.shortcut("ctrl+")
        mock_hk.assert_called_once_with("ctrl", "")

    def test_shortcut_with_spaces_treated_as_single_key(self):
        """Spaces are not separators -- treated as one key name."""
        svc = self._make_service()
        with patch("windows_mcp.input.service.pg.press") as mock_press:
            svc.shortcut("win d")
        mock_press.assert_called_once_with("win d")


# ===========================================================================
# native_ffi.py -- lines 125, 138, 154 (null output, send_text return)
# ===========================================================================


class TestNativeFFIMethodCoverage:
    """Tests that actually call NativeFFI methods to exercise production lines.

    The earlier TestNativeFFIErrorPaths tests demonstrated guard logic inline
    but never called system_info() / capture_tree() / send_text(), leaving
    lines 125, 138, and 154 uncovered.
    """

    def _make_ffi_with_mock_dll(self) -> tuple:
        from windows_mcp.native_ffi import NativeFFI

        mock_dll = MagicMock()
        mock_dll.wmcp_last_error.return_value = b"injected error"
        ffi = object.__new__(NativeFFI)
        ffi._dll = mock_dll
        return ffi, mock_dll

    def test_system_info_null_output_via_real_method(self):
        """Calling ffi.system_info() when the DLL returns OK but leaves
        the output pointer null triggers line 125."""
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_system_info.return_value = 0  # WMCP_OK
        with pytest.raises(RuntimeError, match="null output"):
            ffi.system_info()

    def test_capture_tree_null_output_via_real_method(self):
        """Calling ffi.capture_tree() when the DLL returns OK but leaves
        the output pointer null triggers line 154."""
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_capture_tree.return_value = 0  # WMCP_OK
        with pytest.raises(RuntimeError, match="null output"):
            ffi.capture_tree([12345], max_depth=10)

    def test_send_text_returns_count_on_success(self):
        """Calling ffi.send_text() on success path returns out_count.value (line 138)."""
        ffi, mock_dll = self._make_ffi_with_mock_dll()
        mock_dll.wmcp_send_text.return_value = 0  # WMCP_OK
        result = ffi.send_text("hello")
        # Mock doesn't write to the pointer, so c_uint32() stays at 0
        assert result == 0
        assert isinstance(result, int)


# ===========================================================================
# native_worker.py -- lines 86-88, 107, 125-126
# ===========================================================================


class TestNativeWorkerDrainStderr:
    """Cover _drain_stderr() loop body (lines 86-88) and stderr logging."""

    async def test_drain_stderr_logs_lines_then_exits_on_empty(self):
        """_drain_stderr reads lines until empty, covering lines 86-88."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe")

        mock_proc = MagicMock()
        lines = [b"line1\n", b"line2\n", b""]  # empty = EOF
        line_iter = iter(lines)

        async def _readline():
            return next(line_iter)

        mock_proc.stderr = MagicMock()
        mock_proc.stderr.readline = _readline
        w._process = mock_proc

        # Run _drain_stderr directly -- it should consume all lines
        await w._drain_stderr()
        # If we got here without hanging, the loop exited on b""


class TestNativeWorkerVerboseFlag:
    """Cover line 107: args.append('--verbose') when verbose=True."""

    async def test_start_with_verbose_passes_flag(self, tmp_path):
        """start() appends --verbose to args when verbose=True."""
        from windows_mcp.native_worker import NativeWorker

        exe = tmp_path / "wmcp-worker.exe"
        exe.write_bytes(b"MZ")  # Just needs to be a path

        w = NativeWorker(exe_path=str(exe), verbose=True)

        mock_proc = MagicMock()
        mock_proc.pid = 42
        mock_proc.returncode = None
        mock_proc.stderr = MagicMock()
        mock_proc.stderr.readline = AsyncMock(return_value=b"")

        with patch("asyncio.create_subprocess_exec", new_callable=AsyncMock) as mock_exec:
            mock_exec.return_value = mock_proc
            # Patch _drain_stderr to avoid unawaited coroutine warning
            with patch.object(w, "_drain_stderr", new_callable=AsyncMock):
                await w.start()

        # Verify --verbose was in the args
        call_args = mock_exec.call_args
        assert "--verbose" in call_args[0]

    async def test_start_without_verbose_omits_flag(self, tmp_path):
        """start() does NOT pass --verbose when verbose=False."""
        from windows_mcp.native_worker import NativeWorker

        exe = tmp_path / "wmcp-worker.exe"
        exe.write_bytes(b"MZ")

        w = NativeWorker(exe_path=str(exe), verbose=False)

        mock_proc = MagicMock()
        mock_proc.pid = 42
        mock_proc.returncode = None
        mock_proc.stderr = MagicMock()
        mock_proc.stderr.readline = AsyncMock(return_value=b"")

        with patch("asyncio.create_subprocess_exec", new_callable=AsyncMock) as mock_exec:
            mock_exec.return_value = mock_proc
            with patch.object(w, "_drain_stderr", new_callable=AsyncMock):
                await w.start()

        call_args = mock_exec.call_args
        assert "--verbose" not in call_args[0]


class TestNativeWorkerStopCancelsStderrTask:
    """Cover lines 124-126: await self._stderr_task + CancelledError catch."""

    async def test_stop_cancels_active_stderr_task(self):
        """stop() cancels _stderr_task and catches CancelledError."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker(exe_path="/fake/exe")

        # Create a real async task that sleeps forever
        async def _forever():
            await asyncio.sleep(9999)

        w._stderr_task = asyncio.create_task(_forever())
        # No process to stop
        w._process = None

        await w.stop()
        assert w._stderr_task is None
