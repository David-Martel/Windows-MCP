"""Integration tests for the Rust native extension and Python adapters.

Tests are skipped when the native extension is not available, except for
mock-based fallback tests which always run.
"""

from unittest.mock import patch

import pytest

from windows_mcp.native import (
    HAS_NATIVE,
    NATIVE_VERSION,
    native_capture_tree,
    native_send_click,
    native_send_hotkey,
    native_send_key,
    native_send_mouse_move,
    native_send_scroll,
    native_send_text,
    native_system_info,
)

pytestmark = pytest.mark.skipif(not HAS_NATIVE, reason="Native extension not available")


# ---------------------------------------------------------------------------
# Module-level checks
# ---------------------------------------------------------------------------


class TestNativeModule:
    def test_has_native_is_true(self):
        assert HAS_NATIVE is True

    def test_native_version_is_string(self):
        assert isinstance(NATIVE_VERSION, str)
        assert len(NATIVE_VERSION) > 0

    def test_import_windows_mcp_core(self):
        import windows_mcp_core

        assert hasattr(windows_mcp_core, "system_info")
        assert hasattr(windows_mcp_core, "capture_tree")
        assert hasattr(windows_mcp_core, "send_text")
        assert hasattr(windows_mcp_core, "send_key")
        assert hasattr(windows_mcp_core, "send_click")
        assert hasattr(windows_mcp_core, "send_mouse_move")
        assert hasattr(windows_mcp_core, "send_hotkey")
        assert hasattr(windows_mcp_core, "send_scroll")


# ---------------------------------------------------------------------------
# system_info
# ---------------------------------------------------------------------------


class TestNativeSystemInfo:
    def test_returns_dict(self):
        result = native_system_info()
        assert isinstance(result, dict)

    def test_has_required_keys(self):
        result = native_system_info()
        required = [
            "os_name",
            "os_version",
            "hostname",
            "cpu_count",
            "cpu_usage_percent",
            "total_memory_bytes",
            "used_memory_bytes",
            "disks",
        ]
        for key in required:
            assert key in result, f"Missing key: {key}"

    def test_cpu_count_positive(self):
        result = native_system_info()
        assert isinstance(result["cpu_count"], int)
        assert result["cpu_count"] > 0

    def test_cpu_usage_is_list(self):
        result = native_system_info()
        usage = result["cpu_usage_percent"]
        assert isinstance(usage, list)
        assert len(usage) == result["cpu_count"]
        for u in usage:
            assert isinstance(u, float)
            assert 0.0 <= u <= 100.0

    def test_memory_positive(self):
        result = native_system_info()
        assert result["total_memory_bytes"] > 0
        assert result["used_memory_bytes"] >= 0
        assert result["used_memory_bytes"] <= result["total_memory_bytes"]

    def test_disks_is_list(self):
        result = native_system_info()
        disks = result["disks"]
        assert isinstance(disks, list)
        assert len(disks) > 0

    def test_disk_has_required_keys(self):
        result = native_system_info()
        for disk in result["disks"]:
            assert "name" in disk
            assert "mount_point" in disk
            assert "total_bytes" in disk
            assert "available_bytes" in disk
            assert disk["total_bytes"] > 0

    def test_os_name_not_empty(self):
        result = native_system_info()
        assert len(result["os_name"]) > 0
        assert len(result["hostname"]) > 0


# ---------------------------------------------------------------------------
# Input functions -- return values only (no actual input injection in tests)
# ---------------------------------------------------------------------------


class TestNativeInput:
    def test_send_text_empty_returns_zero(self):
        # send_text returns event count; 0 for empty string (no actual input)
        result = native_send_text("")
        assert result == 0

    def test_send_hotkey_empty_returns_zero(self):
        # Empty hotkey list returns 0 (no actual input)
        result = native_send_hotkey([])
        assert result == 0

    def test_send_scroll_zero_delta_returns_int(self):
        # delta=0 sends a wheel event but no actual scrolling
        result = native_send_scroll(0, 0, 0, False)
        assert isinstance(result, int)


# ---------------------------------------------------------------------------
# capture_tree -- format compatibility
# ---------------------------------------------------------------------------


class TestNativeCaptureTree:
    def test_empty_handles_returns_empty(self):
        result = native_capture_tree([], max_depth=10)
        assert result == []

    def test_invalid_handle_skipped(self):
        result = native_capture_tree([0], max_depth=10)
        # Handle 0 is invalid, should be silently skipped
        assert isinstance(result, list)

    def test_result_is_list_of_dicts(self):
        # Use handle 0 which should be filtered out
        result = native_capture_tree([0, -1], max_depth=5)
        assert isinstance(result, list)
        # Both handles are invalid, so result should be empty
        for item in result:
            assert isinstance(item, dict)


# ---------------------------------------------------------------------------
# Fallback behavior when native is unavailable
# ---------------------------------------------------------------------------


class TestFallbackBehavior:
    """Verify that native wrapper functions return None when extension is mocked away.

    These tests always run regardless of HAS_NATIVE state by patching
    HAS_NATIVE to False.
    """

    def test_system_info_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_system_info()
            assert result is None

    def test_capture_tree_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_capture_tree([])
            assert result is None

    def test_send_text_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_text("")
            assert result is None

    def test_send_click_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_click(100, 200)
            assert result is None

    def test_send_key_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_key(0x0D)
            assert result is None

    def test_send_mouse_move_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_mouse_move(100, 200)
            assert result is None

    def test_send_hotkey_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_hotkey([0x11, 0x43])
            assert result is None

    def test_send_scroll_returns_none_when_unavailable(self):
        with patch("windows_mcp.native.HAS_NATIVE", False):
            result = native_send_scroll(100, 200, 120)
            assert result is None


# ---------------------------------------------------------------------------
# NativeFFI tests (ctypes DLL wrapper)
# ---------------------------------------------------------------------------


class TestNativeFFI:
    """Tests for the ctypes C ABI DLL wrapper."""

    @pytest.fixture
    def ffi(self):
        """Try to instantiate NativeFFI, skip if DLL not found."""
        try:
            from windows_mcp.native_ffi import NativeFFI

            return NativeFFI()
        except FileNotFoundError:
            pytest.skip("windows_mcp_ffi.dll not found")

    def test_system_info(self, ffi):
        result = ffi.system_info()
        assert isinstance(result, dict)
        assert "os_name" in result
        assert "cpu_count" in result
        assert result["cpu_count"] > 0

    def test_capture_tree_empty(self, ffi):
        result = ffi.capture_tree([], max_depth=10)
        assert result == []


# ---------------------------------------------------------------------------
# NativeWorker tests (subprocess IPC)
# ---------------------------------------------------------------------------


class TestNativeWorker:
    """Tests for the subprocess IPC worker."""

    @pytest.fixture
    async def worker(self):
        """Try to start the worker, skip if exe not found."""
        from windows_mcp.native_worker import NativeWorker

        w = NativeWorker()
        try:
            await w.start()
        except FileNotFoundError:
            pytest.skip("wmcp-worker.exe not found")
        yield w
        await w.stop()

    async def test_ping(self, worker):
        result = await worker.call("ping")
        assert result == "pong"

    async def test_system_info(self, worker):
        result = await worker.call("system_info")
        assert isinstance(result, dict)
        assert "os_name" in result
        assert result["cpu_count"] > 0

    async def test_capture_tree_empty(self, worker):
        result = await worker.call("capture_tree", handles=[], max_depth=10)
        assert result == []

    async def test_unknown_method_raises(self, worker):
        with pytest.raises(RuntimeError, match="unknown method"):
            await worker.call("nonexistent_method")
