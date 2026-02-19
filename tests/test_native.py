"""Tests for the optional Rust PyO3 extension module (windows_mcp_core).

These tests are skipped if the native module is not built.
Build with: cd native && maturin develop
"""

import pytest

try:
    import windows_mcp_core

    HAS_NATIVE = True
except ImportError:
    HAS_NATIVE = False

pytestmark = pytest.mark.skipif(not HAS_NATIVE, reason="Native module not built")


class TestNativeSystemInfo:
    """Test the Rust system_info() function."""

    def test_returns_dict(self):
        result = windows_mcp_core.system_info()
        assert isinstance(result, dict)

    def test_has_os_fields(self):
        result = windows_mcp_core.system_info()
        assert "os_name" in result
        assert "os_version" in result
        assert "hostname" in result

    def test_has_cpu_fields(self):
        result = windows_mcp_core.system_info()
        assert "cpu_count" in result
        assert isinstance(result["cpu_count"], int)
        assert result["cpu_count"] > 0

    def test_has_memory_fields(self):
        result = windows_mcp_core.system_info()
        assert "total_memory_bytes" in result
        assert "used_memory_bytes" in result
        assert result["total_memory_bytes"] > 0
        assert result["used_memory_bytes"] >= 0

    def test_has_disks(self):
        result = windows_mcp_core.system_info()
        assert "disks" in result
        assert isinstance(result["disks"], list)
        if result["disks"]:
            disk = result["disks"][0]
            assert "mount_point" in disk
            assert "total_bytes" in disk


class TestNativeCaptureTree:
    """Test the Rust capture_tree() function."""

    def test_callable(self):
        assert hasattr(windows_mcp_core, "capture_tree")
        assert callable(windows_mcp_core.capture_tree)

    def test_empty_hwnd_list(self):
        result = windows_mcp_core.capture_tree([], max_depth=5)
        assert isinstance(result, list)
        assert len(result) == 0

    def test_invalid_hwnd_returns_empty(self):
        # HWND 0 is invalid -- should return empty list (no crash)
        result = windows_mcp_core.capture_tree([0], max_depth=5)
        assert isinstance(result, list)
        assert len(result) == 0

    def test_default_max_depth(self):
        # Should accept call without max_depth kwarg
        result = windows_mcp_core.capture_tree([])
        assert isinstance(result, list)

    def test_returns_dict_with_expected_keys(self):
        import ctypes

        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if hwnd == 0:
            pytest.skip("No foreground window available")
        try:
            result = windows_mcp_core.capture_tree([hwnd], max_depth=2)
        except (TypeError, OSError):
            # Window may have closed or become inaccessible between
            # GetForegroundWindow() and capture_tree()
            pytest.skip("Foreground window became inaccessible during test")
        assert isinstance(result, list)
        if not result:
            pytest.skip("Foreground window returned empty tree")
        elem = result[0]
        assert isinstance(elem, dict)
        for key in [
            "name",
            "automation_id",
            "control_type",
            "class_name",
            "bounding_rect",
            "children",
            "depth",
        ]:
            assert key in elem, f"Missing key: {key}"
        assert isinstance(elem["children"], list)
        assert isinstance(elem["bounding_rect"], list)
        assert len(elem["bounding_rect"]) == 4


class TestNativeModuleStructure:
    """Verify the module exports expected symbols."""

    def test_has_system_info(self):
        assert hasattr(windows_mcp_core, "system_info")
        assert callable(windows_mcp_core.system_info)

    def test_has_capture_tree(self):
        assert hasattr(windows_mcp_core, "capture_tree")
        assert callable(windows_mcp_core.capture_tree)

    def test_has_version(self):
        assert hasattr(windows_mcp_core, "__version__")
        assert isinstance(windows_mcp_core.__version__, str)
