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


class TestNativeModuleStructure:
    """Verify the module exports expected symbols."""

    def test_has_system_info(self):
        assert hasattr(windows_mcp_core, "system_info")
        assert callable(windows_mcp_core.system_info)
