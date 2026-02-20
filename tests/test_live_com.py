"""Live-desktop integration tests for native UIA query functions.

These tests require a running Windows desktop with UIAutomation available.
They are excluded from CI by default.  Run manually with:

    uv run python -m pytest tests/test_live_com.py -m live_desktop

Skip all tests in this module if:
- The native extension is not available
- Not running on Windows with a desktop session
"""

import os
import sys

import pytest

# Skip the entire module if not on Windows or no desktop
_is_windows = sys.platform == "win32"
_has_display = os.environ.get("SESSIONNAME") is not None or _is_windows

pytestmark = [
    pytest.mark.live_desktop,
    pytest.mark.skipif(not _is_windows, reason="Requires Windows"),
]


def _skip_if_no_native():
    """Skip test if native extension is unavailable."""
    from windows_mcp.native import HAS_NATIVE

    if not HAS_NATIVE:
        pytest.skip("Native extension unavailable")


@pytest.mark.live_desktop
class TestLiveElementQuery:
    """Tests against actual Windows desktop."""

    def test_element_from_point_desktop_area(self):
        """ElementFromPoint should return a valid dict for desktop coordinates."""
        _skip_if_no_native()
        from windows_mcp.native import native_element_from_point, native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None, "get_screen_metrics returned None"

        # Query center of screen (should find something)
        result = native_element_from_point(
            metrics["primary_width"] // 2, metrics["primary_height"] // 2
        )
        assert result is not None
        assert "name" in result
        assert "control_type" in result
        assert "bounding_rect" in result
        assert "supported_patterns" in result
        assert isinstance(result["supported_patterns"], list)

    def test_element_from_point_taskbar(self):
        """The Windows taskbar should be findable near bottom of screen."""
        _skip_if_no_native()
        from windows_mcp.native import native_element_from_point, native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_element_from_point(
            metrics["primary_width"] // 2, metrics["primary_height"] - 10
        )
        assert result is not None
        # Taskbar area should have some element
        assert result["control_type"] != ""

    def test_find_elements_by_control_type(self):
        """Should find at least one Button somewhere on the desktop."""
        _skip_if_no_native()
        from windows_mcp.native import native_find_elements

        results = native_find_elements(control_type="Button", limit=5)
        assert results is not None
        assert len(results) > 0
        assert results[0]["control_type"] == "Button"

    def test_find_elements_with_limit(self):
        """Limit should cap the number of results."""
        _skip_if_no_native()
        from windows_mcp.native import native_find_elements

        results = native_find_elements(control_type="Button", limit=2)
        assert results is not None
        assert len(results) <= 2

    def test_find_elements_no_match(self):
        """Should return empty list for non-existent automation ID."""
        _skip_if_no_native()
        from windows_mcp.native import native_find_elements

        results = native_find_elements(automation_id="__nonexistent_id_12345__", limit=5)
        assert results is not None
        assert len(results) == 0


@pytest.mark.live_desktop
class TestLiveScreenMetrics:
    """Tests for screen metrics on live desktop."""

    def test_screen_metrics_valid(self):
        """Screen dimensions should be positive."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None
        assert metrics["primary_width"] > 0
        assert metrics["primary_height"] > 0
        assert metrics["virtual_width"] >= metrics["primary_width"]
        assert metrics["virtual_height"] >= metrics["primary_height"]

    def test_screen_metrics_reasonable_values(self):
        """Screen dimensions should be within reasonable bounds."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None
        # Minimum 640x480, maximum 16K
        assert 640 <= metrics["primary_width"] <= 16384
        assert 480 <= metrics["primary_height"] <= 16384
