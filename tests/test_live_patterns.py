"""Live-desktop integration tests for native UIA pattern invocation.

These tests require a running Windows desktop with UIAutomation available.
They are excluded from CI by default.  Run manually with:

    uv run python -m pytest tests/test_live_patterns.py -m live_desktop
"""

import sys

import pytest

pytestmark = [
    pytest.mark.live_desktop,
    pytest.mark.skipif(sys.platform != "win32", reason="Requires Windows"),
]


def _skip_if_no_native():
    from windows_mcp.native import HAS_NATIVE

    if not HAS_NATIVE:
        pytest.skip("Native extension unavailable")


@pytest.mark.live_desktop
class TestLivePatternInvocation:
    """Tests that pattern invocation works on real UI elements."""

    def test_invoke_returns_result_dict(self):
        """Invoke at taskbar area should return a PatternResult dict."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics, native_invoke_at

        metrics = native_get_screen_metrics()
        assert metrics is not None

        # Query bottom-center (taskbar area)
        result = native_invoke_at(metrics["primary_width"] // 2, metrics["primary_height"] - 10)
        assert result is not None
        assert "element_name" in result
        assert "element_type" in result
        assert "action" in result
        assert result["action"] == "invoke"
        assert "success" in result
        assert "detail" in result

    def test_toggle_returns_result_dict(self):
        """Toggle should return a result with state info."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics, native_toggle_at

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_toggle_at(metrics["primary_width"] // 2, metrics["primary_height"] // 2)
        assert result is not None
        assert result["action"] == "toggle"
        assert isinstance(result["success"], bool)

    def test_set_value_returns_result_dict(self):
        """SetValue should return a result dict."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics, native_set_value_at

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_set_value_at(
            metrics["primary_width"] // 2, metrics["primary_height"] // 2, "test"
        )
        assert result is not None
        assert result["action"] == "set_value"
        assert isinstance(result["success"], bool)

    def test_expand_returns_result_dict(self):
        """Expand should return a result dict."""
        _skip_if_no_native()
        from windows_mcp.native import native_expand_at, native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_expand_at(metrics["primary_width"] // 2, metrics["primary_height"] - 10)
        assert result is not None
        assert result["action"] == "expand"

    def test_select_returns_result_dict(self):
        """Select should return a result dict."""
        _skip_if_no_native()
        from windows_mcp.native import native_get_screen_metrics, native_select_at

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_select_at(metrics["primary_width"] // 2, metrics["primary_height"] - 10)
        assert result is not None
        assert result["action"] == "select"

    def test_collapse_returns_result_dict(self):
        """Collapse should return a result dict."""
        _skip_if_no_native()
        from windows_mcp.native import native_collapse_at, native_get_screen_metrics

        metrics = native_get_screen_metrics()
        assert metrics is not None

        result = native_collapse_at(metrics["primary_width"] // 2, metrics["primary_height"] - 10)
        assert result is not None
        assert result["action"] == "collapse"
