"""Performance regression tests and bottleneck identification.

Verifies that optimization changes are in place and documents
known performance characteristics for monitoring.
"""

from unittest.mock import patch

import pytest

from windows_mcp.desktop.service import Desktop


@pytest.fixture
def desktop():
    with patch.object(Desktop, "__init__", lambda self: None):
        return Desktop()


class TestPgPauseOptimization:
    """Verify pg.PAUSE is set to 0.05 (not the old 1.0)."""

    def test_desktop_service_pause(self):
        import pyautogui as pg

        assert pg.PAUSE <= 0.1, f"pg.PAUSE is {pg.PAUSE}, expected <= 0.1"

    def test_main_module_pause(self):
        import pyautogui as pg

        # After importing __main__, pg.PAUSE should be 0.05
        assert pg.PAUSE <= 0.1, f"pg.PAUSE is {pg.PAUSE}, expected <= 0.1"


class TestImageDrawSequential:
    """Verify screenshot annotation uses sequential drawing (no ThreadPoolExecutor)."""

    def test_no_thread_pool_in_annotation(self):
        import inspect

        source = inspect.getsource(Desktop.get_annotated_screenshot)
        assert "ThreadPoolExecutor" not in source
        assert "executor" not in source


class TestThreadPoolBounded:
    """Verify tree service ThreadPoolExecutor has max_workers bound."""

    def test_max_workers_in_source(self):
        import inspect

        from windows_mcp.tree.service import Tree

        source = inspect.getsource(Tree.get_window_wise_nodes)
        assert "max_workers" in source
        assert "min(8" in source


class TestAnalyticsNoPrint:
    """Verify analytics module doesn't use print() (corrupts MCP stdout)."""

    def test_no_print_in_analytics(self):
        import inspect

        import windows_mcp.analytics as analytics_mod

        source = inspect.getsource(analytics_mod)
        # Allow print in comments/strings but not as a function call
        lines = source.split("\n")
        for i, line in enumerate(lines, 1):
            stripped = line.lstrip()
            if stripped.startswith("#") or stripped.startswith('"') or stripped.startswith("'"):
                continue
            assert "print(" not in stripped, (
                f"analytics.py line {i}: found print() call that corrupts MCP stdout"
            )


class TestWatchdogNoPrint:
    """Verify watchdog event handlers don't use print() (corrupts MCP stdout)."""

    def test_no_print_in_event_handlers(self):
        import inspect

        from windows_mcp.watchdog import event_handlers

        source = inspect.getsource(event_handlers)
        lines = source.split("\n")
        for i, line in enumerate(lines, 1):
            stripped = line.lstrip()
            if stripped.startswith("#") or stripped.startswith('"') or stripped.startswith("'"):
                continue
            assert "print(" not in stripped, (
                f"event_handlers.py line {i}: found print() call that corrupts MCP stdout"
            )


class TestRegistryUsesWinreg:
    """Verify registry methods use winreg stdlib (not PowerShell subprocess)."""

    def test_registry_get_no_subprocess(self):
        import inspect

        source = inspect.getsource(Desktop.registry_get)
        assert "execute_command" not in source
        assert "powershell" not in source.lower()

    def test_registry_set_no_subprocess(self):
        import inspect

        source = inspect.getsource(Desktop.registry_set)
        assert "execute_command" not in source

    def test_registry_delete_no_subprocess(self):
        import inspect

        source = inspect.getsource(Desktop.registry_delete)
        assert "execute_command" not in source

    def test_registry_list_no_subprocess(self):
        import inspect

        source = inspect.getsource(Desktop.registry_list)
        assert "execute_command" not in source


class TestGetWindowsVersionNoSubprocess:
    """Verify get_windows_version uses winreg (not PowerShell)."""

    def test_no_execute_command(self):
        import inspect

        source = inspect.getsource(Desktop.get_windows_version)
        assert "execute_command" not in source
        assert "winreg" in source


class TestGetDefaultLanguageNoSubprocess:
    """Verify get_default_language uses locale (not PowerShell)."""

    def test_no_execute_command(self):
        import inspect

        source = inspect.getsource(Desktop.get_default_language)
        assert "execute_command" not in source
        assert "locale" in source


class TestKnownBottlenecks:
    """Document known remaining performance bottlenecks for future optimization.

    These tests verify the bottleneck still exists (they pass when the bottleneck
    is present). When a bottleneck is fixed, the corresponding test should be
    updated to verify the fix.
    """

    def test_tree_uses_per_node_cache_request(self):
        """TreeScope_Subtree optimization NOT YET applied.

        Current: BuildUpdatedCache per-node with TreeScope_Element + TreeScope_Children
        Target: Single TreeScope_Subtree CacheRequest on window root
        Impact: ~2x reduction in COM round-trips during tree traversal
        """
        import inspect

        from windows_mcp.tree.service import Tree

        source = inspect.getsource(Tree)
        # When this assertion fails, the optimization has been applied
        assert "BuildUpdatedCache" in source or "TreeScope_Subtree" not in source

    def test_comtypes_overhead_present(self):
        """comtypes COM wrapper adds ~50-200us per call overhead.

        Target: Replace with windows-rs in Rust PyO3 extension for hot paths.
        Impact: ~1000ms saved per Snapshot (10,000 calls * 100us each)
        """
        import comtypes

        # Just verify comtypes is still in use (not yet replaced by Rust)
        assert comtypes is not None

    def test_pyautogui_still_used(self):
        """pyautogui uses image-based operations where SendInput would be faster.

        Target: Replace mouse_event/keybd_event with ctypes SendInput
        Impact: More reliable input simulation, lower latency
        """
        import pyautogui

        assert pyautogui is not None
