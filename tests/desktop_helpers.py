"""Shared test helpers for Desktop service unit tests.

Provides factory functions for creating bare Desktop instances and
common test data objects without requiring COM/UIA initialization.
"""

import threading
from unittest.mock import MagicMock

from windows_mcp.desktop.views import DesktopState, Status, Window
from windows_mcp.tree.views import BoundingBox


def make_bare_desktop():
    """Return a Desktop instance that bypasses __init__ (no COM/UIA calls).

    Only the attributes accessed by the methods under test are set.
    All sub-services are MagicMock instances for easy stubbing.
    """
    from windows_mcp.desktop.service import Desktop

    d = Desktop.__new__(Desktop)
    d._state_lock = threading.Lock()
    d.desktop_state = None
    d._app_cache = None
    d._app_cache_time = 0.0
    d._APP_CACHE_TTL = 3600.0
    d._app_cache_lock = threading.Lock()
    # Sub-services - replaced per-test with MagicMock where needed
    d.tree = MagicMock()
    d._input = MagicMock()
    d._registry = MagicMock()
    d._shell = MagicMock()
    d._scraper = MagicMock()
    d._process = MagicMock()
    d._screen = MagicMock()
    d._window = MagicMock()
    return d


def make_window(name="Test App", status=Status.NORMAL, handle=1001, process_id=2001):
    """Create a Window dataclass with sensible defaults."""
    bbox = BoundingBox(left=0, top=0, right=800, bottom=600, width=800, height=600)
    return Window(
        name=name,
        is_browser=False,
        depth=0,
        status=status,
        bounding_box=bbox,
        handle=handle,
        process_id=process_id,
    )


def make_desktop_state(active_window=None, windows=None):
    """Create a DesktopState with minimal defaults."""
    return DesktopState(
        active_desktop={"id": "1", "name": "Desktop 1"},
        all_desktops=[{"id": "1", "name": "Desktop 1"}],
        active_window=active_window,
        windows=windows or [],
    )
