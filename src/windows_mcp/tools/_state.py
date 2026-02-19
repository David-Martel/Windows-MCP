"""Shared mutable state for MCP tool handlers.

Globals are initialised by the FastMCP lifespan in ``__main__.py``.
Tool modules import this module and access attributes at call-time,
so they always see the current (post-lifespan) values.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from windows_mcp.analytics import PostHogAnalytics
    from windows_mcp.desktop.service import Desktop, Size

desktop: Desktop | None = None
analytics: PostHogAnalytics | None = None
screen_size: Size | None = None
