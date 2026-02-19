"""Observation / state-query MCP tools.

Registers: Snapshot, WaitFor, WaitForEvent, Find, VisionAnalyze (5 tools).
"""

import asyncio
import io
import logging
import threading
import time
from textwrap import dedent
from typing import Literal

from fastmcp import Context
from fastmcp.utilities.types import Image
from mcp.types import ToolAnnotations

from windows_mcp.analytics import with_analytics
from windows_mcp.tools import _state
from windows_mcp.tools._helpers import MAX_IMAGE_HEIGHT, MAX_IMAGE_WIDTH, _coerce_bool

logger = logging.getLogger("windows_mcp")


def register(mcp):  # noqa: C901
    """Register observation/state tools on *mcp*."""

    @mcp.tool(
        name="Snapshot",
        description="Captures complete desktop state including: system language, focused/opened windows, interactive elements (buttons, text fields, links, menus with coordinates), and scrollable areas. Set use_vision=True to include screenshot. Set use_dom=True for browser content to get web page elements instead of browser UI. Always call this first to understand the current desktop state before taking actions.",
        annotations=ToolAnnotations(
            title="Snapshot",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "State-Tool")
    def state_tool(
        use_vision: bool | str = False, use_dom: bool | str = False, ctx: Context = None
    ):
        try:
            use_vision = _coerce_bool(use_vision)
            use_dom = _coerce_bool(use_dom)

            # Calculate scale factor to cap resolution at 1080p (1920x1080)
            if _state.screen_size is not None:
                scale_width = (
                    MAX_IMAGE_WIDTH / _state.screen_size.width
                    if _state.screen_size.width > MAX_IMAGE_WIDTH
                    else 1.0
                )
                scale_height = (
                    MAX_IMAGE_HEIGHT / _state.screen_size.height
                    if _state.screen_size.height > MAX_IMAGE_HEIGHT
                    else 1.0
                )
                scale = min(scale_width, scale_height)
            else:
                scale = 1.0

            desktop_state = _state.desktop.get_state(
                use_vision=use_vision, use_dom=use_dom, as_bytes=False, scale=scale
            )

            interactive_elements = desktop_state.tree_state.interactive_elements_to_string()
            scrollable_elements = desktop_state.tree_state.scrollable_elements_to_string()
            windows = desktop_state.windows_to_string()
            active_window = desktop_state.active_window_to_string()
            active_desktop = desktop_state.active_desktop_to_string()
            all_desktops = desktop_state.desktops_to_string()

            # Convert screenshot to bytes for vision response
            screenshot_bytes = None
            if use_vision and desktop_state.screenshot is not None:
                buffered = io.BytesIO()
                desktop_state.screenshot.save(buffered, format="PNG")
                screenshot_bytes = buffered.getvalue()
                buffered.close()
        except Exception as e:
            return [f"Error capturing desktop state: {str(e)}. Please try again."]

        return [
            dedent(f"""
    Active Desktop:
    {active_desktop}

    All Desktops:
    {all_desktops}

    Focused Window:
    {active_window}

    Opened Windows:
    {windows}

    List of Interactive Elements:
    {interactive_elements or "No interactive elements found."}

    List of Scrollable Elements:
    {scrollable_elements or "No scrollable elements found."}""")
        ] + (
            [Image(data=screenshot_bytes, format="png")] if use_vision and screenshot_bytes else []
        )

    @mcp.tool(
        name="WaitFor",
        description=(
            "Waits for a window or UI element to appear within a timeout. "
            "mode='window': waits for a window whose title contains the given name. "
            "mode='element': waits for an interactive element matching the name in the "
            "accessibility tree. "
            "Returns the matched item or times out with an error. "
            "Use this instead of the Wait tool when you need to wait for a specific condition."
        ),
        annotations=ToolAnnotations(
            title="WaitFor",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "WaitFor-Tool")
    async def waitfor_tool(
        mode: Literal["window", "element"],
        name: str,
        timeout: int = 10,
        ctx: Context = None,
    ) -> str:
        # Cap timeout to prevent unbounded task lifetime
        timeout = min(max(timeout, 1), 300)
        poll_interval = 0.5
        deadline = time.monotonic() + timeout

        while time.monotonic() < deadline:
            try:
                # Run blocking get_state in thread pool to avoid blocking the event loop
                desktop_state = await asyncio.to_thread(
                    _state.desktop.get_state, use_vision=False, use_dom=False
                )
                if mode == "window":
                    for w in desktop_state.windows:
                        if name.lower() in w.name.lower():
                            return f"Window found: '{w.name}'"
                    if (
                        desktop_state.active_window
                        and name.lower() in desktop_state.active_window.name.lower()
                    ):
                        return f"Window found: '{desktop_state.active_window.name}'"
                elif mode == "element":
                    tree_state = desktop_state.tree_state
                    if tree_state:
                        for node in tree_state.interactive_nodes:
                            if name.lower() in node.name.lower():
                                center = node.center
                                return (
                                    f"Element found: '{node.name}' "
                                    f"({node.control_type}) at ({center.x},{center.y})"
                                )
            except Exception:
                logger.debug("WaitFor poll error", exc_info=True)
            await asyncio.sleep(poll_interval)

        return f"Timeout: '{name}' {mode} not found within {timeout}s."

    @mcp.tool(
        name="WaitForEvent",
        description=(
            "Subscribes to a Windows UI Automation event and waits for it to fire. "
            "Supported events: 'window_opened', 'window_closed', 'menu_opened', "
            "'menu_closed', 'text_changed', 'focus_changed'. "
            "Optionally filter by element name (title substring match). "
            "Returns the matched event details or times out. "
            "More efficient than polling with WaitFor for window lifecycle events."
        ),
        annotations=ToolAnnotations(
            title="WaitForEvent",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "WaitForEvent-Tool")
    async def wait_for_event_tool(
        event: Literal[
            "window_opened",
            "window_closed",
            "menu_opened",
            "menu_closed",
            "text_changed",
            "focus_changed",
        ],
        name: str = "",
        timeout: int = 30,
        ctx: Context = None,
    ) -> str:
        from windows_mcp.uia import Control
        from windows_mcp.uia.events import EventId

        _EVENT_MAP = {
            "window_opened": EventId.UIA_Window_WindowOpenedEventId,
            "window_closed": EventId.UIA_Window_WindowClosedEventId,
            "menu_opened": EventId.UIA_MenuOpenedEventId,
            "menu_closed": EventId.UIA_MenuClosedEventId,
            "text_changed": EventId.UIA_Text_TextChangedEventId,
            "focus_changed": EventId.UIA_AutomationFocusChangedEventId,
        }

        event_id = _EVENT_MAP.get(event)
        if event_id is None:
            return f"Error: unknown event '{event}'. Supported: {', '.join(_EVENT_MAP)}."

        timeout = min(max(timeout, 1), 300)
        matched = threading.Event()
        result_holder: list[str] = []

        def on_event(sender, fired_event_id):
            try:
                element = Control.CreateControlFromElement(sender)
                element_name = element.Name or ""
                if name and name.lower() not in element_name.lower():
                    return
                result_holder.append(
                    f"Event '{event}' fired: '{element_name}' "
                    f"({element.ControlTypeName})"
                )
                matched.set()
            except Exception:
                # If we can't read element details, still match if no name filter
                if not name:
                    result_holder.append(f"Event '{event}' fired (element details unavailable)")
                    matched.set()

        watchdog = _state.watchdog
        if watchdog is None:
            return "Error: WatchDog service not available."

        # Subscribe to the event via WatchDog.
        prev_callback = watchdog._automation_callback
        prev_event_id = watchdog._automation_event_id
        prev_element = watchdog._automation_element
        watchdog.set_automation_callback(on_event, event_id=event_id)

        try:
            deadline = time.monotonic() + timeout
            while not matched.is_set():
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                await asyncio.sleep(min(0.2, remaining))
        finally:
            # Restore previous callback (or clear).
            watchdog.set_automation_callback(prev_callback, prev_event_id, prev_element)

        if result_holder:
            return result_holder[0]
        return f"Timeout: no '{event}' event{f' matching {name!r}' if name else ''} within {timeout}s."

    @mcp.tool(
        name="Find",
        description=(
            "Searches for UI elements by name, control type, or window. "
            "Returns matching interactive elements with their coordinates. "
            "Useful when you know what you're looking for and only need matching results."
        ),
        annotations=ToolAnnotations(
            title="Find",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Find-Tool")
    def find_tool(
        name: str | None = None,
        control_type: str | None = None,
        window: str | None = None,
        limit: int = 20,
        ctx: Context = None,
    ) -> str:
        if not name and not control_type and not window:
            return "Error: at least one of name, control_type, or window must be specified."

        desktop_state = _state.desktop.get_state(use_vision=False, use_dom=False)
        tree_state = desktop_state.tree_state
        if not tree_state:
            return "Error: could not capture desktop state."

        matches = []
        for i, node in enumerate(tree_state.interactive_nodes):
            if name and name.lower() not in node.name.lower():
                continue
            if control_type and control_type.lower() not in node.control_type.lower():
                continue
            if window and window.lower() not in node.window_name.lower():
                continue
            center = node.center
            matches.append(
                f"[{i}] '{node.name}' ({node.control_type}) "
                f"at ({center.x},{center.y}) window='{node.window_name}'"
            )
            if len(matches) >= limit:
                break

        if not matches:
            criteria = []
            if name:
                criteria.append(f"name='{name}'")
            if control_type:
                criteria.append(f"type='{control_type}'")
            if window:
                criteria.append(f"window='{window}'")
            return f"No elements found matching {', '.join(criteria)}."

        return f"Found {len(matches)} element(s):\n" + "\n".join(matches)

    @mcp.tool(
        name="VisionAnalyze",
        description=(
            "Analyze the current screen using a vision-capable LLM. "
            "Takes a screenshot and sends it to the configured vision API for analysis. "
            "Modes: 'describe' returns a natural-language description, "
            "'elements' returns a JSON list of detected UI elements with coordinates, "
            "'query' answers a specific question about what's on screen. "
            "Requires VISION_API_URL and VISION_API_KEY environment variables. "
            "Useful when the accessibility tree is sparse or for custom-rendered UIs."
        ),
        annotations=ToolAnnotations(
            title="VisionAnalyze",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        ),
    )
    @with_analytics(lambda: _state.analytics, "VisionAnalyze-Tool")
    def vision_analyze(
        mode: Literal["describe", "elements", "query"] = "describe",
        query: str = "",
        target: str = "",
        ctx: Context = None,
    ) -> str:
        """Capture a screenshot and analyze it with a vision LLM."""
        from windows_mcp.vision import VisionService

        vision = VisionService()
        if not vision.is_configured:
            return (
                "Error: Vision API not configured. "
                "Set VISION_API_URL and VISION_API_KEY environment variables. "
                "Supports any OpenAI-compatible endpoint (GPT-4o, Claude via proxy, "
                "Ollama, llama.cpp server, or PC-AI pcai-inference)."
            )

        # Capture screenshot
        screenshot = _state.desktop.get_screenshot()
        img_buffer = io.BytesIO()
        screenshot.save(img_buffer, format="PNG")
        image_bytes = img_buffer.getvalue()

        if mode == "describe":
            return vision.describe_screen(image_bytes, context=query)
        elif mode == "elements":
            import json

            elements = vision.identify_elements(image_bytes, target=target)
            if not elements:
                return "No UI elements detected by vision analysis."
            return json.dumps(elements, indent=2)
        elif mode == "query":
            if not query:
                return "Error: query parameter required for mode='query'."
            return vision.analyze(image_bytes, prompt=query)
        else:
            return f"Error: unknown mode '{mode}'. Use 'describe', 'elements', or 'query'."
