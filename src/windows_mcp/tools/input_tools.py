"""Input / interaction MCP tools.

Registers: Click, Type, Scroll, Move, Shortcut, Wait,
MultiSelect, MultiEdit, Invoke (9 tools).
"""

from typing import Literal

import pyautogui as pg
from fastmcp import Context
from mcp.types import ToolAnnotations

from windows_mcp.analytics import with_analytics
from windows_mcp.tools import _state
from windows_mcp.tools._helpers import _coerce_bool, _validate_loc


def register(mcp):  # noqa: C901
    """Register input/interaction tools on *mcp*."""

    @mcp.tool(
        name="Click",
        description="Performs mouse clicks at specified coordinates [x, y]. Supports button types: 'left' for selection/activation, 'right' for context menus, 'middle'. Supports clicks: 0=hover only (no click), 1=single click (select/focus), 2=double click (open/activate).",
        annotations=ToolAnnotations(
            title="Click",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Click-Tool")
    def click_tool(
        loc: list[int],
        button: Literal["left", "right", "middle"] = "left",
        clicks: int = 1,
        ctx: Context = None,
    ) -> str:
        x, y = _validate_loc(loc)
        _state.desktop.click(loc=(x, y), button=button, clicks=clicks)
        num_clicks = {0: "Hover", 1: "Single", 2: "Double"}
        return f"{num_clicks.get(clicks, clicks)} {button} clicked at ({x},{y})."

    @mcp.tool(
        name="Type",
        description="Types text at specified coordinates [x, y]. Set clear=True to clear existing text first, False to append. Set press_enter=True to submit after typing. Set caret_position to 'start' (beginning), 'end' (end), or 'idle' (default).",
        annotations=ToolAnnotations(
            title="Type",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Type-Tool")
    def type_tool(
        loc: list[int],
        text: str,
        clear: bool | str = False,
        caret_position: Literal["start", "idle", "end"] = "idle",
        press_enter: bool | str = False,
        ctx: Context = None,
    ) -> str:
        clear = _coerce_bool(clear)
        press_enter = _coerce_bool(press_enter)
        x, y = _validate_loc(loc)
        _state.desktop.type(
            loc=(x, y),
            text=text,
            caret_position=caret_position,
            clear=clear,
            press_enter=press_enter,
        )
        return f"Typed {text} at ({x},{y})."

    @mcp.tool(
        name="Scroll",
        description="Scrolls at coordinates [x, y] or current mouse position if loc=None. Type: vertical (default) or horizontal. Direction: up/down for vertical, left/right for horizontal. wheel_times controls amount (1 wheel ≈ 3-5 lines). Use for navigating long content, lists, and web pages.",
        annotations=ToolAnnotations(
            title="Scroll",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Scroll-Tool")
    def scroll_tool(
        loc: list[int] = None,
        type: Literal["horizontal", "vertical"] = "vertical",
        direction: Literal["up", "down", "left", "right"] = "down",
        wheel_times: int = 1,
        ctx: Context = None,
    ) -> str:
        validated_loc = None
        if loc is not None:
            x, y = _validate_loc(loc)
            validated_loc = (x, y)
        response = _state.desktop.scroll(validated_loc, type, direction, wheel_times)
        if response:
            return response
        msg = f"Scrolled {type} {direction} by {wheel_times} wheel times"
        if validated_loc:
            msg += f" at ({validated_loc[0]},{validated_loc[1]})"
        return msg + "."

    @mcp.tool(
        name="Move",
        description="Moves mouse cursor to coordinates [x, y]. Set drag=True to perform a drag-and-drop operation from the current mouse position to the target coordinates. Default (drag=False) is a simple cursor move (hover).",
        annotations=ToolAnnotations(
            title="Move",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Move-Tool")
    def move_tool(loc: list[int], drag: bool | str = False, ctx: Context = None) -> str:
        drag = _coerce_bool(drag)
        x, y = _validate_loc(loc)
        if drag:
            _state.desktop.drag((x, y))
            return f"Dragged to ({x},{y})."
        else:
            _state.desktop.move((x, y))
            return f"Moved the mouse pointer to ({x},{y})."

    @mcp.tool(
        name="Shortcut",
        description='Executes keyboard shortcuts using key combinations separated by +. Examples: "ctrl+c" (copy), "ctrl+v" (paste), "alt+tab" (switch apps), "win+r" (Run dialog), "win" (Start menu), "ctrl+shift+esc" (Task Manager). Use for quick actions and system commands.',
        annotations=ToolAnnotations(
            title="Shortcut",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Shortcut-Tool")
    def shortcut_tool(shortcut: str, ctx: Context = None):
        if not shortcut or not shortcut.strip():
            return "Error: shortcut must not be empty."
        _state.desktop.shortcut(shortcut)
        return f"Pressed {shortcut}."

    @mcp.tool(
        name="Wait",
        description="Pauses execution for specified duration in seconds. Use when waiting for: applications to launch/load, UI animations to complete, page content to render, dialogs to appear, or between rapid actions. Helps ensure UI is ready before next interaction.",
        annotations=ToolAnnotations(
            title="Wait",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Wait-Tool")
    def wait_tool(duration: int, ctx: Context = None) -> str:
        duration = max(duration, 0)
        pg.sleep(duration)
        return f"Waited for {duration} seconds."

    @mcp.tool(
        name="MultiSelect",
        description="Selects multiple items such as files, folders, or checkboxes if press_ctrl=True, or performs multiple clicks if False.",
        annotations=ToolAnnotations(
            title="MultiSelect",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Multi-Select-Tool")
    def multi_select_tool(
        locs: list[list[int]], press_ctrl: bool | str = True, ctx: Context = None
    ) -> str:
        press_ctrl = _coerce_bool(press_ctrl, default=True)
        if not locs:
            return "Error: at least one location is required."
        validated = [_validate_loc(loc, label=f"locs[{i}]") for i, loc in enumerate(locs)]
        _state.desktop.multi_select(press_ctrl, validated)
        elements_str = "\n".join([f"({x},{y})" for x, y in validated])
        return f"Multi-selected elements at:\n{elements_str}"

    @mcp.tool(
        name="MultiEdit",
        description="Enters text into multiple input fields at specified coordinates [[x,y,text], ...].",
        annotations=ToolAnnotations(
            title="MultiEdit",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Multi-Edit-Tool")
    def multi_edit_tool(locs: list[list], ctx: Context = None) -> str:
        if not locs:
            return "Error: at least one location is required."
        validated = []
        for i, entry in enumerate(locs):
            if not isinstance(entry, (list, tuple)) or len(entry) < 3:
                raise ValueError(f"locs[{i}] must be [x, y, text], got {entry!r}")
            x, y = _validate_loc(entry[:2], label=f"locs[{i}]")
            validated.append((x, y, str(entry[2])))
        _state.desktop.multi_edit(validated)
        elements_str = ", ".join([f"({x},{y}) with text '{text}'" for x, y, text in validated])
        return f"Multi-edited elements at: {elements_str}"

    def _format_pattern_result(r: dict, x: int, y: int) -> str:
        """Format a Rust PatternResult dict into a user-facing string."""
        name = r.get("element_name") or "unnamed"
        etype = r.get("element_type") or "unknown"
        action = r.get("action", "")
        if not r.get("success"):
            return f"Error: element '{name}' does not support {r.get('detail', action)}."
        if action == "invoke":
            return f"Invoked '{name}' ({etype}) at ({x},{y})."
        if action == "toggle":
            state = r.get("detail", "")
            return f"Toggled '{name}' ({etype}) at ({x},{y}). {state}."
        if action == "set_value":
            detail = r.get("detail", "")
            return f"Set value on '{name}' ({etype}) at ({x},{y}). {detail}."
        if action == "expand":
            return f"Expanded '{name}' ({etype}) at ({x},{y})."
        if action == "collapse":
            return f"Collapsed '{name}' ({etype}) at ({x},{y})."
        if action == "select":
            return f"Selected '{name}' ({etype}) at ({x},{y})."
        return f"{action} on '{name}' ({etype}) at ({x},{y})."

    def _try_native_pattern(x: int, y: int, action: str, value: str | None) -> str | None:
        """Try Rust native UIA pattern invocation. Returns formatted string or None."""
        from windows_mcp.native import (
            native_collapse_at,
            native_expand_at,
            native_invoke_at,
            native_select_at,
            native_set_value_at,
            native_toggle_at,
        )

        if action == "set_value":
            if value is None:
                return "Error: value parameter required for set_value action."
            if len(value) > 10000:
                return f"Error: value too long ({len(value)} chars, max 10000)."
            result = native_set_value_at(x, y, value)
        else:
            fn = {
                "invoke": native_invoke_at,
                "toggle": native_toggle_at,
                "expand": native_expand_at,
                "collapse": native_collapse_at,
                "select": native_select_at,
            }.get(action)
            if fn is None:
                return None  # Unknown action, let fallback handle it
            result = fn(x, y)

        if result is None:
            return None  # Native unavailable, fall back to Python UIA
        return _format_pattern_result(result, x, y)

    @mcp.tool(
        name="Invoke",
        description=(
            "Invokes a UIA automation pattern on an element at coordinates [x, y]. "
            "Actions: 'invoke' (click buttons), 'toggle' (checkboxes/switches), "
            "'set_value' (type into fields without clicking), 'expand'/'collapse' (dropdowns/trees), "
            "'select' (list items). More reliable than coordinate-based Click for UI automation."
        ),
        annotations=ToolAnnotations(
            title="Invoke",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Invoke-Tool")
    def invoke_tool(
        loc: list[int],
        action: Literal["invoke", "toggle", "set_value", "expand", "collapse", "select"] = "invoke",
        value: str | None = None,
        ctx: Context = None,
    ) -> str:
        try:
            x, y = _validate_loc(loc)
        except ValueError:
            return "Error: loc must be [x, y] with integer coordinates."

        # --- Rust native fast-path ---
        result = _try_native_pattern(x, y, action, value)
        if result is not None:
            return result

        # --- Python UIA fallback ---
        from windows_mcp.uia import ControlFromPoint, PatternId

        element = ControlFromPoint(x, y)
        if not element:
            return f"Error: no element found at ({x},{y})."

        element_name = element.Name or element.AutomationId or "unnamed"
        element_type = element.LocalizedControlType or "unknown"

        try:
            if action == "invoke":
                pattern = element.GetPattern(PatternId.InvokePattern)
                if not pattern:
                    return f"Error: element '{element_name}' does not support InvokePattern."
                pattern.Invoke()
                return f"Invoked '{element_name}' ({element_type}) at ({x},{y})."

            elif action == "toggle":
                pattern = element.GetPattern(PatternId.TogglePattern)
                if not pattern:
                    return f"Error: element '{element_name}' does not support TogglePattern."
                pattern.Toggle()
                state = pattern.ToggleState
                state_name = {0: "off", 1: "on", 2: "indeterminate"}.get(state, str(state))
                return (
                    f"Toggled '{element_name}' ({element_type}) at ({x},{y}). State: {state_name}."
                )

            elif action == "set_value":
                if value is None:
                    return "Error: value parameter required for set_value action."
                if len(value) > 10000:
                    return f"Error: value too long ({len(value)} chars, max 10000)."
                pattern = element.GetPattern(PatternId.ValuePattern)
                if not pattern:
                    return f"Error: element '{element_name}' does not support ValuePattern."
                pattern.SetValue(value)
                preview = value[:50] + "..." if len(value) > 50 else value
                return f"Set value '{preview}' on '{element_name}' ({element_type}) at ({x},{y})."

            elif action == "expand":
                pattern = element.GetPattern(PatternId.ExpandCollapsePattern)
                if not pattern:
                    return (
                        f"Error: element '{element_name}' does not support ExpandCollapsePattern."
                    )
                pattern.Expand()
                return f"Expanded '{element_name}' ({element_type}) at ({x},{y})."

            elif action == "collapse":
                pattern = element.GetPattern(PatternId.ExpandCollapsePattern)
                if not pattern:
                    return (
                        f"Error: element '{element_name}' does not support ExpandCollapsePattern."
                    )
                pattern.Collapse()
                return f"Collapsed '{element_name}' ({element_type}) at ({x},{y})."

            elif action == "select":
                pattern = element.GetPattern(PatternId.SelectionItemPattern)
                if not pattern:
                    return f"Error: element '{element_name}' does not support SelectionItemPattern."
                pattern.Select()
                return f"Selected '{element_name}' ({element_type}) at ({x},{y})."

            else:
                return f"Error: unknown action '{action}'."

        except Exception as e:
            return f"Error invoking {action} on '{element_name}': {str(e)}"
