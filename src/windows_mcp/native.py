"""Centralised native extension adapter.

All Python modules that want to use the Rust extension should import from
here rather than importing ``windows_mcp_core`` directly.  This module
handles the try-import and provides typed wrapper functions that return
``None`` when the extension is unavailable, allowing callers to fall back
to pure Python.

Usage::

    from windows_mcp.native import HAS_NATIVE, native_system_info

    result = native_system_info()
    if result is None:
        # fallback to psutil / PowerShell
        ...
"""

import logging

logger = logging.getLogger(__name__)

try:
    import windows_mcp_core

    HAS_NATIVE = True
    NATIVE_VERSION = windows_mcp_core.__version__
    logger.info("Native extension loaded: windows_mcp_core %s", NATIVE_VERSION)
except ImportError:
    windows_mcp_core = None  # type: ignore[assignment]
    HAS_NATIVE = False
    NATIVE_VERSION = None
    logger.debug("Native extension not available, using pure Python fallbacks")


# ---------------------------------------------------------------------------
# system_info
# ---------------------------------------------------------------------------


def native_system_info() -> dict | None:
    """Collect system information via Rust sysinfo crate.

    Returns a dict with keys: os_name, os_version, hostname, cpu_count,
    cpu_usage_percent, total_memory_bytes, used_memory_bytes, disks.
    Returns None if the native extension is unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.system_info()
    except Exception:
        logger.warning("native_system_info failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# capture_tree
# ---------------------------------------------------------------------------


def native_capture_tree(handles: list[int], max_depth: int = 50) -> list[dict] | None:
    """Capture UIA accessibility tree via Rust + Rayon.

    Args:
        handles: List of window HWNDs as integers.
        max_depth: Maximum tree recursion depth.

    Returns a list of nested dicts, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.capture_tree(handles, max_depth=max_depth)
    except Exception:
        logger.warning("native_capture_tree failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# Input functions
# ---------------------------------------------------------------------------


def native_send_text(text: str) -> int | None:
    """Type Unicode text via Win32 SendInput.

    Returns the number of events injected, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_text(text)
    except Exception:
        logger.warning("native_send_text failed, falling back", exc_info=True)
        return None


def native_send_click(x: int, y: int, button: str = "left") -> int | None:
    """Click at absolute screen coordinates via Win32 SendInput.

    Returns the number of events injected, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_click(x, y, button)
    except Exception:
        logger.warning("native_send_click failed, falling back", exc_info=True)
        return None


def native_send_key(vk_code: int, key_up: bool = False) -> int | None:
    """Press or release a virtual key code via Win32 SendInput.

    Returns 1 on success, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_key(vk_code, key_up)
    except Exception:
        logger.warning("native_send_key failed, falling back", exc_info=True)
        return None


def native_send_mouse_move(x: int, y: int) -> int | None:
    """Move cursor to screen coordinates via Win32 SendInput.

    Returns 1 on success, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_mouse_move(x, y)
    except Exception:
        logger.warning("native_send_mouse_move failed, falling back", exc_info=True)
        return None


def native_send_hotkey(vk_codes: list[int]) -> int | None:
    """Send a key combination via Win32 SendInput.

    Returns the number of events injected, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_hotkey(vk_codes)
    except Exception:
        logger.warning("native_send_hotkey failed, falling back", exc_info=True)
        return None


def native_send_scroll(x: int, y: int, delta: int, horizontal: bool = False) -> int | None:
    """Scroll the mouse wheel via Win32 SendInput.

    Args:
        x, y: Screen coordinates.
        delta: Wheel delta in WHEEL_DELTA units (120 = one notch).
        horizontal: True for horizontal scrolling.

    Returns the number of events injected, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_scroll(x, y, delta, horizontal)
    except Exception:
        logger.warning("native_send_scroll failed, falling back", exc_info=True)
        return None


def native_send_drag(to_x: int, to_y: int, steps: int = 10) -> int | None:
    """Drag the mouse from current position to destination via Win32 SendInput.

    Args:
        to_x, to_y: Destination screen coordinates.
        steps: Reserved for future interpolation.

    Returns the number of events injected, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.send_drag(to_x, to_y, steps)
    except Exception:
        logger.warning("native_send_drag failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# Window functions
# ---------------------------------------------------------------------------


def native_list_windows() -> list[dict] | None:
    """List all visible windows with full info.

    Returns a list of dicts (same shape as native_get_window_info),
    or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.list_windows()
    except Exception:
        logger.warning("native_list_windows failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# Screenshot functions
# ---------------------------------------------------------------------------


def native_capture_screenshot_png(monitor_index: int = 0) -> bytes | None:
    """Capture a screenshot via DXGI/GDI and return PNG bytes.

    Args:
        monitor_index: Zero-based monitor index (0 = primary).

    Returns PNG file bytes, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.capture_screenshot_png(monitor_index)
    except Exception:
        logger.warning("native_capture_screenshot_png failed, falling back", exc_info=True)
        return None


def native_capture_screenshot_raw(monitor_index: int = 0) -> dict | None:
    """Capture a screenshot via DXGI/GDI and return raw BGRA data.

    Args:
        monitor_index: Zero-based monitor index (0 = primary).

    Returns dict with keys: width, height, data (bytes), or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.capture_screenshot_raw(monitor_index)
    except Exception:
        logger.warning("native_capture_screenshot_raw failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# UIA query functions
# ---------------------------------------------------------------------------


def native_element_from_point(x: int, y: int) -> dict | None:
    """Query the UIA element at screen coordinates via Rust.

    Returns a dict with keys: name, automation_id, control_type,
    localized_control_type, class_name, bounding_rect, is_enabled,
    is_offscreen, has_keyboard_focus, supported_patterns.
    Returns None if the native extension is unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.element_from_point(x, y)
    except Exception:
        logger.warning("native_element_from_point failed, falling back", exc_info=True)
        return None


def native_find_elements(
    name: str | None = None,
    control_type: str | None = None,
    automation_id: str | None = None,
    window_handle: int | None = None,
    limit: int = 20,
) -> list[dict] | None:
    """Search for UIA elements matching criteria via Rust.

    Args:
        name: Substring match (case-insensitive).
        control_type: Exact match (e.g. "Button").
        automation_id: Exact match.
        window_handle: Scope to specific window.
        limit: Maximum results (default 20, max 100).

    Returns a list of ElementInfo dicts, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.find_elements(
            name=name,
            control_type=control_type,
            automation_id=automation_id,
            window_handle=window_handle,
            limit=limit,
        )
    except Exception:
        logger.warning("native_find_elements failed, falling back", exc_info=True)
        return None


def native_get_screen_metrics() -> dict | None:
    """Query primary and virtual screen dimensions via Rust.

    Returns dict with keys: primary_width, primary_height,
    virtual_width, virtual_height. Returns None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.get_screen_metrics()
    except Exception:
        logger.warning("native_get_screen_metrics failed, falling back", exc_info=True)
        return None


# ---------------------------------------------------------------------------
# UIA pattern functions
# ---------------------------------------------------------------------------


def native_invoke_at(x: int, y: int) -> dict | None:
    """Invoke the InvokePattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.invoke_at(x, y)
    except Exception:
        logger.warning("native_invoke_at failed, falling back", exc_info=True)
        return None


def native_toggle_at(x: int, y: int) -> dict | None:
    """Toggle the TogglePattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.toggle_at(x, y)
    except Exception:
        logger.warning("native_toggle_at failed, falling back", exc_info=True)
        return None


def native_set_value_at(x: int, y: int, value: str) -> dict | None:
    """Set value via ValuePattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.set_value_at(x, y, value)
    except Exception:
        logger.warning("native_set_value_at failed, falling back", exc_info=True)
        return None


def native_expand_at(x: int, y: int) -> dict | None:
    """Expand via ExpandCollapsePattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.expand_at(x, y)
    except Exception:
        logger.warning("native_expand_at failed, falling back", exc_info=True)
        return None


def native_collapse_at(x: int, y: int) -> dict | None:
    """Collapse via ExpandCollapsePattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.collapse_at(x, y)
    except Exception:
        logger.warning("native_collapse_at failed, falling back", exc_info=True)
        return None


def native_select_at(x: int, y: int) -> dict | None:
    """Select via SelectionItemPattern on the element at (x, y) via Rust.

    Returns a PatternResult dict, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.select_at(x, y)
    except Exception:
        logger.warning("native_select_at failed, falling back", exc_info=True)
        return None
