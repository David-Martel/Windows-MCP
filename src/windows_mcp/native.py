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


def native_enumerate_windows() -> list[int] | None:
    """Enumerate all visible top-level windows via Win32 EnumWindows.

    Returns a list of window handles (HWNDs), or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.enumerate_windows()
    except Exception:
        logger.warning("native_enumerate_windows failed, falling back", exc_info=True)
        return None


def native_get_window_info(hwnd: int) -> dict | None:
    """Get detailed info about a window (title, class, pid, rect, state).

    Returns a dict with keys: hwnd, title, class_name, pid, rect,
    is_minimized, is_maximized, is_visible.  None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.get_window_info(hwnd)
    except Exception:
        logger.warning("native_get_window_info failed", exc_info=True)
        return None


def native_get_foreground_window() -> int | None:
    """Get the foreground (active) window handle via Win32.

    Returns the HWND as int, or None if unavailable.
    """
    if not HAS_NATIVE:
        return None
    try:
        return windows_mcp_core.get_foreground_window()
    except Exception:
        logger.warning("native_get_foreground_window failed", exc_info=True)
        return None


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
