//! Keyboard and mouse input simulation via Win32 `SendInput`.
//!
//! Replaces `pyautogui` calls which add ~50ms+ of Python overhead per action.
//! Each function releases the GIL via `py.allow_threads()` so other Python
//! threads can run during the (fast) Win32 call.
//!
//! # Functions
//!
//! | Function | Purpose |
//! |----------|---------|
//! | [`send_text`] | Type Unicode text via `KEYEVENTF_UNICODE` |
//! | [`send_key`] | Press/release a virtual key code |
//! | [`send_click`] | Mouse click at absolute screen coordinates |
//! | [`send_mouse_move`] | Move cursor to absolute screen coordinates |
//!
//! # Performance
//!
//! `SendInput` is the modern Win32 input injection API (replacing deprecated
//! `keybd_event` / `mouse_event`).  A single `SendInput` call can batch many
//! events atomically, avoiding per-event overhead.
//!
//! # Safety
//!
//! All `SendInput` calls require the calling process to have the foreground
//! window or be running with UI Access permissions.  If the process does not
//! have input injection rights, `SendInput` returns 0 (no events sent).

use std::mem;

use pyo3::prelude::*;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    SendInput, INPUT, INPUT_0, INPUT_KEYBOARD, INPUT_MOUSE, KEYBDINPUT, KEYBD_EVENT_FLAGS,
    KEYEVENTF_KEYUP, KEYEVENTF_UNICODE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN,
    MOUSEEVENTF_LEFTUP, MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP, MOUSEEVENTF_MOVE,
    MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, MOUSEINPUT, MOUSE_EVENT_FLAGS, VIRTUAL_KEY,
};
use windows::Win32::UI::WindowsAndMessaging::{GetSystemMetrics, SM_CXSCREEN, SM_CYSCREEN};

// WindowsMcpError::InputError available for future use (e.g. UIPI failures).

// ---------------------------------------------------------------------------
// Helper: build keyboard INPUT
// ---------------------------------------------------------------------------

/// Create a `KEYEVENTF_UNICODE` input event for a UTF-16 code unit.
fn unicode_key_input(scan_code: u16, key_up: bool) -> INPUT {
    let flags = if key_up {
        KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
    } else {
        KEYEVENTF_UNICODE
    };

    INPUT {
        r#type: INPUT_KEYBOARD,
        Anonymous: INPUT_0 {
            ki: KEYBDINPUT {
                wVk: VIRTUAL_KEY(0),
                wScan: scan_code,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

/// Create a virtual-key input event.
fn virtual_key_input(vk: u16, key_up: bool) -> INPUT {
    let flags = if key_up {
        KEYEVENTF_KEYUP
    } else {
        KEYBD_EVENT_FLAGS(0)
    };

    INPUT {
        r#type: INPUT_KEYBOARD,
        Anonymous: INPUT_0 {
            ki: KEYBDINPUT {
                wVk: VIRTUAL_KEY(vk),
                wScan: 0,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

// ---------------------------------------------------------------------------
// Helper: build mouse INPUT
// ---------------------------------------------------------------------------

/// Create an absolute-position mouse input event.
///
/// `abs_x` and `abs_y` are in the 0..65535 normalised coordinate space
/// required by `MOUSEEVENTF_ABSOLUTE`.
fn mouse_input(abs_x: i32, abs_y: i32, flags: MOUSE_EVENT_FLAGS) -> INPUT {
    INPUT {
        r#type: INPUT_MOUSE,
        Anonymous: INPUT_0 {
            mi: MOUSEINPUT {
                dx: abs_x,
                dy: abs_y,
                mouseData: 0,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

/// Convert pixel coordinates to the 0..65535 normalised space.
fn normalise_coords(x: i32, y: i32) -> (i32, i32) {
    let screen_w = unsafe { GetSystemMetrics(SM_CXSCREEN) };
    let screen_h = unsafe { GetSystemMetrics(SM_CYSCREEN) };

    if screen_w <= 0 || screen_h <= 0 {
        return (0, 0);
    }

    // The normalised coordinate formula per MSDN:
    // abs = (pixel * 65536 / screen_dim) + 1   (rounds up to avoid off-by-one)
    let abs_x = ((x as i64 * 65536) / screen_w as i64 + 1) as i32;
    let abs_y = ((y as i64 * 65536) / screen_h as i64 + 1) as i32;
    (abs_x, abs_y)
}

// ---------------------------------------------------------------------------
// Public pyfunctions
// ---------------------------------------------------------------------------

/// Type Unicode text by injecting `KEYEVENTF_UNICODE` events via `SendInput`.
///
/// This is **much** faster than pyautogui's character-by-character approach
/// because all characters are batched into a single `SendInput` call.
///
/// # Arguments
///
/// * `text` -- The string to type.  Supports any Unicode character including
///   emoji, CJK, and combining marks.
///
/// # Returns
///
/// The number of input events successfully injected (should equal
/// `2 * len(text)` -- one down + one up per character).
///
/// # Example
///
/// ```python
/// import windows_mcp_core
/// sent = windows_mcp_core.send_text("Hello, World!")
/// assert sent == 26  # 13 chars * 2 events each
/// ```
#[pyfunction]
#[pyo3(signature = (text,))]
pub fn send_text(py: Python<'_>, text: &str) -> PyResult<u32> {
    let chars: Vec<u16> = text.encode_utf16().collect();

    if chars.is_empty() {
        return Ok(0);
    }

    let sent = py.allow_threads(move || -> u32 {
        let mut inputs: Vec<INPUT> = Vec::with_capacity(chars.len() * 2);
        for &ch in &chars {
            inputs.push(unicode_key_input(ch, false));
            inputs.push(unicode_key_input(ch, true));
        }
        unsafe { SendInput(&inputs, mem::size_of::<INPUT>() as i32) }
    });

    Ok(sent)
}

/// Press or release a virtual key code via `SendInput`.
///
/// # Arguments
///
/// * `vk_code` -- Win32 virtual key code (e.g. `0x0D` for Enter, `0x09` for
///   Tab).  See
///   <https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes>.
/// * `key_up` -- `True` to release the key, `False` (default) to press it.
///
/// # Returns
///
/// Number of events injected (1 on success, 0 on failure).
///
/// # Example
///
/// ```python
/// import windows_mcp_core
/// VK_RETURN = 0x0D
/// windows_mcp_core.send_key(VK_RETURN)        # key down
/// windows_mcp_core.send_key(VK_RETURN, True)   # key up
/// ```
#[pyfunction]
#[pyo3(signature = (vk_code, key_up=false))]
pub fn send_key(py: Python<'_>, vk_code: u16, key_up: bool) -> PyResult<u32> {
    let sent = py.allow_threads(move || -> u32 {
        let input = virtual_key_input(vk_code, key_up);
        unsafe { SendInput(&[input], mem::size_of::<INPUT>() as i32) }
    });
    Ok(sent)
}

/// Click the mouse at absolute screen coordinates.
///
/// Sends a move + button-down + button-up sequence in a single `SendInput`
/// call for atomicity.
///
/// # Arguments
///
/// * `x`, `y` -- Pixel coordinates on the primary monitor.
/// * `button` -- `"left"` (default), `"right"`, or `"middle"`.
///
/// # Returns
///
/// Number of events injected (3 on success: move + down + up).
///
/// # Example
///
/// ```python
/// import windows_mcp_core
/// windows_mcp_core.send_click(500, 300)                # left click
/// windows_mcp_core.send_click(500, 300, "right")       # right click
/// ```
#[pyfunction]
#[pyo3(signature = (x, y, button="left"))]
pub fn send_click(py: Python<'_>, x: i32, y: i32, button: &str) -> PyResult<u32> {
    let button_owned = button.to_lowercase();

    let sent = py.allow_threads(move || -> u32 {
        let (abs_x, abs_y) = normalise_coords(x, y);

        let (down_flag, up_flag) = match button_owned.as_str() {
            "right" => (MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP),
            "middle" => (MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP),
            _ => (MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP),
        };

        let move_flags = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE;

        let inputs = [
            mouse_input(abs_x, abs_y, move_flags | down_flag),
            mouse_input(abs_x, abs_y, move_flags | up_flag),
        ];

        unsafe { SendInput(&inputs, mem::size_of::<INPUT>() as i32) }
    });

    Ok(sent)
}

/// Move the mouse cursor to absolute screen coordinates without clicking.
///
/// # Arguments
///
/// * `x`, `y` -- Pixel coordinates on the primary monitor.
///
/// # Returns
///
/// Number of events injected (1 on success).
///
/// # Example
///
/// ```python
/// import windows_mcp_core
/// windows_mcp_core.send_mouse_move(960, 540)  # center of 1920x1080
/// ```
#[pyfunction]
#[pyo3(signature = (x, y))]
pub fn send_mouse_move(py: Python<'_>, x: i32, y: i32) -> PyResult<u32> {
    let sent = py.allow_threads(move || -> u32 {
        let (abs_x, abs_y) = normalise_coords(x, y);
        let flags = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE;
        let input = mouse_input(abs_x, abs_y, flags);
        unsafe { SendInput(&[input], mem::size_of::<INPUT>() as i32) }
    });
    Ok(sent)
}

/// Send a key combination (e.g. Ctrl+C, Alt+Tab).
///
/// Presses all modifier keys, then the main key, then releases in reverse
/// order -- all in a single atomic `SendInput` call.
///
/// # Arguments
///
/// * `vk_codes` -- List of virtual key codes.  The last element is the main
///   key; all preceding elements are modifiers held during the press.
///
/// # Returns
///
/// Number of events injected.
///
/// # Example
///
/// ```python
/// import windows_mcp_core
/// VK_CONTROL = 0x11
/// VK_C = 0x43
/// windows_mcp_core.send_hotkey([VK_CONTROL, VK_C])  # Ctrl+C
/// ```
#[pyfunction]
#[pyo3(signature = (vk_codes,))]
pub fn send_hotkey(py: Python<'_>, vk_codes: Vec<u16>) -> PyResult<u32> {
    if vk_codes.is_empty() {
        return Ok(0);
    }

    let sent = py.allow_threads(move || -> u32 {
        let mut inputs: Vec<INPUT> = Vec::with_capacity(vk_codes.len() * 2);

        // Press all keys in order.
        for &vk in &vk_codes {
            inputs.push(virtual_key_input(vk, false));
        }
        // Release in reverse order.
        for &vk in vk_codes.iter().rev() {
            inputs.push(virtual_key_input(vk, true));
        }

        unsafe { SendInput(&inputs, mem::size_of::<INPUT>() as i32) }
    });

    Ok(sent)
}
