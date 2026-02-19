//! Keyboard and mouse input simulation via Win32 `SendInput`.
//!
//! All functions are pure Rust with no PyO3 dependency.  PyO3 wrappers
//! in `wmcp-pyo3` call these via `py.allow_threads()`.
//!
//! # Performance
//!
//! `SendInput` batches multiple events atomically, avoiding per-event
//! overhead.  Each function completes in <1ms.

use windows::Win32::UI::Input::KeyboardAndMouse::{
    SendInput, INPUT, INPUT_0, INPUT_KEYBOARD, INPUT_MOUSE, KEYBDINPUT, KEYBD_EVENT_FLAGS,
    KEYEVENTF_KEYUP, KEYEVENTF_UNICODE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN,
    MOUSEEVENTF_LEFTUP, MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP, MOUSEEVENTF_MOVE,
    MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, MOUSEEVENTF_VIRTUALDESK, MOUSEINPUT,
    MOUSE_EVENT_FLAGS, VIRTUAL_KEY,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    MOUSEEVENTF_HWHEEL, MOUSEEVENTF_WHEEL,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetSystemMetrics, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN,
    SM_YVIRTUALSCREEN,
};

/// Maximum text length to prevent unbounded allocation.
const MAX_TEXT_LENGTH: usize = 10_000;

/// Maximum hotkey combo length (no real hotkey uses more than 5-6 keys).
const MAX_HOTKEY_KEYS: usize = 8;

/// Pre-computed size of `INPUT` struct for `SendInput` calls.
const INPUT_SIZE: i32 = std::mem::size_of::<INPUT>() as i32;

/// Query virtual screen dimensions and origin (covers all monitors).
///
/// Returns `(origin_x, origin_y, width, height)`.  On multi-monitor setups
/// where a monitor is left of or above the primary, origin can be negative.
fn screen_geometry() -> (i32, i32, i32, i32) {
    unsafe {
        let x = GetSystemMetrics(SM_XVIRTUALSCREEN);
        let y = GetSystemMetrics(SM_YVIRTUALSCREEN);
        let w = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        let h = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        // Fallback: GetSystemMetrics returns 0 on failure
        if w > 0 && h > 0 {
            (x, y, w, h)
        } else {
            (0, 0, 1920, 1080)
        }
    }
}

// ---------------------------------------------------------------------------
// Helpers: build INPUT structs
// ---------------------------------------------------------------------------

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

fn mouse_input_with_data(abs_x: i32, abs_y: i32, data: i32, flags: MOUSE_EVENT_FLAGS) -> INPUT {
    INPUT {
        r#type: INPUT_MOUSE,
        Anonymous: INPUT_0 {
            mi: MOUSEINPUT {
                dx: abs_x,
                dy: abs_y,
                // Win32 treats mouseData as signed for WHEEL/HWHEEL events.
                // Rust `as u32` is a bitwise reinterpret, preserving the sign bits.
                mouseData: data as u32,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

/// Convert pixel coordinates to 0..65535 normalised space for the virtual desktop.
///
/// Accounts for the virtual screen origin (which can be negative on multi-monitor
/// setups where a monitor is left of or above the primary).
///
/// Uses the MSDN formula: `((pixel - origin) * 65535) / (screen_size - 1)`.
/// Result is clamped to `[0, 65535]` to prevent out-of-range values.
fn normalise_coords(x: i32, y: i32) -> (i32, i32) {
    let (origin_x, origin_y, screen_w, screen_h) = screen_geometry();

    if screen_w <= 1 || screen_h <= 1 {
        return (0, 0);
    }

    let abs_x = (((x - origin_x) as i64 * 65535) / (screen_w as i64 - 1)).clamp(0, 65535) as i32;
    let abs_y = (((y - origin_y) as i64 * 65535) / (screen_h as i64 - 1)).clamp(0, 65535) as i32;
    (abs_x, abs_y)
}

/// Flags for absolute mouse positioning on the virtual desktop.
const ABSOLUTE_MOVE: MOUSE_EVENT_FLAGS =
    MOUSE_EVENT_FLAGS(MOUSEEVENTF_ABSOLUTE.0 | MOUSEEVENTF_MOVE.0 | MOUSEEVENTF_VIRTUALDESK.0);

// ---------------------------------------------------------------------------
// Public API -- raw functions (no PyO3)
// ---------------------------------------------------------------------------

/// Type Unicode text via `KEYEVENTF_UNICODE` events.
///
/// Returns the number of input events successfully injected.
/// Returns 0 if text is empty or exceeds `MAX_TEXT_LENGTH` (10,000 chars).
pub fn send_text_raw(text: &str) -> u32 {
    if text.is_empty() || text.len() > MAX_TEXT_LENGTH {
        return 0;
    }

    let chars: Vec<u16> = text.encode_utf16().collect();
    let mut inputs: Vec<INPUT> = Vec::with_capacity(chars.len() * 2);
    for &ch in &chars {
        inputs.push(unicode_key_input(ch, false));
        inputs.push(unicode_key_input(ch, true));
    }
    unsafe { SendInput(&inputs, INPUT_SIZE) }
}

/// Press or release a virtual key code.
///
/// Returns 1 on success, 0 on failure.
pub fn send_key_raw(vk_code: u16, key_up: bool) -> u32 {
    let input = virtual_key_input(vk_code, key_up);
    unsafe { SendInput(&[input], INPUT_SIZE) }
}

/// Click the mouse at absolute screen coordinates.
///
/// Returns the number of events injected (2 on success: down + up).
pub fn send_click_raw(x: i32, y: i32, button: &str) -> u32 {
    let (abs_x, abs_y) = normalise_coords(x, y);

    let (down_flag, up_flag) = match button {
        "right" => (MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP),
        "middle" => (MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP),
        _ => (MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP),
    };

    let inputs = [
        mouse_input(abs_x, abs_y, MOUSE_EVENT_FLAGS(ABSOLUTE_MOVE.0 | down_flag.0)),
        mouse_input(abs_x, abs_y, MOUSE_EVENT_FLAGS(ABSOLUTE_MOVE.0 | up_flag.0)),
    ];

    unsafe { SendInput(&inputs, INPUT_SIZE) }
}

/// Move the mouse cursor to absolute screen coordinates without clicking.
///
/// Returns 1 on success.
pub fn send_mouse_move_raw(x: i32, y: i32) -> u32 {
    let (abs_x, abs_y) = normalise_coords(x, y);
    let input = mouse_input(abs_x, abs_y, ABSOLUTE_MOVE);
    unsafe { SendInput(&[input], INPUT_SIZE) }
}

/// Send a key combination (e.g. Ctrl+C, Alt+Tab).
///
/// Presses all keys in order, releases in reverse -- all in a single
/// atomic `SendInput` call.
///
/// Returns 0 if `vk_codes` is empty or exceeds `MAX_HOTKEY_KEYS` (8).
pub fn send_hotkey_raw(vk_codes: &[u16]) -> u32 {
    if vk_codes.is_empty() || vk_codes.len() > MAX_HOTKEY_KEYS {
        return 0;
    }

    let mut inputs: Vec<INPUT> = Vec::with_capacity(vk_codes.len() * 2);

    for &vk in vk_codes {
        inputs.push(virtual_key_input(vk, false));
    }
    for &vk in vk_codes.iter().rev() {
        inputs.push(virtual_key_input(vk, true));
    }

    unsafe { SendInput(&inputs, INPUT_SIZE) }
}

/// Scroll the mouse wheel at absolute screen coordinates.
///
/// `delta` is in WHEEL_DELTA units (120 = one notch).
/// `horizontal` selects horizontal vs vertical scrolling.
///
/// Returns the number of events injected (2: move + wheel).
pub fn send_scroll_raw(x: i32, y: i32, delta: i32, horizontal: bool) -> u32 {
    let (abs_x, abs_y) = normalise_coords(x, y);

    let wheel_flag = if horizontal {
        MOUSEEVENTF_HWHEEL
    } else {
        MOUSEEVENTF_WHEEL
    };

    // Move and wheel MUST be separate INPUT events -- combining
    // MOUSEEVENTF_MOVE with MOUSEEVENTF_WHEEL is undefined behavior.
    let inputs = [
        mouse_input(abs_x, abs_y, ABSOLUTE_MOVE),
        mouse_input_with_data(0, 0, delta, wheel_flag),
    ];
    unsafe { SendInput(&inputs, INPUT_SIZE) }
}

/// Drag the mouse from current position to (`to_x`, `to_y`).
///
/// Sends: left-button-down at current position, move to destination,
/// left-button-up at destination.  The caller must ensure the cursor is
/// already at the desired drag origin.
///
/// `steps` is reserved for future interpolation (currently ignored).
///
/// Returns total events injected (3 on success).
pub fn send_drag_raw(to_x: i32, to_y: i32, _steps: u32) -> u32 {
    let (abs_to_x, abs_to_y) = normalise_coords(to_x, to_y);

    let inputs = [
        // Press left button at current position (relative 0,0)
        mouse_input(0, 0, MOUSEEVENTF_LEFTDOWN),
        // Move to destination while holding
        mouse_input(abs_to_x, abs_to_y, ABSOLUTE_MOVE),
        // Release left button at destination
        mouse_input(abs_to_x, abs_to_y, MOUSE_EVENT_FLAGS(ABSOLUTE_MOVE.0 | MOUSEEVENTF_LEFTUP.0)),
    ];

    unsafe { SendInput(&inputs, INPUT_SIZE) }
}
