//! C ABI DLL for windows-mcp -- loadable by ctypes, C#, or any FFI consumer.
//!
//! All exported functions follow the convention:
//! - Return `i32` status code: `WMCP_OK=0`, `WMCP_ERROR=-1`
//! - String outputs allocated by Rust, freed via `wmcp_free_string()`
//! - Last error retrievable via `wmcp_last_error()`

use std::ffi::{c_char, CStr, CString};
use std::ptr;
use std::cell::RefCell;

pub const WMCP_OK: i32 = 0;
pub const WMCP_ERROR: i32 = -1;

/// Maximum handles to process in `wmcp_capture_tree` to prevent
/// unreasonable allocations from corrupted input.
const MAX_HANDLE_COUNT: usize = 256;

/// Maximum text length for `wmcp_send_text`.
const MAX_TEXT_LENGTH: usize = 10_000;

thread_local! {
    static LAST_ERROR: RefCell<Option<CString>> = const { RefCell::new(None) };
}

fn set_last_error(msg: &str) {
    LAST_ERROR.with(|e| {
        *e.borrow_mut() = CString::new(msg).ok();
    });
}

/// Retrieve the last error message (thread-local).
///
/// Returns a **heap-allocated** copy of the error string.  The caller owns
/// the returned pointer and **must** free it with `wmcp_free_string()`.
/// Returns null if no error has occurred.
///
/// This avoids the dangling-pointer hazard of returning a borrow into
/// thread-local storage.
#[no_mangle]
pub extern "C" fn wmcp_last_error() -> *mut c_char {
    LAST_ERROR.with(|e| {
        e.borrow()
            .as_ref()
            .and_then(|s| CString::new(s.as_bytes()).ok())
            .map(|copy| copy.into_raw())
            .unwrap_or(ptr::null_mut())
    })
}

/// Free a string previously allocated by a wmcp_* function.
///
/// # Safety
///
/// `ptr` must be a pointer returned by a wmcp_* function or null.
#[no_mangle]
pub unsafe extern "C" fn wmcp_free_string(ptr: *mut c_char) {
    if !ptr.is_null() {
        drop(unsafe { CString::from_raw(ptr) });
    }
}

/// Collect system information as a JSON string.
///
/// # Safety
///
/// `out_json` must be a valid pointer to a `*mut c_char`.
/// On success, `*out_json` is set to a heap-allocated JSON C string.
/// Caller must free with `wmcp_free_string()`.
#[no_mangle]
pub unsafe extern "C" fn wmcp_system_info(out_json: *mut *mut c_char) -> i32 {
    if out_json.is_null() {
        set_last_error("out_json is null");
        return WMCP_ERROR;
    }

    match wmcp_core::system_info::collect_system_info() {
        Ok(snapshot) => match serde_json::to_string(&snapshot) {
            Ok(json) => match CString::new(json) {
                Ok(cstr) => {
                    unsafe { *out_json = cstr.into_raw() };
                    WMCP_OK
                }
                Err(e) => {
                    set_last_error(&format!("CString conversion failed: {e}"));
                    WMCP_ERROR
                }
            },
            Err(e) => {
                set_last_error(&format!("JSON serialization failed: {e}"));
                WMCP_ERROR
            }
        },
        Err(e) => {
            set_last_error(&e.to_string());
            WMCP_ERROR
        }
    }
}

/// Send Unicode text via SendInput.
///
/// # Safety
///
/// `text` must be a valid null-terminated UTF-8 C string.
/// `out_count` is optional (may be null).
#[no_mangle]
pub unsafe extern "C" fn wmcp_send_text(text: *const c_char, out_count: *mut u32) -> i32 {
    if text.is_null() {
        set_last_error("text is null");
        return WMCP_ERROR;
    }

    let text_str = match unsafe { CStr::from_ptr(text) }.to_str() {
        Ok(s) => s,
        Err(e) => {
            set_last_error(&format!("Invalid UTF-8: {e}"));
            return WMCP_ERROR;
        }
    };

    if text_str.len() > MAX_TEXT_LENGTH {
        set_last_error(&format!(
            "text length {} exceeds maximum {MAX_TEXT_LENGTH}",
            text_str.len()
        ));
        return WMCP_ERROR;
    }

    let count = wmcp_core::input::send_text_raw(text_str);
    if !out_count.is_null() {
        unsafe { *out_count = count };
    }
    WMCP_OK
}

/// Click the mouse at absolute screen coordinates.
///
/// Returns `WMCP_OK` on success, `WMCP_ERROR` if SendInput failed.
#[no_mangle]
pub extern "C" fn wmcp_send_click(x: i32, y: i32, button: i32) -> i32 {
    let button_str = match button {
        1 => "right",
        2 => "middle",
        _ => "left",
    };
    let count = wmcp_core::input::send_click_raw(x, y, button_str);
    if count == 0 {
        set_last_error("SendInput returned 0 events for click");
        WMCP_ERROR
    } else {
        WMCP_OK
    }
}

/// Move the mouse cursor to absolute screen coordinates.
///
/// Returns `WMCP_OK` on success.
#[no_mangle]
pub extern "C" fn wmcp_send_mouse_move(x: i32, y: i32) -> i32 {
    wmcp_core::input::send_mouse_move_raw(x, y);
    WMCP_OK
}

/// Scroll the mouse wheel at absolute screen coordinates.
///
/// `delta` is in WHEEL_DELTA units (120 = one notch).
/// `horizontal`: 0 = vertical, nonzero = horizontal.
#[no_mangle]
pub extern "C" fn wmcp_send_scroll(x: i32, y: i32, delta: i32, horizontal: i32) -> i32 {
    wmcp_core::input::send_scroll_raw(x, y, delta, horizontal != 0);
    WMCP_OK
}

/// Send a key combination (e.g. Ctrl+C).
///
/// # Safety
///
/// `vk_codes` must point to `count` contiguous `u16` values.
#[no_mangle]
pub unsafe extern "C" fn wmcp_send_hotkey(vk_codes: *const u16, count: usize) -> i32 {
    if vk_codes.is_null() || count == 0 {
        set_last_error("null or empty vk_codes");
        return WMCP_ERROR;
    }
    if count > 8 {
        set_last_error("hotkey count exceeds maximum 8");
        return WMCP_ERROR;
    }
    let codes = unsafe { std::slice::from_raw_parts(vk_codes, count) };
    wmcp_core::input::send_hotkey_raw(codes);
    WMCP_OK
}

/// Enumerate visible windows as a JSON array of handle integers.
///
/// # Safety
///
/// `out_json` must be a valid pointer. Caller must free with `wmcp_free_string()`.
#[no_mangle]
pub unsafe extern "C" fn wmcp_enumerate_windows(out_json: *mut *mut c_char) -> i32 {
    if out_json.is_null() {
        set_last_error("out_json is null");
        return WMCP_ERROR;
    }
    match wmcp_core::window::enumerate_visible_windows() {
        Ok(handles) => match serde_json::to_string(&handles) {
            Ok(json) => match CString::new(json) {
                Ok(cstr) => {
                    unsafe { *out_json = cstr.into_raw() };
                    WMCP_OK
                }
                Err(e) => {
                    set_last_error(&format!("CString conversion failed: {e}"));
                    WMCP_ERROR
                }
            },
            Err(e) => {
                set_last_error(&format!("JSON serialization failed: {e}"));
                WMCP_ERROR
            }
        },
        Err(e) => {
            set_last_error(&e.to_string());
            WMCP_ERROR
        }
    }
}

/// List all visible windows with details as a JSON array.
///
/// # Safety
///
/// `out_json` must be a valid pointer. Caller must free with `wmcp_free_string()`.
#[no_mangle]
pub unsafe extern "C" fn wmcp_list_windows(out_json: *mut *mut c_char) -> i32 {
    if out_json.is_null() {
        set_last_error("out_json is null");
        return WMCP_ERROR;
    }
    match wmcp_core::window::list_windows() {
        Ok(windows) => match serde_json::to_string(&windows) {
            Ok(json) => match CString::new(json) {
                Ok(cstr) => {
                    unsafe { *out_json = cstr.into_raw() };
                    WMCP_OK
                }
                Err(e) => {
                    set_last_error(&format!("CString conversion failed: {e}"));
                    WMCP_ERROR
                }
            },
            Err(e) => {
                set_last_error(&format!("JSON serialization failed: {e}"));
                WMCP_ERROR
            }
        },
        Err(e) => {
            set_last_error(&e.to_string());
            WMCP_ERROR
        }
    }
}

/// Capture a screenshot as PNG bytes.
///
/// # Safety
///
/// `out_buf` must be a valid pointer to a `*mut u8`.
/// `out_len` must be a valid pointer to a `usize`.
/// On success, `*out_buf` is set to a heap-allocated buffer and `*out_len` to its length.
/// Caller must free the buffer with `wmcp_free_string()` (cast to `*mut c_char`).
#[no_mangle]
pub unsafe extern "C" fn wmcp_capture_screenshot_png(
    monitor_index: u32,
    out_buf: *mut *mut u8,
    out_len: *mut usize,
) -> i32 {
    if out_buf.is_null() || out_len.is_null() {
        set_last_error("null pointer argument");
        return WMCP_ERROR;
    }
    match wmcp_core::screenshot::capture_png(monitor_index) {
        Ok(png_bytes) => {
            let len = png_bytes.len();
            let boxed = png_bytes.into_boxed_slice();
            let ptr = Box::into_raw(boxed) as *mut u8;
            unsafe {
                *out_buf = ptr;
                *out_len = len;
            }
            WMCP_OK
        }
        Err(e) => {
            set_last_error(&e.to_string());
            WMCP_ERROR
        }
    }
}

/// Free a byte buffer allocated by `wmcp_capture_screenshot_png`.
///
/// # Safety
///
/// `ptr` must be a buffer returned by `wmcp_capture_screenshot_png` or null.
/// `len` must be the corresponding length.
#[no_mangle]
pub unsafe extern "C" fn wmcp_free_buffer(ptr: *mut u8, len: usize) {
    if !ptr.is_null() && len > 0 {
        drop(unsafe { Box::from_raw(std::ptr::slice_from_raw_parts_mut(ptr, len)) });
    }
}

/// Capture the UIA tree for window handles as a JSON string.
///
/// # Safety
///
/// `handles` must point to `handle_count` contiguous, initialized `isize` values.
/// `*out_json` will be set to a heap-allocated string; free with
/// `wmcp_free_string()`.
#[no_mangle]
pub unsafe extern "C" fn wmcp_capture_tree(
    handles: *const isize,
    handle_count: usize,
    max_depth: usize,
    out_json: *mut *mut c_char,
) -> i32 {
    if handles.is_null() || out_json.is_null() {
        set_last_error("null pointer argument");
        return WMCP_ERROR;
    }

    if handle_count == 0 {
        // Return empty JSON array for zero handles
        match CString::new("[]") {
            Ok(cstr) => {
                unsafe { *out_json = cstr.into_raw() };
                return WMCP_OK;
            }
            Err(_) => {
                set_last_error("CString allocation failed");
                return WMCP_ERROR;
            }
        }
    }

    if handle_count > MAX_HANDLE_COUNT {
        set_last_error(&format!(
            "handle_count {handle_count} exceeds maximum {MAX_HANDLE_COUNT}"
        ));
        return WMCP_ERROR;
    }

    // Validate pointer alignment
    if (handles as usize) % std::mem::align_of::<isize>() != 0 {
        set_last_error("handles pointer is not properly aligned");
        return WMCP_ERROR;
    }

    let handle_slice = unsafe { std::slice::from_raw_parts(handles, handle_count) };
    let snapshots = wmcp_core::tree::capture_tree_raw(handle_slice, max_depth);

    match serde_json::to_string(&snapshots) {
        Ok(json) => {
            // Sanitize null bytes that would break CString
            let json_sanitized = json.replace('\0', "\\u0000");
            match CString::new(json_sanitized) {
                Ok(cstr) => {
                    unsafe { *out_json = cstr.into_raw() };
                    WMCP_OK
                }
                Err(e) => {
                    set_last_error(&format!("CString conversion failed: {e}"));
                    WMCP_ERROR
                }
            }
        },
        Err(e) => {
            set_last_error(&format!("JSON serialization failed: {e}"));
            WMCP_ERROR
        }
    }
}
