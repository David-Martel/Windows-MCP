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
/// Returns a pointer valid until the next wmcp_* call on this thread.
/// Returns null if no error has occurred.
#[no_mangle]
pub extern "C" fn wmcp_last_error() -> *const c_char {
    LAST_ERROR.with(|e| {
        e.borrow()
            .as_ref()
            .map(|s| s.as_ptr())
            .unwrap_or(ptr::null())
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

    let count = wmcp_core::input::send_text_raw(text_str);
    if !out_count.is_null() {
        unsafe { *out_count = count };
    }
    WMCP_OK
}

/// Click the mouse at absolute screen coordinates.
#[no_mangle]
pub extern "C" fn wmcp_send_click(x: i32, y: i32, button: i32) -> i32 {
    let button_str = match button {
        1 => "right",
        2 => "middle",
        _ => "left",
    };
    wmcp_core::input::send_click_raw(x, y, button_str);
    WMCP_OK
}

/// Capture the UIA tree for window handles as a JSON string.
///
/// # Safety
///
/// `handles` must point to `handle_count` valid `isize` values.
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

    let handle_slice = unsafe { std::slice::from_raw_parts(handles, handle_count) };
    let snapshots = wmcp_core::tree::capture_tree_raw(handle_slice, max_depth);

    match serde_json::to_string(&snapshots) {
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
    }
}
