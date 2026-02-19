//! Window enumeration and management via Win32 API.
//!
//! Provides Rust-native implementations of window operations that currently
//! require Python `win32gui` or ctypes calls.  All functions return owned
//! structs, never raw handles.

use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;

use serde::Serialize;
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, RECT, TRUE};
use windows::Win32::Graphics::Dwm::{DwmGetWindowAttribute, DWMWA_CLOAKED};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetClassNameW, GetForegroundWindow, GetWindow, GetWindowLongW, GetWindowRect,
    GetWindowTextLengthW, GetWindowTextW, GetWindowThreadProcessId, IsIconic, IsWindow,
    IsWindowVisible, IsZoomed, GWL_EXSTYLE, GWL_STYLE, GW_OWNER, WS_EX_APPWINDOW,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_VISIBLE,
};

use crate::errors::WindowsMcpError;

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

/// Owned snapshot of a visible window.
#[derive(Debug, Clone, Serialize)]
pub struct WindowInfo {
    pub hwnd: isize,
    pub title: String,
    pub class_name: String,
    pub pid: u32,
    pub rect: WindowRect,
    pub is_minimized: bool,
    pub is_maximized: bool,
    pub is_visible: bool,
}

/// Window bounding rectangle in screen coordinates.
#[derive(Debug, Clone, Serialize)]
pub struct WindowRect {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Read the window title (up to 512 chars).
fn read_window_title(hwnd: HWND) -> String {
    let len = unsafe { GetWindowTextLengthW(hwnd) };
    if len <= 0 {
        return String::new();
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    let copied = unsafe { GetWindowTextW(hwnd, &mut buf) };
    if copied <= 0 {
        return String::new();
    }
    OsString::from_wide(&buf[..copied as usize])
        .to_string_lossy()
        .into_owned()
}

/// Read the window class name (up to 256 chars).
fn read_class_name(hwnd: HWND) -> String {
    let mut buf = [0u16; 256];
    let len = unsafe { GetClassNameW(hwnd, &mut buf) };
    if len <= 0 {
        return String::new();
    }
    OsString::from_wide(&buf[..len as usize])
        .to_string_lossy()
        .into_owned()
}

/// Get the process ID for a window handle.
fn read_pid(hwnd: HWND) -> u32 {
    let mut pid: u32 = 0;
    unsafe { GetWindowThreadProcessId(hwnd, Some(&mut pid)) };
    pid
}

/// Check if a window is cloaked (hidden by DWM).
///
/// On Windows 10/11, UWP apps and windows on other virtual desktops are
/// "cloaked" -- they pass `IsWindowVisible` but are invisible to the user.
fn is_cloaked(hwnd: HWND) -> bool {
    let mut cloaked: u32 = 0;
    let hr = unsafe {
        DwmGetWindowAttribute(
            hwnd,
            DWMWA_CLOAKED,
            &mut cloaked as *mut u32 as *mut _,
            std::mem::size_of::<u32>() as u32,
        )
    };
    hr.is_ok() && cloaked != 0
}

/// Check if a window has a visible owner (popup/dialog filter).
///
/// Windows owned by other visible windows should not appear in Alt+Tab
/// unless they have WS_EX_APPWINDOW style.
fn has_visible_owner(hwnd: HWND) -> bool {
    let owner = unsafe { GetWindow(hwnd, GW_OWNER) };
    match owner {
        Ok(h) if !h.is_invalid() => unsafe { IsWindowVisible(h) }.as_bool(),
        _ => false,
    }
}

/// Check if a window is a normal top-level application window (not a tool
/// window, cloaked, or otherwise invisible to the taskbar).
///
/// Implements the canonical Alt+Tab filter (Raymond Chen):
/// 1. IsWindowVisible
/// 2. Not WS_EX_TOOLWINDOW
/// 3. Not cloaked (DwmGetWindowAttribute DWMWA_CLOAKED)
/// 4. No visible owner unless WS_EX_APPWINDOW
fn is_alt_tab_window(hwnd: HWND) -> bool {
    let style = unsafe { GetWindowLongW(hwnd, GWL_STYLE) } as u32;
    let ex_style = unsafe { GetWindowLongW(hwnd, GWL_EXSTYLE) } as u32;

    // Must be visible
    if style & WS_VISIBLE.0 == 0 {
        return false;
    }

    // Skip tool windows and non-activatable windows
    if ex_style & WS_EX_TOOLWINDOW.0 != 0 {
        return false;
    }
    if ex_style & WS_EX_NOACTIVATE.0 != 0 {
        return false;
    }

    // Skip cloaked windows (other virtual desktops, suspended UWP apps)
    if is_cloaked(hwnd) {
        return false;
    }

    // Skip windows with a visible owner, unless marked as WS_EX_APPWINDOW
    if has_visible_owner(hwnd) && (ex_style & WS_EX_APPWINDOW.0 == 0) {
        return false;
    }

    true
}

/// Callback for EnumWindows that collects visible window handles.
///
/// # Safety
///
/// `EnumWindows` calls this callback synchronously on the calling thread.
/// The raw pointer targets the `Vec` struct on the caller's stack frame,
/// so even if the `Vec` reallocates its backing buffer, the pointer
/// remains valid for the duration of the enumeration.
unsafe extern "system" fn enum_callback(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let handles = unsafe { &mut *(lparam.0 as *mut Vec<HWND>) };

    if unsafe { IsWindowVisible(hwnd) }.as_bool() && is_alt_tab_window(hwnd) {
        // Skip windows with no title
        let title_len = unsafe { GetWindowTextLengthW(hwnd) };
        if title_len > 0 {
            handles.push(hwnd);
        }
    }

    TRUE // continue enumeration
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Enumerate all visible top-level windows.
///
/// Returns a list of window handles for windows that are visible, have a
/// title, and appear in the Alt+Tab list (not tool windows, not cloaked).
pub fn enumerate_visible_windows() -> Result<Vec<isize>, WindowsMcpError> {
    let mut handles: Vec<HWND> = Vec::with_capacity(64);
    let result = unsafe {
        EnumWindows(
            Some(enum_callback),
            LPARAM(&mut handles as *mut Vec<HWND> as isize),
        )
    };

    result.map_err(|e| {
        WindowsMcpError::ComError(format!("EnumWindows failed: {e}"))
    })?;

    Ok(handles.iter().map(|h| h.0 as isize).collect())
}

/// Get detailed information about a window by its handle.
///
/// Returns an error if the handle is invalid or the window has been destroyed.
pub fn get_window_info(handle: isize) -> Result<WindowInfo, WindowsMcpError> {
    let hwnd = HWND(handle as *mut core::ffi::c_void);

    // Validate that the handle refers to an existing window
    if !unsafe { IsWindow(hwnd) }.as_bool() {
        return Err(WindowsMcpError::ComError(format!(
            "Invalid window handle: {handle}"
        )));
    }

    let title = read_window_title(hwnd);
    let class_name = read_class_name(hwnd);
    let pid = read_pid(hwnd);

    let mut rect_raw = RECT::default();
    unsafe { GetWindowRect(hwnd, &mut rect_raw) }.map_err(|e| {
        WindowsMcpError::ComError(format!("GetWindowRect failed for handle {handle}: {e}"))
    })?;

    let is_minimized = unsafe { IsIconic(hwnd) }.as_bool();
    let is_maximized = unsafe { IsZoomed(hwnd) }.as_bool();
    let is_visible = unsafe { IsWindowVisible(hwnd) }.as_bool();

    Ok(WindowInfo {
        hwnd: handle,
        title,
        class_name,
        pid,
        rect: WindowRect {
            left: rect_raw.left,
            top: rect_raw.top,
            right: rect_raw.right,
            bottom: rect_raw.bottom,
        },
        is_minimized,
        is_maximized,
        is_visible,
    })
}

/// Get the foreground (active) window handle.
///
/// Returns 0 if no window is in the foreground.
pub fn get_foreground_hwnd() -> isize {
    let hwnd = unsafe { GetForegroundWindow() };
    hwnd.0 as isize
}

/// Get information about all visible windows.
///
/// Convenience function that enumerates windows and collects info for each.
/// Windows that become invalid between enumeration and info-gathering are
/// silently skipped (TOCTOU race inherent to Win32 window enumeration).
pub fn list_windows() -> Result<Vec<WindowInfo>, WindowsMcpError> {
    let handles = enumerate_visible_windows()?;
    let mut windows = Vec::with_capacity(handles.len());
    for handle in handles {
        match get_window_info(handle) {
            Ok(info) => windows.push(info),
            Err(_) => continue, // skip inaccessible windows (TOCTOU race)
        }
    }
    Ok(windows)
}
