//! Error types for `wmcp_core`.
//!
//! All Rust-side failures are funnelled through [`WindowsMcpError`], which
//! uses `thiserror` for `Display` and `Error` derives.  PyO3 conversion
//! is handled in the `wmcp-pyo3` crate, keeping this crate PyO3-free.

use thiserror::Error;
use windows::core::Error as WindowsError;

/// Top-level error type for the `wmcp_core` library.
///
/// Each variant corresponds to a distinct subsystem.
#[derive(Debug, Error)]
pub enum WindowsMcpError {
    /// Failure while collecting system information via the `sysinfo` crate.
    #[error("SystemInfoError: {0}")]
    SystemInfoError(String),

    /// COM / UIAutomation error.
    #[error("ComError: {0}")]
    ComError(String),

    /// Accessibility tree traversal or element lookup failure.
    #[error("TreeError: {0}")]
    TreeError(String),

    /// Input simulation failure (SendInput / keyboard / mouse).
    #[error("InputError: {0}")]
    InputError(String),

    /// Screenshot capture failure (GDI / DXGI).
    #[error("ScreenshotError: {0}")]
    ScreenshotError(String),
}

/// Convert a `windows::core::Error` (COM / Win32 HRESULT failure) into a
/// `WindowsMcpError::ComError`.
impl From<WindowsError> for WindowsMcpError {
    fn from(err: WindowsError) -> Self {
        WindowsMcpError::ComError(format!("Windows COM error: {err}"))
    }
}
