//! Python exception mapping for `windows_mcp_core`.
//!
//! All Rust-side failures are funnelled through [`WindowsMcpError`], which
//! implements [`From<WindowsMcpError> for pyo3::PyErr`] so that the `?`
//! operator can be used directly inside `#[pyfunction]` bodies.
//!
//! # Example
//!
//! ```rust,ignore
//! use crate::errors::WindowsMcpError;
//!
//! #[pyo3::pyfunction]
//! fn do_something() -> pyo3::PyResult<()> {
//!     some_fallible_op().map_err(|e| WindowsMcpError::SystemInfoError(e.to_string()))?;
//!     Ok(())
//! }
//! ```

use pyo3::exceptions::PyRuntimeError;
use pyo3::PyErr;
use windows::core::Error as WindowsError;

/// Top-level error type for the `windows_mcp_core` extension.
///
/// Each variant corresponds to a distinct subsystem so that Python callers
/// can `except RuntimeError` and inspect the message to determine the origin.
/// When the `windows` crate COM subsystem is added, a dedicated
/// `PyException` subclass per variant would allow more precise catching.
#[derive(Debug)]
pub enum WindowsMcpError {
    /// Failure while collecting system information via the `sysinfo` crate.
    SystemInfoError(String),

    /// COM / UIAutomation error (reserved for the future `windows-rs` tree
    /// traversal subsystem).
    ComError(String),

    /// Accessibility tree traversal or element lookup failure.
    TreeError(String),

    /// Input simulation failure (SendInput / keyboard / mouse).
    InputError(String),

    /// Screenshot capture failure (GDI / DXGI).
    ScreenshotError(String),
}

impl std::fmt::Display for WindowsMcpError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SystemInfoError(msg) => write!(f, "SystemInfoError: {msg}"),
            Self::ComError(msg) => write!(f, "ComError: {msg}"),
            Self::TreeError(msg) => write!(f, "TreeError: {msg}"),
            Self::InputError(msg) => write!(f, "InputError: {msg}"),
            Self::ScreenshotError(msg) => write!(f, "ScreenshotError: {msg}"),
        }
    }
}

impl std::error::Error for WindowsMcpError {}

/// Convert a [`windows::core::Error`] (COM / Win32 HRESULT failure) into a
/// [`WindowsMcpError::ComError`].
///
/// This allows `?` to be used on `windows-rs` fallible calls inside
/// functions that return `Result<_, WindowsMcpError>`.
///
/// # Example
///
/// ```rust,ignore
/// use windows::Win32::UI::Accessibility::IUIAutomation;
/// use crate::errors::WindowsMcpError;
///
/// fn get_root(uia: &IUIAutomation) -> Result<(), WindowsMcpError> {
///     let _root = unsafe { uia.GetRootElement() }?;  // From<windows::core::Error>
///     Ok(())
/// }
/// ```
impl From<WindowsError> for WindowsMcpError {
    fn from(err: WindowsError) -> Self {
        WindowsMcpError::ComError(format!("Windows COM error: {err}"))
    }
}

/// Convert any [`WindowsMcpError`] into a Python `RuntimeError`.
///
/// PyO3 requires `From<E> for PyErr` so that `map_err` / `?` in
/// `#[pyfunction]` bodies can lift Rust errors into Python exceptions
/// automatically.
impl From<WindowsMcpError> for PyErr {
    fn from(err: WindowsMcpError) -> Self {
        PyRuntimeError::new_err(err.to_string())
    }
}
