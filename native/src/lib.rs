//! `windows_mcp_core` -- PyO3 Rust extension for Windows-MCP.
//!
//! This crate provides a native Python extension module that accelerates the
//! hot paths in the Windows-MCP Python server.  The goal is to keep the MCP
//! protocol layer in Python (FastMCP) while offloading CPU-bound and
//! OS-interaction work to Rust via PyO3 / Maturin.
//!
//! # Modules
//!
//! | Module | Purpose |
//! |--------|---------|
//! | [`errors`] | [`WindowsMcpError`] enum and `From<> for PyErr` impl |
//! | [`system_info`] | Replace PowerShell subprocess calls with `sysinfo` |
//!
//! # Planned modules (Phase 3)
//!
//! - `tree` -- UIA accessibility tree traversal via `windows-rs` + Rayon
//! - `screenshot` -- DXGI Output Duplication capture
//! - `input` -- `SendInput` (keyboard/mouse) replacing pyautogui
//!
//! # Building
//!
//! ```bash
//! # From the `native/` directory:
//! maturin develop --release          # install into the active venv
//! maturin build --release            # produce a wheel in target/wheels/
//! ```
//!
//! # Usage (Python)
//!
//! ```python
//! import windows_mcp_core
//!
//! info = windows_mcp_core.system_info()
//! print(info["os_name"])       # e.g. "Windows 11 Pro"
//! print(info["cpu_count"])     # e.g. 16
//! print(info["disks"])         # list of dicts
//! ```

pub mod errors;
pub mod system_info;

use pyo3::prelude::*;

/// Register the `windows_mcp_core` Python module.
///
/// All public `#[pyfunction]` items from submodules are added here so that
/// they are available at the top level:
///
/// ```python
/// import windows_mcp_core
/// info = windows_mcp_core.system_info()
/// ```
///
/// Submodule objects are also added for dot-access and for future expansion:
///
/// ```python
/// # (future) import windows_mcp_core.tree
/// ```
#[pymodule]
fn windows_mcp_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // -----------------------------------------------------------------------
    // Top-level functions
    // -----------------------------------------------------------------------

    // system_info: replaces 200-500 ms PowerShell subprocess calls
    m.add_function(wrap_pyfunction!(system_info::system_info, m)?)?;

    // -----------------------------------------------------------------------
    // Module metadata
    // -----------------------------------------------------------------------
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add("__doc__", "Native Rust acceleration layer for Windows-MCP.")?;

    Ok(())
}
