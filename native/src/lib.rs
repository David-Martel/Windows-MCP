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
//! | [`input`] | `SendInput` keyboard/mouse simulation (replaces pyautogui) |
//! | [`system_info`] | Replace PowerShell subprocess calls with `sysinfo` |
//! | [`tree`] | UIA accessibility tree traversal via `windows-rs` + Rayon |
//!
//! # Planned modules
//!
//! - `screenshot` -- DXGI Output Duplication capture
//!
//! # Building
//!
//! ```bash
//! # From the `native/` directory:
//! RUSTC_WRAPPER="" cargo build --release
//! # Then copy target/release/windows_mcp_core.dll to
//! # .venv/Lib/site-packages/windows_mcp_core.pyd
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
//!
//! import ctypes
//! hwnd = ctypes.windll.user32.GetForegroundWindow()
//! trees = windows_mcp_core.capture_tree([hwnd], max_depth=10)
//! print(trees[0]["name"], trees[0]["control_type"])
//! ```

pub mod errors;
pub mod input;
pub mod system_info;
pub mod tree;

use pyo3::prelude::*;

/// Register the `windows_mcp_core` Python module.
///
/// All public `#[pyfunction]` items from submodules are added here so that
/// they are available at the top level:
///
/// ```python
/// import windows_mcp_core
///
/// info = windows_mcp_core.system_info()
///
/// trees = windows_mcp_core.capture_tree([hwnd])
/// ```
#[pymodule]
fn windows_mcp_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // -----------------------------------------------------------------------
    // Top-level functions
    // -----------------------------------------------------------------------

    // system_info: replaces 200-500 ms PowerShell subprocess calls
    m.add_function(wrap_pyfunction!(system_info::system_info, m)?)?;

    // capture_tree: single-RPC UIA subtree capture with Rayon parallelism.
    // Replaces the Python comtypes per-node BuildUpdatedCache loop.
    m.add_function(wrap_pyfunction!(tree::capture_tree, m)?)?;

    // input: SendInput-based keyboard/mouse simulation (replaces pyautogui).
    m.add_function(wrap_pyfunction!(input::send_text, m)?)?;
    m.add_function(wrap_pyfunction!(input::send_key, m)?)?;
    m.add_function(wrap_pyfunction!(input::send_click, m)?)?;
    m.add_function(wrap_pyfunction!(input::send_mouse_move, m)?)?;
    m.add_function(wrap_pyfunction!(input::send_hotkey, m)?)?;

    // -----------------------------------------------------------------------
    // Module metadata
    // -----------------------------------------------------------------------
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add("__doc__", "Native Rust acceleration layer for Windows-MCP.")?;

    Ok(())
}
