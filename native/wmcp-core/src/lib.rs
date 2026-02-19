//! `wmcp_core` -- Pure Rust core library for Windows-MCP.
//!
//! This crate contains all business logic with **no PyO3 dependency**.
//! It can be consumed by:
//! - `wmcp-pyo3` (PyO3 Python extension)
//! - `wmcp-ffi` (C ABI DLL for ctypes / other languages)
//! - `wmcp-cli` (standalone CLI tools)
//!
//! # Modules
//!
//! | Module | Purpose |
//! |--------|---------|
//! | [`errors`] | `WindowsMcpError` enum via `thiserror` |
//! | [`com`] | `COMGuard` RAII wrapper for COM apartment init |
//! | [`system_info`] | System telemetry via `sysinfo` crate |
//! | [`input`] | `SendInput` keyboard/mouse simulation |
//! | [`tree`] | UIA accessibility tree traversal via `windows-rs` + Rayon |

pub mod com;
pub mod errors;
pub mod input;
pub mod system_info;
pub mod tree;
