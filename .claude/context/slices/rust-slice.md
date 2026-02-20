# Rust Context Slice -- Windows-MCP (2026-02-20)

## Current Rust Workspace (native/)

4-crate workspace, all compiling with zero warnings:

```
wmcp-core/     Pure Rust lib (NO PyO3)
  src/lib.rs, errors.rs, com.rs, system_info.rs, input.rs,
  tree/, screenshot.rs, query.rs, pattern.rs
wmcp-pyo3/     PyO3 wrappers -> windows_mcp_core.pyd (24 functions)
wmcp-ffi/      C ABI DLL -> windows_mcp_ffi.dll (12 exports)
wmcp-cli/      CLI binaries (wmcp-worker: JSON-RPC over stdin/stdout)
```

## Implemented Modules

| Module | Functions | Status |
|--------|-----------|--------|
| system_info | system_info() | Complete, Rust fast-path in Desktop.get_system_info() |
| input | send_text_raw, send_key_raw, send_hotkey_raw, send_click_raw, send_mouse_move_raw, send_scroll_raw, send_drag_raw | Complete, wired into InputService |
| tree | capture_tree_raw | Complete, wired into Tree.get_window_wise_nodes() |
| screenshot | capture_raw, capture_png | Complete, DXGI Output Duplication |
| query | element_from_point, find_elements, get_screen_metrics | NEW -- wired into Find tool |
| pattern | invoke_at, toggle_at, set_value_at, expand_at, collapse_at, select_at | NEW -- wired into Invoke tool |

## Build Patterns
- **sccache port 4226 blocked**: Use `RUSTC_WRAPPER=""` to bypass
- **PYO3_PYTHON**: Must point to `.venv/Scripts/python.exe`
- **Install**: cargo build, then copy .pyd/.dll to `.venv/Lib/site-packages/`
- **Target dir**: `T:\RustCache\cargo-target\release/`

## Key Crates
- `windows` v0.62+ (Win32_UI_Accessibility, Win32_Graphics_Gdi, etc.)
- `pyo3` v0.23+ with extension-module feature
- `rayon` for parallel tree traversal
- `sysinfo` for system_info
- `thiserror` for error types

## COM Threading
- COMGuard RAII pattern: `CoInitializeEx(COINIT_MULTITHREADED)` per thread
- `py.allow_threads()` releases GIL during all COM operations
- Each function creates its own `CUIAutomation` instance (no sharing)

## Rust Unit Tests (25 total)
- query.rs: FindCriteria defaults, ScreenMetrics struct, ElementInfo serde
- pattern.rs: PatternResult serde, toggle state names
- input.rs: empty string, too-long text, empty/too-many hotkeys, normalise_coords
- tree/mod.rs: control_type_name mapping, empty handles, zero handle, max depth

## Lessons Learned
- ctypes `c_char_p` causes double-free with `wmcp_free_string` -- use `c_void_p`
- sysinfo CPU: needs double-refresh + 200ms gap for non-zero readings
- SM_CXVIRTUALSCREEN (not SM_CXSCREEN) for SendInput MOUSEEVENTF_ABSOLUTE
- OnceLock bad for screen dims (resolution can change) -- call GetSystemMetrics each time
- `windows::core::Interface` trait required for `.cast()` on COM objects
- `UIA_PATTERN_ID` is a newtype wrapper -- use `UIA_InvokePatternId.0` for raw i32
- `CreatePropertyCondition` returns `IUIAutomationPropertyCondition` -- must `.cast::<IUIAutomationCondition>()` for `FindAll`
