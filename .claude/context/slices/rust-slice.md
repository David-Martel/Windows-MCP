# Rust Migration Context Slice -- Windows-MCP

## Recommended Hybrid Architecture
- Python FastMCP stays as protocol/orchestration layer
- Rust PyO3 extension (`windows_mcp_core.pyd`) handles hot paths
- Build via Maturin, separate sub-crate with own Cargo.toml

## Primary Rust Target: Tree Traversal
- `capture_tree(window_handles: list[int], options: dict) -> list[dict]`
- Uses `windows-rs` IUIAutomation COM bindings (zero-cost vtable calls)
- `rayon` for parallel per-window traversal (each thread: own STA apartment)
- `py.allow_threads()` releases GIL during entire traversal
- Estimated: 200-800ms (Python optimized) -> 50-200ms (Rust)

## Secondary Targets
- Screenshot: DXGI Output Duplication + `image` crate (55-150ms -> 12-35ms)
- Win32 operations: Toast notifications (WinRT), locale, OS version, app launch

## Key Crates
- `windows` v0.62+ (Microsoft official, `Win32::UI::Accessibility`)
- `uiautomation-rs` v0.24+ (higher-level wrapper, `!Send + !Sync` by design)
- `pyo3` v0.23+ with `extension-module` feature
- `rayon` v1.10+ for parallel iterators
- `image` v0.25+ for screenshot annotation

## COM Threading Risk
- Python process has STA apartment on main thread (via comtypes `CoInitialize`)
- Rust extension threads must call `CoInitializeEx(COINIT_MULTITHREADED)` independently
- `uiautomation-rs` UIAutomation is `!Send + !Sync` -- one instance per thread required
- Cleanest option: dedicated Rust threads (not rayon pool) with explicit COM init

## Build Integration
- Current: hatchling build backend
- Recommended: Keep hatchling for Python, separate Rust sub-crate with Maturin
- Install Rust extension as optional dependency

## Port Complexity
- `tree_traversal` is ~400 lines with complex branching (browser detection, DOM mode, role filtering)
- Must replicate classification logic exactly (interactive/scrollable/informative)
- Mitigation: shadow mode -- run both implementations, diff outputs per-window
