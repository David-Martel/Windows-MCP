# Windows-MCP Phase 2 Execution Plan

**Date:** 2026-02-19
**Status:** Proposed
**Prerequisite:** Phase 1 complete (auth, perf fixes, PyO3 scaffold, Build.ps1)

---

## Context

Session 3 completed Python quick wins, auth system, and PyO3 scaffold. 184 tests passing.
PC-AI exploration revealed reusable Rust patterns: rayon parallelism, OnceLock singletons,
FFI string marshaling, streaming search, and error status codes.

**Remaining high-impact items** from TODO.md (sorted by value/effort):

| Item | Impact | Effort | Category |
|------|--------|--------|----------|
| TreeScope_Subtree optimization | 60-80% tree speedup | 1 day | Python |
| Rust tree traversal | 10-50x tree speedup | 1 week | Rust |
| WaitFor/Find/Invoke tools | New capabilities | 2-3 days | Python |
| COM threading fixes | Eliminates UB | 1 day | Python |
| Desktop God Object decomposition | Maintainability | 2-3 days | Architecture |
| Security hardening (SSRF, shell) | Safety | 1-2 days | Security |
| 85% test coverage | Quality | 2-3 days | Testing |

---

## Phase 2A: Python Optimization (No Rust Changes)

### Step 1: TreeScope_Subtree CacheRequest

**Files:** `src/windows_mcp/tree/service.py`, `tree/cache_utils.py`

**Change:** Replace per-node `BuildUpdatedCache(TreeScope_Element | TreeScope_Children)` with
single `BuildUpdatedCache(TreeScope_Subtree)` on window root, then walk the cached result in-process.

**Impact:** Collapses thousands of COM round-trips into one per window. 60-80% tree traversal reduction.

**Risk:** May change element ordering. Needs shadow-mode validation (run both paths, diff outputs).

**Verification:**
- Add shadow-mode test: compare element lists from old vs new traversal
- Measure timing with `time.perf_counter()` before/after
- Run on complex apps (VS Code, Chrome, Explorer) to verify completeness

### Step 2: LegacyIAccessiblePattern Dedup

**Files:** `src/windows_mcp/tree/service.py` (~lines 439-482)

**Change:** Call `GetLegacyIAccessiblePattern()` once per element, store in local variable.
Currently called up to 3x per interactive element (live COM round-trip each time).

**Impact:** 10-20% reduction in per-element overhead.

### Step 3: COM Threading Fix

**Files:** `src/windows_mcp/tree/service.py`, `desktop/service.py`

**Changes:**
- Make `self.dom` and `self.dom_bounding_box` local to traversal context (not instance fields)
- Add `threading.Lock` for `Tree.tree_state` and `Desktop.desktop_state`
- Fix `_AutomationClient.instance()` TOCTOU with a lock

### Step 4: New MCP Tools (WaitFor, Find, Invoke)

**Files:** `src/windows_mcp/__main__.py`, `desktop/service.py`

**WaitFor tool:**
```python
@mcp.tool(name="WaitFor")
async def waitfor_tool(mode: Literal["window", "element"], name: str, timeout: int = 10):
    # Uses uia.Exists(maxSearchSeconds=timeout) for elements
    # Uses watchdog focus events for windows
```

**Find tool:**
```python
@mcp.tool(name="Find")
async def find_tool(name: str = None, control_type: str = None,
                    window: str = None, automation_id: str = None):
    # Returns coordinates + xpath without full Snapshot
```

**Invoke tool:**
```python
@mcp.tool(name="Invoke")
async def invoke_tool(element_index: int, action: str = "invoke"):
    # InvokePattern.Invoke() for buttons
    # TogglePattern.Toggle() for checkboxes
    # ValuePattern.SetValue() for text fields
```

---

## Phase 2B: Rust Acceleration (Borrowing PC-AI Patterns)

### Step 5: Expand windows_mcp_core with Tree Traversal

**Files:** `native/src/tree/mod.rs`, `native/src/tree/cache.rs`, `native/Cargo.toml`

**Architecture** (adapted from PC-AI):
```
native/src/
├── lib.rs              # PyO3 module (add capture_tree)
├── errors.rs           # WindowsMcpError (add ComError variants)
├── system_info.rs      # Already done
└── tree/
    ├── mod.rs          # capture_tree() entry point
    ├── cache.rs        # CacheRequest with TreeScope_Subtree
    ├── walker.rs       # Recursive tree walking (adapted from PC-AI search/walker.rs)
    └── element.rs      # TreeElementNode serialization
```

**Key patterns from PC-AI to adopt:**
1. **Rayon parallelism** for per-window traversal (from `search/duplicates.rs`)
2. **OnceLock<Mutex<>>** for COM instance caching (from `telemetry/mod.rs`)
3. **Streaming element collection** (from `search/content.rs` buffered reader pattern)
4. **Error status mapping** (from `error.rs` `PcaiStatus` enum -> extend `WindowsMcpError`)

**COM threading strategy:**
```rust
// Each rayon thread gets its own COM apartment + UIAutomation instance
thread_local! {
    static UIA: RefCell<Option<IUIAutomation>> = RefCell::new(None);
}

fn get_uia() -> &IUIAutomation {
    UIA.with(|cell| {
        if cell.borrow().is_none() {
            unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).ok(); }
            let uia: IUIAutomation = unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) }.unwrap();
            *cell.borrow_mut() = Some(uia);
        }
        cell.borrow().as_ref().unwrap()
    })
}
```

**Cargo.toml additions:**
```toml
[dependencies]
windows = { version = "0.58", features = [
    "Win32_UI_Accessibility",
    "Win32_System_Com",
    "Win32_Foundation",
] }
rayon = "1.10"
```

**PyO3 entry point:**
```rust
#[pyfunction]
fn capture_tree(py: Python, window_handles: Vec<isize>, options: &Bound<PyDict>) -> PyResult<Vec<PyObject>> {
    py.allow_threads(|| {
        window_handles.par_iter().map(|&hwnd| {
            let uia = get_uia();
            let element = uia.ElementFromHandle(HWND(hwnd))?;
            let cache_request = build_subtree_cache_request(uia)?;
            let cached = element.BuildUpdatedCache(&cache_request)?;
            walk_cached_tree(cached)
        }).collect()
    })
}
```

**Impact:** 500-5000ms -> 50-200ms (10-25x speedup)

### Step 6: Rust Screenshot Module

**Files:** `native/src/screenshot.rs`

**Approach:** DXGI Output Duplication (from performance-slice recommendations)

```rust
#[pyfunction]
fn capture_screenshot(py: Python, monitor: i32, scale: f64) -> PyResult<Vec<u8>> {
    py.allow_threads(|| {
        // DXGI Output Duplication API
        // Returns PNG bytes
    })
}
```

**Impact:** 55-150ms -> 12-35ms

### Step 7: Build Pipeline Update

Update `Build.ps1` to handle the expanded Rust crate:
- Add `cargo test` for native/ (tree traversal tests with mock UIA)
- Add `cargo clippy --all-targets -- -D warnings`
- Add benchmarks: `cargo bench` with criterion

---

## Phase 2C: Security & Quality

### Step 8: SSRF Protection in Scrape Tool

**File:** `src/windows_mcp/desktop/service.py` (scrape method)

**Change:** Validate URL scheme (http/https only), block private IPs, block metadata endpoints.

### Step 9: Shell Tool Sandboxing

**File:** `src/windows_mcp/desktop/service.py` (execute_command method)

**Change:** Add command allowlist/blocklist with configurable policy via env vars.

### Step 10: Test Coverage to 85%

**New test files needed:**
- `tests/test_mcp_tools.py` -- Mock Desktop, test all 19 tool return formats
- `tests/test_watchdog.py` -- Test threading, event dispatch
- `tests/test_desktop_orchestration.py` -- Test get_state composition

---

## Phase 2D: Architecture (If Time Permits)

### Step 11: Decompose Desktop God Object

Extract into:
- `services/input_service.py` -- click, type, scroll, drag, move, shortcut
- `services/window_service.py` -- get_windows, switch_app, resize_app
- `services/registry_service.py` -- registry_get/set/delete/list
- `services/screen_service.py` -- screenshots, annotations

### Step 12: PC-AI Integration Bridge

**Concept:** Share `windows_mcp_core` Rust crate with PC-AI's `pcai_core_lib`:
- Extract common `windows-sys` bindings into shared workspace member
- Expose PC-AI's file search/duplicate detection as MCP tools
- Share sysinfo singleton and telemetry infrastructure

```
shared-workspace/
├── Cargo.toml (workspace)
├── windows-mcp-core/     # MCP-specific (tree, screenshot, input)
├── pcai-core/             # PC-AI-specific (search, diagnostics)
└── common/                # Shared (sysinfo, error, string, path)
```

---

## Commit Strategy

1. `perf: implement TreeScope_Subtree optimization` (Step 1-2)
2. `fix: resolve COM threading violations` (Step 3)
3. `feat: add WaitFor, Find, Invoke tools` (Step 4)
4. `feat(native): add Rust tree traversal module` (Step 5)
5. `feat(native): add Rust screenshot capture` (Step 6)
6. `security: add SSRF protection and shell sandboxing` (Step 8-9)
7. `test: expand coverage to 85%` (Step 10)
8. `refactor: decompose Desktop into focused services` (Step 11)

---

## Verification Checklist

After each step:
1. `uv run python -m pytest tests/` -- all tests pass
2. `ruff check . && ruff format --check .` -- no new lint errors
3. `.\Build.ps1 -Action Check` -- Rust checks pass
4. `.\Build.ps1 -Action Test` -- full suite including native tests
5. Manual test: `uv run windows-mcp` launches without errors

---

## Risk Assessment

| Risk | Mitigation |
|------|-----------|
| TreeScope_Subtree changes element ordering | Shadow mode: run both, diff |
| COM threading in Rust extension | thread_local! with explicit CoInitializeEx |
| sccache port conflict on build machine | Build.ps1 Find-NativeDll fallback |
| Large refactor breaks existing tools | Incremental extraction, test after each service |
| PC-AI workspace merge complexity | Shared crate as optional dependency first |
