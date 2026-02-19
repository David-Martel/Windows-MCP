# Windows-MCP TODO

**Generated:** 2026-02-18 from REVIEW.md findings
**Last updated:** 2026-02-19 -- 1770 tests, 64% coverage, Rust workspace (4 crates, 15 PyO3 + 12 FFI exports), tools/ decomposition complete
**Reference:** See [REVIEW.md](REVIEW.md) for full context on each item.

---

## P0 -- Must Fix Before Any Shared Deployment

- [x] **[S2] Add authentication for SSE/HTTP transport** -- Bearer token auth via `BearerAuthMiddleware`. DPAPI key storage via `AuthKeyManager`. CLI: `--api-key`, `--generate-key`, `--rotate-key`. Refuses `0.0.0.0` without auth.
- [x] **[A4] Fix analytics decorator binding bug** -- Changed `with_analytics` to accept a callable factory (`lambda: analytics`). Resolves at call time, not decoration time. Also fixed [Q5] traceback formatting.
- [x] **[P4] Fix PIL ImageDraw thread safety** -- Removed `ThreadPoolExecutor` from `get_annotated_screenshot`. Now draws sequentially (<5ms).
- [x] **[A5] Remove `ipykernel` from dependencies** -- Removed from `pyproject.toml` dependencies.
- [x] **[Q6] Remove debug `print()` from analytics.py:97** -- Removed. Also fixed 3 `print()` calls in `watchdog/event_handlers.py` (replaced with `logger.debug()`).

---

## P1 -- High Impact Performance & Safety

- [x] **[P6] Cache Start Menu app list** -- Added thread-safe TTL cache (1 hour) to `get_apps_from_start_menu()` with double-checked locking pattern.
- [x] **[P5] Eliminate duplicate VDM desktop enumeration** -- Added `get_desktop_info()` that returns `(current, all)` from a single `_enumerate_desktops()` call. `get_state()` uses it.
- [x] **[P2] Bound ThreadPoolExecutor** -- Added `max_workers=min(8, os.cpu_count() or 4)` to tree traversal executor in `tree/service.py`.
- [x] **[P3] Convert tree_traversal to iterative** -- Replaced recursion with explicit stack (7-tuple frames). Eliminates Python's 1000-frame limit on complex UIs.
- [x] **[P9] Make pyautogui.PAUSE configurable** -- Reduced from 1.0s to 0.05s in both `desktop/service.py` and `__main__.py`. Saves 1-6s per input operation.
- [x] **[S7] Block SSRF in Scrape tool** -- Done. `ScraperService.validate_url()` blocks non-HTTP schemes, private/reserved IPs, DNS rebinding, cloud metadata endpoints. Extracted to `scraper/service.py`.
- [x] **[T1] Fix COM apartment threading violations** -- Verified: `dom_context` dict is already thread-local in `get_nodes()`. COM init/uninit properly scoped per-thread. No cross-apartment sharing.
- [x] **[T2] Add synchronization to shared mutable state** -- Added `threading.Lock` to `Tree._state_lock` and `Desktop._state_lock`. State built locally, then published atomically under lock.

---

## P2 -- Architectural Improvements

- [x] **[A1] Decompose Desktop God Object** -- All 8 services extracted (Desktop down from 1039 to ~680 lines):
  - [x] `RegistryService` -- registry_get/set/delete/list (`registry/service.py`)
  - [x] `ShellService` -- execute, check_blocklist, ps_quote (`shell/service.py`)
  - [x] `ScraperService` -- validate_url, scrape (`scraper/service.py`)
  - [x] `InputService` -- click, type, scroll, drag, move, shortcut, multi_select, multi_edit (`input/service.py`)
  - [x] `WindowService` -- get_windows, get_active_window, bring_window_to_top, auto_minimize (`window/service.py`)
  - [x] `ScreenService` -- get_screenshot, get_annotated_screenshot, get_screen_size, get_dpi_scaling (`screen/service.py`)
  - [x] `VisionService` -- LLM-powered screenshot analysis via OpenAI-compatible API (`vision/service.py`)
  - [x] `ProcessService` -- list_processes, kill_process with protected process blocklist (`process/service.py`)
- [ ] **[A3] Replace module globals with dependency injection** -- Use FastMCP's lifespan context to pass `desktop`, `watchdog`, `analytics` via `ctx.request_context.lifespan_context`.
- [x] **[A2] Decompose __main__.py** -- Extracted 23 tool registrations into `tools/` package: `input_tools.py` (9), `state_tools.py` (4), `system_tools.py` (10). Shared state via `tools/_state.py`. Reduced __main__.py from 1122 to 207 lines (81%).
- [x] **[P5] Parallelize get_state** -- VDM desktop query runs in parallel with window enumeration chain via ThreadPoolExecutor. Saves ~100-200ms per get_state call.
- [ ] **[Q1] Deduplicate UIA constants** -- Extract 13 shared constants from `uia/core.py`, `uia/controls.py`, `uia/patterns.py` into `uia/constants.py`.
- [ ] **[Q2] Extract boolean coercion utility** -- Replace 6+ instances of `x is True or (isinstance(x, str) and x.lower() == "true")` with a single `to_bool()` helper.

---

## P3 -- Security Hardening

- [x] **[S1] Implement Shell tool sandboxing** -- Done. `ShellService.check_blocklist()` with 16 regex patterns (format, diskpart, bcdedit, rm -rf, net user/localgroup, IEX+DownloadString, etc). Configurable via `WINDOWS_MCP_SHELL_BLOCKLIST` env var. Extracted to `shell/service.py`.
- [ ] **[S4] Add path-scoping to File tool** -- Define allowed directory scope. Validate resolved paths remain within scope.
- [ ] **[S3] Restrict Registry tool access** -- Deny security-sensitive keys (Run, RunOnce, Services, Policies, SAM). Require confirmation for write/delete.
- [ ] **[S5] Rotate PostHog API key** -- Move to environment variable. Disable GeoIP, disable exception auto-capture. Default telemetry to opt-in.
- [ ] **[S6] Enforce HTTPS for auth client** -- Change default from `http://` to `https://`. Add certificate verification. Reduce `__repr__` key exposure to 4 chars.
- [x] **[S8] Add protected process list** -- `ProcessService.is_protected()` blocks csrss, lsass, services, svchost, winlogon, MsMpEng, smss, wininit, system, registry. Extracted to `process/service.py`.
- [ ] **Implement security audit logging** -- Log all tool invocations with timestamps, parameters, results.
- [ ] **Implement rate limiting** -- Per-tool rate limits to prevent abuse.

---

## P4 -- Test Coverage

- [x] **[A7] Add MCP tool handler integration tests** -- Done. 145 headless integration tests in `test_mcp_integration.py` covering all 23 tools via `tool.fn()`. 4 tiers: server structure, tool dispatch, error handling, transport+auth.
- [ ] **Add WatchDog service tests** -- Test threading, COM event handling, callback dispatch.
- [ ] **Add Desktop.get_state orchestration tests** -- Test the composition of windows, tree state, screenshots, VDM.
- [ ] **Add tree_traversal unit tests** -- Test with mock UIA trees. Cover DOM/interactive/scrollable classification.
- [ ] **Achieve 85% overall test coverage** -- Currently at 64% overall (1671 tests). Testable modules at 82-100%. COM/UIA modules at 14-42% (require live desktop).

---

## P5 -- Code Quality

- [ ] **[Q3] Remove wildcard imports in controls.py** -- Replace `from .enums import *` with explicit imports.
- [x] **[Q4] Remove dead config** -- `STRUCTURAL_CONTROL_TYPE_NAMES` is actually used in tree/service.py -- not dead. Verified, no action needed.
- [x] **[Q5] Fix broken traceback in analytics.py:107** -- Now uses `traceback.format_exception()` for full stack trace string.
- [x] **[Q7] Fix resource leaks** -- Fixed HTTP response leak in `scrape()` (close intermediate redirect responses and final response in `finally` block). `auto_minimize` handle guard already present. COM init guard deferred (comtypes handles internally).
- [ ] **[P10] Cache COM pattern calls in tree traversal** -- Store `GetLegacyIAccessiblePattern()` result in local variable instead of calling 3 times per node.
- [ ] **[P10] Fix BuildUpdatedCache** -- Check for existing cached state before issuing round-trip in `get_cached_children`.
- [ ] **Use set literals in tree/config.py** -- Replace `set([...])` with `{...}`.
- [x] **[A6] Consolidate input simulation** -- Rust Win32 `SendInput` is primary path in `InputService` (click, type, scroll, drag, move, shortcut). Falls back to pyautogui when native unavailable.
- [ ] **Add type annotations to all public APIs**.
- [ ] **[T3] Fix singleton TOCTOU in _AutomationClient** -- Add threading lock to `instance()`.

---

## P1.5 -- Capability Gaps (from Framework Comparison)

*Source: Phase 2 comparative analysis (rust-pro, backend-architect, performance-engineer agents)*

- [ ] **Replace `Wait(duration)` with event-driven `WaitFor`** -- Extract the pattern already used in `App` tool (lines 299-311) into a general `WaitFor(mode="window"|"element", name=..., timeout=...)` tool backed by `uia.Exists(maxSearchSeconds=N)`. Deprecate fixed-sleep `Wait`.
- [ ] **Add `ValuePattern.SetValue()` as primary text input** -- In `type()` method, attempt `ValuePattern.SetValue(text)` before falling back to `pg.typewrite()`. Eliminates 20ms/char typing for all UIA-compliant text fields.
- [x] **Reduce `pg.PAUSE` from 1.0 to 0.05** -- Done. Changed in both `desktop/service.py:45` and `__main__.py:33`.
- [ ] **Add `Find` tool for semantic element lookup** -- `Find(name, control_type, window, automation_id)` resolves to coordinates + xpath without full Snapshot traversal.
- [ ] **Add `Invoke` tool for UIA pattern actions** -- `InvokePattern.Invoke()` for buttons, `TogglePattern.Toggle()` for checkboxes, `ExpandCollapsePattern.Expand()` for dropdowns. More reliable than coordinate clicking.
- [ ] **Use `TreeScope_Subtree` for single-shot cached traversal** -- Currently `BuildUpdatedCache` called per-node (TWO COM round-trips each). Single `TreeScope_Subtree` on window root would collapse thousands of COM calls into one. Estimated 60-80% tree traversal reduction.

---

## P2.5 -- Performance Caching Layer

- [ ] **Per-window tree cache with WatchDog invalidation** -- Cache TreeState per window handle. On focus change, only re-traverse changed windows. 80-95% reduction on repeated `get_state()` calls.
- [ ] **Cache Start Menu app list with filesystem watcher** -- Build app map once at init, monitor Start Menu folders via `ReadDirectoryChangesW`. Invalidate on change.
- [ ] **Element coordinate cache** -- Reuse `TreeElementNode.bounding_box` for click/type instead of re-querying XPath. Invalidate on window move/resize.
- [ ] **Cache VDM desktop list with TTL** -- `get_current_desktop()` calls `get_all_desktops()` redundantly. Cache with 5s TTL, invalidated by WatchDog.
- [x] **Replace `ImageGrab.grab` with Rust DXGI** -- Rust DXGI Output Duplication used as primary screenshot path in `ScreenService.get_screenshot()`. Falls back to ImageGrab then pyautogui.

---

## P3.5 -- Rust Acceleration (Cargo Workspace)

*Workspace: `native/` with 4 crates -- wmcp-core (pure Rust lib), wmcp-pyo3 (PyO3 bindings), wmcp-ffi (C ABI DLL), wmcp-cli (standalone tools). 15 PyO3 functions + 12 FFI exports. Clippy clean with `-D warnings`.*

- [x] **Rust tree traversal module** -- `windows_mcp_core.capture_tree(handles, max_depth)` via PyO3. Uses `windows-rs` IUIAutomation + `rayon` per-HWND parallel traversal. COMGuard RAII for per-thread COM init. Wired into Python via `native_capture_tree()`.
- [x] **Rust screenshot module** -- DXGI Output Duplication + GDI BitBlt fallback + `image` crate PNG encoding. Wired into `ScreenService.get_screenshot()` as fast-path. `capture_png()` and `capture_raw()` exposed via PyO3 and FFI.
- [x] **Rust input module** -- `send_text`, `send_click`, `send_key`, `send_mouse_move`, `send_hotkey`, `send_scroll`, `send_drag` via Win32 `SendInput`. Wired into `InputService` as fast-path with pyautogui fallback.
- [x] **Rust window module** -- `enumerate_visible_windows`, `get_window_info`, `get_foreground_window`, `list_windows` with Alt+Tab filter + DWM cloaked detection. Exposed via PyO3 and FFI.
- [x] **Rust system_info** -- Wired into `Desktop.get_system_info()` as fast-path (avoids 1s blocking `cpu_percent`). Falls back to psutil.
- [x] **Replace PowerShell registry with `winreg`** -- All 4 registry methods (`registry_get/set/delete/list`) rewritten to use `winreg` stdlib. 200-500ms saved per operation.
- [x] **Replace PowerShell sysinfo with Python stdlib** -- `get_windows_version()` uses `winreg`, `get_default_language()` uses `locale.getlocale()`, `get_user_account_type()` uses `winreg`.
- [ ] **Eliminate remaining PowerShell** -- Toast notifications via WinRT COM, app launch via `CreateProcessW`, culture via `GetUserDefaultLocaleName`.

---

## P4.5 -- Enterprise & Integration

- [ ] **Playwright-MCP bridge** -- Detect browser windows and forward to `playwright-mcp` for DOM tasks. Use existing ProxyClient pattern from remote mode.
- [ ] **Power Automate Desktop flow trigger** -- `PADFlow(flow_name, inputs)` tool via Dataverse REST API POST to `RunDesktopFlow`.
- [ ] **Capability/permission manifest** -- Tauri-inspired `WINDOWS_MCP_ALLOW`/`WINDOWS_MCP_DENY` env vars. Evaluated in `lifespan` context manager.
- [ ] **WatchDog-backed WaitForEvent** -- Subscribe to `Window_WindowOpened`/`Window_WindowClosed` events. Expose as `WaitForEvent(event, name, timeout)`.
- [ ] **Win32 message fallback** -- `PostMessage(WM_COMMAND/WM_LBUTTONDOWN)` for legacy apps without UIAutomation. AutoHotkey-level coverage.
- [ ] **Session recording** -- Middleware decorator logging tool calls + before/after screenshots when `RECORD_SESSION=true`.
