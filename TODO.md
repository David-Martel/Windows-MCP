# Windows-MCP TODO

**Generated:** 2026-02-18 from REVIEW.md findings
**Last updated:** 2026-02-19 -- 2099 tests, 64% coverage, Rust workspace (4 crates, 15 PyO3 + 12 FFI exports), tools/ decomposition complete
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
- [x] **[Q1] Deduplicate UIA constants** -- Extracted 14 shared constants from `uia/core.py`, `uia/controls.py`, `uia/patterns.py`, `uia/enums.py` into `uia/constants.py`. Removed unused imports.
- [x] **[Q2] Extract boolean coercion utility** -- `_coerce_bool()` in `tools/_helpers.py` used by all 23 tool handlers. Service layer callers (input, desktop) use simplified checks since tool layer already coerces.

---

## P3 -- Security Hardening

- [x] **[S1] Implement Shell tool sandboxing** -- Done. `ShellService.check_blocklist()` with 16 regex patterns (format, diskpart, bcdedit, rm -rf, net user/localgroup, IEX+DownloadString, etc). Configurable via `WINDOWS_MCP_SHELL_BLOCKLIST` env var. Extracted to `shell/service.py`.
- [x] **[S4] Add path-scoping to File tool** -- `WINDOWS_MCP_ALLOWED_PATHS` env var (semicolon-separated). All 8 filesystem functions validate resolved paths against allowed scope. 28 new tests.
- [x] **[S3] Restrict Registry tool access** -- 11 regex patterns block writes to Run/RunOnce/Services/Policies/SAM/SECURITY. `WINDOWS_MCP_REGISTRY_UNRESTRICTED=true` bypasses. 51 new tests.
- [x] **[S5] Rotate PostHog API key** -- `POSTHOG_API_KEY` env var overrides hardcoded default. Disabled GeoIP, disabled exception auto-capture.
- [x] **[S6] Enforce HTTPS for auth client** -- Default dashboard URL changed to `https://windowsmcp.io` (configurable via `DASHBOARD_URL`). `__repr__` exposes only last 4 chars of API key.
- [x] **[S8] Add protected process list** -- `ProcessService.is_protected()` blocks csrss, lsass, services, svchost, winlogon, MsMpEng, smss, wininit, system, registry. Extracted to `process/service.py`.
- [x] **Implement security audit logging** -- `WINDOWS_MCP_AUDIT_LOG` env var enables file-based audit log. Tab-separated: timestamp, OK/ERR, tool_name, duration_ms, error_type. Integrated into `with_analytics` decorator.
- [x] **Implement rate limiting** -- `RateLimiter` class with sliding window per-tool. Defaults: 60/min general, 10/min Shell, 5/min Registry-Set/Delete. Configurable via `WINDOWS_MCP_RATE_LIMITS` env var. Integrated into `with_analytics` decorator.

---

## P4 -- Test Coverage

- [x] **[A7] Add MCP tool handler integration tests** -- Done. 145 headless integration tests in `test_mcp_integration.py` covering all 23 tools via `tool.fn()`. 4 tiers: server structure, tool dispatch, error handling, transport+auth.
- [ ] **Add WatchDog service tests** -- Test threading, COM event handling, callback dispatch.
- [x] **Add Desktop.get_state orchestration tests** -- 60 tests across 9 classes covering structure, no-windows, screenshot/tree failure, VDM info, thread safety, multi-window, vision paths, bool coercion.
- [x] **Add tree_traversal unit tests** -- 54 tests covering tree_traversal (28), focus debounce (6), DOM correction (14), plus mock infrastructure. `test_tree_service.py` up from 4 to 58 tests.
- [ ] **Achieve 85% overall test coverage** -- Currently at 64% overall (1671 tests). Testable modules at 82-100%. COM/UIA modules at 14-42% (require live desktop).

---

## P5 -- Code Quality

- [x] **[Q3] Remove wildcard imports in controls.py** -- Replaced 4 wildcard imports with ~90 explicit named imports from constants, core, enums, patterns.
- [x] **[Q4] Remove dead config** -- `STRUCTURAL_CONTROL_TYPE_NAMES` is actually used in tree/service.py -- not dead. Verified, no action needed.
- [x] **[Q5] Fix broken traceback in analytics.py:107** -- Now uses `traceback.format_exception()` for full stack trace string.
- [x] **[Q7] Fix resource leaks** -- Fixed HTTP response leak in `scrape()` (close intermediate redirect responses and final response in `finally` block). `auto_minimize` handle guard already present. COM init guard deferred (comtypes handles internally).
- [x] **[P10] Cache COM pattern calls in tree traversal** -- Already done in P3 iterative rewrite. `legacy_pattern = None` set once per node, guarded `if legacy_pattern is None:` before each call.
- [x] **[P10] Fix BuildUpdatedCache** -- Subtree caching attempted first (line 916), falls back to per-node only on failure. `get_cached_children` already checks scope flags.
- [x] **Use set literals in tree/config.py** -- Already uses `{...}` set literals. Verified no `set([...])` remains.
- [x] **[A6] Consolidate input simulation** -- Rust Win32 `SendInput` is primary path in `InputService` (click, type, scroll, drag, move, shortcut). Falls back to pyautogui when native unavailable.
- [x] **Add type annotations to all public APIs** -- 25 methods annotated across InputService (7), WindowService (3), Desktop (15). Includes Generator import for context managers.
- [x] **[T3] Fix singleton TOCTOU in _AutomationClient** -- Added double-checked locking with `threading.Lock()` to `instance()`.

---

## P1.5 -- Capability Gaps (from Framework Comparison)

*Source: Phase 2 comparative analysis (rust-pro, backend-architect, performance-engineer agents)*

- [x] **Replace `Wait(duration)` with event-driven `WaitFor`** -- Done. `WaitFor(mode="window"|"element", name, timeout)` in `tools/state_tools.py`. Polls `get_state()` with 0.5s interval, timeout capped at 300s.
- [x] **Add `ValuePattern.SetValue()` as primary text input** -- `_try_value_pattern()` in `InputService.type()` attempts `ControlFromPoint` + `ValuePattern.SetValue()` before falling back to SendInput/pyautogui. Handles clear/append semantics.
- [x] **Reduce `pg.PAUSE` from 1.0 to 0.05** -- Done. Changed in both `desktop/service.py:45` and `__main__.py:33`.
- [x] **Add `Find` tool for semantic element lookup** -- Done. `Find(name, control_type, window, limit)` in `tools/state_tools.py`. Searches interactive nodes by name, type, and window.
- [x] **Add `Invoke` tool for UIA pattern actions** -- Done. `Invoke(loc, action, value)` in `tools/input_tools.py`. Supports invoke, toggle, set_value, expand, collapse, select patterns.
- [x] **Use `TreeScope_Subtree` for single-shot cached traversal** -- Implemented in `tree/service.py:916`. Attempts `TreeScope_Subtree` first, falls back to per-node caching on failure. `CacheRequestFactory.create_subtree_cache()` in `cache_utils.py`.

---

## P2.5 -- Performance Caching Layer

- [ ] **Per-window tree cache with WatchDog invalidation** -- Cache TreeState per window handle. On focus change, only re-traverse changed windows. 80-95% reduction on repeated `get_state()` calls.
- [ ] **Cache Start Menu app list with filesystem watcher** -- Build app map once at init, monitor Start Menu folders via `ReadDirectoryChangesW`. Invalidate on change.
- [ ] **Element coordinate cache** -- Reuse `TreeElementNode.bounding_box` for click/type instead of re-querying XPath. Invalidate on window move/resize.
- [x] **Cache VDM desktop list with TTL** -- `_enumerate_desktops()` results cached for 5s via `_CACHE_TTL`. Cache invalidated on create/remove/rename/switch. Avoids redundant COM round-trips within rapid `get_state()` calls.
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
- [x] **Eliminate remaining PowerShell** -- `launch_app` uses `ShellExecuteExW`/`os.startfile`, `send_notification` uses `Shell_NotifyIconW`, `get_apps_from_start_menu` uses shortcut scan as primary path. Only `Shell` tool and `Get-StartApps` fallback remain (intentional).

---

## P4.5 -- Enterprise & Integration

- [ ] **Playwright-MCP bridge** -- Detect browser windows and forward to `playwright-mcp` for DOM tasks. Use existing ProxyClient pattern from remote mode.
- [ ] **Power Automate Desktop flow trigger** -- `PADFlow(flow_name, inputs)` tool via Dataverse REST API POST to `RunDesktopFlow`.
- [x] **Capability/permission manifest** -- Tauri-inspired `WINDOWS_MCP_ALLOW`/`WINDOWS_MCP_DENY` env vars. Evaluated in `with_analytics` decorator. `ToolNotAllowedError` raised for blocked tools. 39 tests.
- [ ] **WatchDog-backed WaitForEvent** -- Subscribe to `Window_WindowOpened`/`Window_WindowClosed` events. Expose as `WaitForEvent(event, name, timeout)`.
- [ ] **Win32 message fallback** -- `PostMessage(WM_COMMAND/WM_LBUTTONDOWN)` for legacy apps without UIAutomation. AutoHotkey-level coverage.
- [ ] **Session recording** -- Middleware decorator logging tool calls + before/after screenshots when `RECORD_SESSION=true`.
