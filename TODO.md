# Windows-MCP TODO

**Generated:** 2026-02-18 from REVIEW.md findings
**Reference:** See [REVIEW.md](REVIEW.md) for full context on each item.

---

## P0 -- Must Fix Before Any Shared Deployment

- [x] **[S2] Add authentication for SSE/HTTP transport** -- Bearer token auth via `BearerAuthMiddleware`. DPAPI key storage via `AuthKeyManager`. CLI: `--api-key`, `--generate-key`, `--rotate-key`. Refuses `0.0.0.0` without auth.
- [ ] **[A4] Fix analytics decorator binding bug** -- `@with_analytics(analytics, ...)` captures `None` at decoration time. Pass a callable/lambda that resolves the global at call time, or use DI through FastMCP context.
- [x] **[P4] Fix PIL ImageDraw thread safety** -- Removed `ThreadPoolExecutor` from `get_annotated_screenshot`. Now draws sequentially (<5ms).
- [ ] **[A5] Remove `ipykernel` from dependencies** -- Unused, pulls in jupyter/zmq/tornado. Move to `[dev]` optional deps if needed.
- [x] **[Q6] Remove debug `print()` from analytics.py:97** -- Removed. Also fixed 3 `print()` calls in `watchdog/event_handlers.py` (replaced with `logger.debug()`).

---

## P1 -- High Impact Performance & Safety

- [ ] **[P6] Cache Start Menu app list** -- Add TTL-based cache (30-60s) to `get_apps_from_start_menu`. Currently shells out to PowerShell on every `launch_app` call.
- [ ] **[P5] Eliminate duplicate VDM desktop enumeration** -- `get_state` calls `get_current_desktop()` and `get_all_desktops()` separately; both enumerate all desktops. Combine into a single call.
- [x] **[P2] Bound ThreadPoolExecutor** -- Added `max_workers=min(8, os.cpu_count() or 4)` to tree traversal executor in `tree/service.py`.
- [ ] **[P3] Convert tree_traversal to iterative** -- Replace recursion with an explicit stack to avoid Python's 1000-frame limit on complex UIs.
- [x] **[P9] Make pyautogui.PAUSE configurable** -- Reduced from 1.0s to 0.05s in both `desktop/service.py` and `__main__.py`. Saves 1-6s per input operation.
- [x] **[S7] Block SSRF in Scrape tool** -- Done. `ScraperService.validate_url()` blocks non-HTTP schemes, private/reserved IPs, DNS rebinding, cloud metadata endpoints. Extracted to `scraper/service.py`.
- [ ] **[T1] Fix COM apartment threading violations** -- Make `self.dom` and `self.dom_bounding_box` local variables in traversal context, not shared instance fields.
- [ ] **[T2] Add synchronization to shared mutable state** -- Lock or per-request context for `Tree.tree_state`, `Desktop.desktop_state`.

---

## P2 -- Architectural Improvements

- [ ] **[A1] Decompose Desktop God Object** -- Extract into focused services (3/6 done):
  - [x] `RegistryService` -- registry_get/set/delete/list (`registry/service.py`)
  - [x] `ShellService` -- execute, check_blocklist, ps_quote (`shell/service.py`)
  - [x] `ScraperService` -- validate_url, scrape (`scraper/service.py`)
  - [ ] `InputService` -- click, type, scroll, drag, move, shortcut, multi_select, multi_edit
  - [ ] `WindowService` -- get_windows, get_active_window, switch_app, resize_app
  - [ ] `ScreenService` -- get_screenshot, get_annotated_screenshot
- [ ] **[A3] Replace module globals with dependency injection** -- Use FastMCP's lifespan context to pass `desktop`, `watchdog`, `analytics` via `ctx.request_context.lifespan_context`.
- [ ] **[A2] Decompose __main__.py** -- Extract tool registrations into domain modules (e.g., `tools/input_tools.py`, `tools/window_tools.py`, `tools/file_tools.py`).
- [ ] **[P5] Parallelize get_state** -- Run window enumeration, VDM queries, and tree traversal concurrently with `asyncio.gather`.
- [ ] **[Q1] Deduplicate UIA constants** -- Extract 13 shared constants from `uia/core.py`, `uia/controls.py`, `uia/patterns.py` into `uia/constants.py`.
- [ ] **[Q2] Extract boolean coercion utility** -- Replace 6+ instances of `x is True or (isinstance(x, str) and x.lower() == "true")` with a single `to_bool()` helper.

---

## P3 -- Security Hardening

- [x] **[S1] Implement Shell tool sandboxing** -- Done. `ShellService.check_blocklist()` with 16 regex patterns (format, diskpart, bcdedit, rm -rf, net user/localgroup, IEX+DownloadString, etc). Configurable via `WINDOWS_MCP_SHELL_BLOCKLIST` env var. Extracted to `shell/service.py`.
- [ ] **[S4] Add path-scoping to File tool** -- Define allowed directory scope. Validate resolved paths remain within scope.
- [ ] **[S3] Restrict Registry tool access** -- Deny security-sensitive keys (Run, RunOnce, Services, Policies, SAM). Require confirmation for write/delete.
- [ ] **[S5] Rotate PostHog API key** -- Move to environment variable. Disable GeoIP, disable exception auto-capture. Default telemetry to opt-in.
- [ ] **[S6] Enforce HTTPS for auth client** -- Change default from `http://` to `https://`. Add certificate verification. Reduce `__repr__` key exposure to 4 chars.
- [ ] **[S8] Add protected process list** -- Prevent killing csrss, lsass, services, svchost, winlogon, MsMpEng.
- [ ] **Implement security audit logging** -- Log all tool invocations with timestamps, parameters, results.
- [ ] **Implement rate limiting** -- Per-tool rate limits to prevent abuse.

---

## P4 -- Test Coverage

- [x] **[A7] Add MCP tool handler integration tests** -- Done. 145 headless integration tests in `test_mcp_integration.py` covering all 22 tools via `tool.fn()`. 4 tiers: server structure, tool dispatch, error handling, transport+auth.
- [ ] **Add WatchDog service tests** -- Test threading, COM event handling, callback dispatch.
- [ ] **Add Desktop.get_state orchestration tests** -- Test the composition of windows, tree state, screenshots, VDM.
- [ ] **Add tree_traversal unit tests** -- Test with mock UIA trees. Cover DOM/interactive/scrollable classification.
- [ ] **Achieve 85% overall test coverage** -- Current coverage unknown (only data model tests exist).

---

## P5 -- Code Quality

- [ ] **[Q3] Remove wildcard imports in controls.py** -- Replace `from .enums import *` with explicit imports.
- [ ] **[Q4] Remove dead config** -- Delete `STRUCTURAL_CONTROL_TYPE_NAMES` from `tree/config.py`.
- [ ] **[Q5] Fix broken traceback in analytics.py:107** -- Use `traceback.format_exception()` instead of `str(error)`.
- [ ] **[Q7] Fix resource leaks** -- Close HTTP response in `scrape()`. Guard `handle` in `auto_minimize`. Guard `CoUninitialize` on init failure.
- [ ] **[P10] Cache COM pattern calls in tree traversal** -- Store `GetLegacyIAccessiblePattern()` result in local variable instead of calling 3 times per node.
- [ ] **[P10] Fix BuildUpdatedCache** -- Check for existing cached state before issuing round-trip in `get_cached_children`.
- [ ] **Use set literals in tree/config.py** -- Replace `set([...])` with `{...}`.
- [ ] **[A6] Consolidate input simulation** -- Choose pyautogui OR uia for mouse/keyboard. Remove dual-path confusion.
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
- [ ] **Replace `ImageGrab.grab` with `mss` library** -- DirectX DDA on Windows is 2-5x faster than PIL/GDI for multi-monitor capture.

---

## P3.5 -- Rust Acceleration (PyO3 Extension)

- [ ] **Rust tree traversal module** -- `windows_mcp_core.capture_tree(handles, opts)` via PyO3. Uses `windows-rs` IUIAutomation + `rayon` for true parallel window traversal. Estimated: 200-800ms -> 50-200ms. *(PyO3 scaffold in `native/` ready -- `system_info()` function implemented as Phase 1)*
- [ ] **Rust screenshot module** -- DXGI Output Duplication + `image` crate for capture + annotation. 55-150ms -> 12-35ms.
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
