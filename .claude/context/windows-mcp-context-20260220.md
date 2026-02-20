# Windows-MCP Context -- 2026-02-20

## Project State

**Version:** 0.6.2 | **Branch:** main @ b6ee6e2 | **Type:** Python + Rust mixed
**Tests:** 2115 Python + 25 Rust + 13 live-desktop = 2153 total, all passing (~23s Python, ~1s Rust)
**Coverage:** ~64% overall; testable modules 82-100%, native.py 90%, COM/UIA 14-42%
**Tools:** 24 MCP tools registered via FastMCP

## Summary

Completed the Rust-Native UIA Replacement Plan (5 phases). Added `query.rs` (ElementFromPoint, FindElements, ScreenMetrics) and `pattern.rs` (Invoke, Toggle, SetValue, Expand, Collapse, Select) to wmcp-core. Wired 9 new PyO3 functions through native.py into Find and Invoke tools with automatic Python UIA fallback. Added 25 Rust unit tests and 13 live-desktop integration tests. Fixed test isolation for Rust fast-paths in 3 test classes. Multiple background agents completed: ProcessService tests (51), VDM tests (14), audit logging tests (31), rate limiting implementation (50 tests), permission manifest tests (39), ValuePattern tests (50), tree traversal tests (58), Desktop.get_state tests (60), return type annotations (25 methods).

## Recent Changes (Last 10 Commits)

| Commit | Description |
|--------|-------------|
| b6ee6e2 | test: add live-COM integration tests and fix native fast-path isolation |
| 3faf2b4 | feat(native): wire Rust UIA query/patterns into Python tools with fallback |
| b8ead70 | feat(native): add Rust UIA query and pattern modules with unit tests |
| b54b34c | refactor(desktop): remove redundant bool coercion from get_state |
| 4443ef4 | refactor(quality): remove dead code, hoist structs, fix typos |
| e542f66 | docs: update TODO.md header with accurate test count |
| e84ab61 | feat(input): wire Rust SendInput for keyboard shortcuts |
| f880ea3 | docs: mark WaitForEvent complete |
| c4fdf2a | feat: add WaitForEvent tool for UIA automation event subscriptions |
| 5af0860 | fix(vision): close HTTP response to prevent connection pool leak |

## Architecture

### Cargo Workspace (native/, 4 crates)
```
wmcp-core/    Pure Rust lib -- system_info, input, tree, screenshot, query, pattern
wmcp-pyo3/    PyO3 wrappers (24 functions) -- windows_mcp_core.pyd
wmcp-ffi/     C ABI DLL (12 exports) -- windows_mcp_ffi.dll
wmcp-cli/     CLI binaries -- wmcp-worker (JSON-RPC)
```

### Python Services (src/windows_mcp/)
```
desktop/service.py    Facade orchestrator (683 lines)
input/service.py      Click, type, scroll, drag, move, shortcut
window/service.py     Rust fast-path + UIA fallback, focus, overlay
screen/service.py     Screenshot, annotated screenshot, DPI
vision/service.py     LLM-powered screenshot analysis
process/service.py    List/kill processes, protected list
shell/service.py      PowerShell execution, blocklist
scraper/service.py    Web fetching, SSRF protection
registry/service.py   CRUD, sensitive key blocking
tree/service.py       UIA accessibility tree traversal
native.py             Centralized Rust adapter (24 functions)
analytics.py          PostHog telemetry + audit logging + rate limiting + permissions
tools/                MCP tool registration (3 modules)
```

### Tool Module Decomposition
- `tools/__init__.py`: register_all_tools(mcp)
- `tools/input_tools.py`: 9 tools (Click, Type, Scroll, Move, Shortcut, Wait, MultiSelect, MultiEdit, Invoke)
- `tools/state_tools.py`: 5 tools (Snapshot, WaitFor, WaitForEvent, Find, VisionAnalyze)
- `tools/system_tools.py`: 10 tools (App, Shell, File, Scrape, Process, SystemInfo, Notification, LockScreen, Clipboard, Registry)

## Decisions Made This Session

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Rust UIA query approach | ElementFromPoint + FindAll with PropertyConditions | Direct COM, no Python overhead |
| Pattern invocation | Point-based (x,y) rather than element handle | Matches existing tool API (loc=[x,y]) |
| Test isolation for native | @patch + autouse fixtures | Prevents live desktop queries in unit tests |
| Commit strategy | 3 logical clusters (Rust core, Python integration, tests) | Clean git history |
| Live test marker | `live_desktop` with `-m "not live_desktop"` default | CI-safe, manual live testing |

## Security Features Active
- S1: Shell blocklist (16 patterns)
- S2: BearerAuth + DPAPI key storage
- S3: Registry sensitive key blocking
- S4: File path scoping (WINDOWS_MCP_ALLOWED_PATHS)
- S7: SSRF protection (private IP, DNS rebinding)
- S8: Protected process list
- Rate limiting: sliding window, per-tool locks
- Permission manifest: WINDOWS_MCP_ALLOW/DENY env vars
- Audit logging: WINDOWS_MCP_AUDIT_LOG env var

## Background Agent Work (Uncommitted)

Several background agents produced test files that may not yet be committed:

| Agent Task | Tests Added | File |
|------------|-------------|------|
| ProcessService + VDM tests | 51 + 14 | test_process_service.py, test_vdm_service.py |
| Audit logging tests | 31 | test_analytics.py |
| Rate limiting impl + tests | 50 | analytics.py, test_analytics.py |
| Permission manifest tests | 39 | test_analytics.py |
| ValuePattern + input tests | 50 | test_input_service.py |
| Tree traversal tests | 58 | test_tree_service.py |
| Desktop.get_state tests | 60 | test_get_state.py |
| Return type annotations | 25 methods | input/service.py, window/service.py, desktop/service.py |

## Roadmap

### Immediate
- Commit/push background agent work (tests, rate limiting, permissions, annotations)
- Verify all background agent changes don't conflict

### This Week
- Convert tree_traversal to iterative (eliminate recursion limit)
- Parallelize get_state orchestration (asyncio.gather)

### Tech Debt
- VDM duplicate enumeration elimination
- Module globals -> dependency injection (FastMCP lifespan context)
- UIA constant deduplication (partially done)

### Performance
- Single TreeScope_Subtree CacheRequest (deferred, needs live testing)
- Deduplicate LegacyIAccessiblePattern calls
