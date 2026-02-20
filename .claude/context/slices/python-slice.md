# Python Context Slice -- Windows-MCP (2026-02-20)

## Architecture (Post-Decomposition)

Desktop service fully decomposed into 8 extracted services:
- InputService, WindowService, ScreenService, VisionService
- ProcessService, ShellService, ScraperService, RegistryService
- Desktop remains as thin facade/orchestrator (683 lines)

Tool registration decomposed into 3 modules:
- tools/input_tools.py (9 tools), tools/state_tools.py (5 tools), tools/system_tools.py (10 tools)

## Performance Status (All Phase 1 DONE)
- pg.PAUSE = 0.05 (not 1.0)
- ImageDraw thread safety fixed (sequential)
- PowerShell replaced with stdlib (winreg, locale, platform)
- analytics print() removed
- watchdog print() -> logger.debug
- ThreadPoolExecutor bounded (max_workers=min(8, cpu_count))

## Remaining Perf Targets
- Single TreeScope_Subtree CacheRequest (deferred)
- Deduplicate LegacyIAccessiblePattern calls (deferred)
- Parallelize get_state orchestration (asyncio.gather)

## Security Features
- Shell blocklist (16 patterns, configurable)
- BearerAuth + DPAPI key storage
- Registry sensitive key blocking
- File path scoping (WINDOWS_MCP_ALLOWED_PATHS)
- SSRF protection (private IP, DNS rebinding)
- Protected process list
- Rate limiting (sliding window, per-tool locks, configurable)
- Permission manifest (WINDOWS_MCP_ALLOW/DENY)
- Audit logging (tab-separated file)

## Key Patterns
- `@with_analytics(lambda: _state.analytics, "Tool-Name")` for telemetry
- `_state.desktop` patched in tests (not `main_module.desktop`)
- Rust fast-paths: native functions return None on failure, Python UIA fallback
- Test isolation: `@patch("windows_mcp.native.native_*", return_value=None)` or autouse fixture with `patch.multiple`
- No print() to stdout (corrupts MCP protocol)

## Test Coverage
- 2115 Python tests + 25 Rust + 13 live-desktop
- Testable modules: 82-100% coverage
- native.py: 90%
- COM/UIA modules: 14-42% (inherently untestable without live desktop)
- desktop/service.py: ~73%

## Well-Designed Modules
- filesystem/service.py -- stateless pure functions
- tree/ -- proper config/cache/views/service separation
- analytics.py -- protocol pattern, rate limiting, audit logging, permissions
- native.py -- centralized adapter with HAS_NATIVE guard
