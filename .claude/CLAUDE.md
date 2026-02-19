# Development Best Practices -- Windows-MCP

Agent-facing reference for code development standards, optimization goals, and architectural decisions.

## Commands (Quick Reference)

```bash
uv sync --extra dev                        # Install all deps
uv run python -m pytest tests/             # Run tests (140 tests, ~6s)
uv run python -m pytest tests/ -x -v       # Stop on first failure, verbose
uv run python -m pytest tests/ --cov       # With coverage
ruff check . --fix && ruff format .        # Lint + format
uv run windows-mcp                         # Launch server (stdio)
uv run windows-mcp --transport sse --port 8000  # Launch server (SSE)
```

## Development Rules

1. **Always `uv run python`** -- never bare `python` or `pip`
2. **Read before modifying** -- understand existing patterns before changing code
3. **Run tests after changes** -- `uv run python -m pytest tests/` must pass
4. **COM threading** -- each thread needs its own `CoInitialize` and UIA instance. Never share COM objects across threads
5. **No wildcard imports** except in `uia/*.py` (ruff config exempts these)
6. **Line length 100** -- ruff enforces this
7. **Type hints on all function signatures** -- existing convention

## Current Optimization Goals (Priority Order)

### Phase 1: Python Quick Wins -- DONE
1. ~~`pg.PAUSE = 0.05`~~ -- Done (both `desktop/service.py:45` and `__main__.py:33`)
2. ~~Fix ImageDraw thread safety~~ -- Done (sequential loop replaces ThreadPoolExecutor)
3. ~~Replace PowerShell subprocess with stdlib~~ -- Done (winreg, locale, platform)
4. ~~Fix analytics print() corrupting stdout~~ -- Done
5. ~~Fix watchdog event_handlers.py print() calls~~ -- Done (replaced with logger.debug)
6. ~~Bound ThreadPoolExecutor in tree/service.py~~ -- Done (max_workers=min(8, cpu_count))
7. Single `TreeScope_Subtree` CacheRequest on window root instead of per-node `BuildUpdatedCache` -- DEFERRED (needs live UIA testing)
8. Deduplicate `LegacyIAccessiblePattern` calls (called up to 3x per element) -- DEFERRED

### Phase 2: Capability Gaps to Fill
1. **WaitFor tool** -- event-driven waiting (like Playwright's `waitForSelector`)
2. **Find tool** -- semantic element search by role, name, or property
3. **Invoke tool** -- UIA pattern invocation (InvokePattern, ValuePattern, TogglePattern)
4. **Win32 message fallback** -- `SendMessage`/`PostMessage` when UIA patterns unavailable

### Phase 3: Rust Acceleration (PyO3 Extension)
- Target: `windows_mcp_core.pyd` via Maturin
- Hot path: tree traversal (500-5000ms -> 50-200ms with `windows-rs`)
- Secondary: screenshot (DXGI Output Duplication), Win32 ops
- Key: `py.allow_threads()` to release GIL during COM traversal

## Architecture Decisions (Settled)

| Decision | Choice | Rationale |
|----------|--------|-----------|
| MCP framework | FastMCP (Python) | Already working, protocol layer stays Python |
| COM interop | comtypes (current), windows-rs (future Rust) | comtypes works, Rust for perf-critical paths |
| Input simulation | pyautogui (current), SendInput (future) | SendInput is modern Win32 API |
| Build system | Hatchling (Python), Maturin (future Rust ext) | Separate concerns |
| Playwright | Complementary bridge, not replacement | Playwright = browser-only, Windows-MCP = all apps |
| FlaUI patterns | Adopt InvokePattern/ValuePattern approach | More reliable than coordinate clicking |

## Key File Paths for Common Tasks

| Task | Files |
|------|-------|
| Add/modify MCP tools | `src/windows_mcp/__main__.py` |
| Desktop automation logic | `src/windows_mcp/desktop/service.py` |
| Tree traversal / a11y | `src/windows_mcp/tree/service.py`, `tree/config.py`, `tree/cache_utils.py` |
| COM / UIAutomation | `src/windows_mcp/uia/core.py`, `uia/patterns.py`, `uia/controls.py` |
| Input simulation | `src/windows_mcp/uia/core.py` (mouse_event, keybd_event calls) |
| Screenshot | `src/windows_mcp/desktop/service.py` (~line 875) |
| Shell execution | `src/windows_mcp/desktop/service.py` (~line 209) |
| File operations | `src/windows_mcp/filesystem/service.py` |
| Virtual desktops | `src/windows_mcp/vdm/core.py` |
| Focus monitoring | `src/windows_mcp/watchdog/service.py` |
| Telemetry | `src/windows_mcp/analytics.py` |
| Auth (remote mode) | `src/windows_mcp/auth/service.py` |
| Tests | `tests/` |

## Agent Context Slices

Pre-built context for specialist agents in `.claude/context/slices/`:
- `python-slice.md` -- Python code quality, bugs, performance issues
- `security-slice.md` -- Security findings (2C/4H/3M/3L)
- `rust-slice.md` -- Rust migration strategy, crates, COM threading
- `performance-slice.md` -- Bottleneck ranking, latency budget, caching opportunities

## Known Patterns to Follow

- **Service pattern**: Each module has a `service.py` with a class that owns the logic, `views.py` for data models
- **Tool registration**: `@mcp.tool()` decorator in `__main__.py`, delegates to `Desktop` methods
- **Analytics wrapping**: `@with_analytics(analytics, "tool_name")` on each tool (NOTE: captures None at decoration time)
- **Error handling**: FastMCP handles exceptions and returns them as tool errors
- **Browser detection**: Check window title for Chrome/Edge/Firefox to enable DOM mode
- **Element classification**: `tree/config.py` classifies UIA control types as interactive, scrollable, or informative
- **Registry ops**: Use `winreg` stdlib (not PowerShell) -- supports PS-style paths (HKCU:\, HKLM:\)
- **No print() to stdout**: Use `logging` module only -- print corrupts MCP protocol stream
- **Auth (local HTTP)**: API key auth via DPAPI storage, Starlette middleware for SSE/HTTP transports

## Recommended Agents & Tools

| Task Domain | Agent/Tool | When to Use |
|-------------|-----------|-------------|
| Python code changes | `python-pro` | Refactoring service.py, fixing bugs |
| Rust extension dev | `rust-pro` | PyO3 crate work in `native/` |
| Security review | `security-auditor` | Auth system, registry ops, shell execution |
| Performance tuning | `performance-engineer` | Profiling tree traversal, COM overhead |
| Test generation | `test-automator` | Adding coverage for new auth/native modules |
| Code review | `superpowers:code-reviewer` | After completing major steps |
| Context7 MCP | `context7:resolve-library-id` + `query-docs` | Latest FastMCP, PyO3, maturin docs |
| Architecture | `architect-reviewer` | Reviewing auth middleware, Rust FFI design |

### MCP Servers Available
- **context7** -- Up-to-date library docs (FastMCP, PyO3, maturin, windows-rs)
- **Cloudflare** -- Remote deployment (D1, KV, Workers) if cloud hosting needed
- **Hugging Face** -- ML model integration if AI features added
- **Notion** -- Documentation/project management integration
