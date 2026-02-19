# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Windows-MCP is a Python MCP (Model Context Protocol) server that bridges AI LLM agents with the Windows OS, enabling direct desktop automation. It exposes 19 tools (App, Shell, Snapshot, Click, Type, Scroll, Move, Shortcut, Wait, Scrape, MultiSelect, MultiEdit, Clipboard, Process, SystemInfo, Notification, LockScreen, Registry) via FastMCP. Uses Windows Accessibility Tree (UIAutomation COM) for element discovery -- works on ALL Windows apps, not just browsers.

## Build & Development Commands

```bash
uv sync                              # Install dependencies
uv sync --extra dev                  # Install with dev deps (ruff, pytest)
uv run windows-mcp                   # Run the MCP server (stdio transport)
uv run python -m pytest tests/       # Run all tests (140 tests, ~6s)
uv run python -m pytest tests/test_foo.py  # Run a single test file
ruff format .                        # Format code
ruff check .                         # Lint code
ruff check --fix .                   # Lint and auto-fix
```

**Package manager**: UV (not pip/bare python). **Python**: 3.13+. **Build backend**: Hatchling. **Test framework**: pytest + pytest-asyncio (async mode: auto).

## Architecture

The codebase follows a layered service architecture under `src/windows_mcp/`:

**Entry point** -- `__main__.py` (~700 lines): Registers all 19 MCP tools on a FastMCP server instance. Uses an async lifespan to initialize Desktop, WatchDog, and Analytics services. Each tool function delegates to `Desktop` methods. The `@with_analytics` decorator wraps tools for telemetry.

**Desktop service** -- `desktop/service.py` (~1087 lines): High-level orchestrator (God Object -- refactor target). Manages window operations (launch, resize, switch), screenshots, mouse/keyboard actions, clipboard, registry, and shell execution. `desktop/views.py` defines data models: `DesktopState`, `Window`, `Size`, `BoundingBox`, `Status`.

**Tree service** -- `tree/service.py`: Captures the Windows accessibility tree from active and background windows. Identifies interactive, informative, and scrollable elements. Uses `ThreadPoolExecutor` for multi-threaded UI traversal. `tree/views.py` defines `TreeElementNode`, `ScrollElementNode`, `TreeState`. `tree/config.py` has control type classifications. `tree/cache_utils.py` has CacheRequest factory.

**UIAutomation wrapper** -- `uia/`: Low-level abstraction over the Windows UIAutomation COM API via `comtypes`. `core.py` wraps the main automation object, `controls.py` has control-specific logic, `patterns.py` wraps UIAutomation patterns, `enums.py` has COM enumerations, `events.py` handles event subscriptions.

**Auth** -- `auth/service.py`: Authentication client for remote mode (windowsmcp.io). HTTP-based API key validation.

**Filesystem** -- `filesystem/service.py`: Stateless file operations. Well-designed module with pure functions and clean error handling. `filesystem/views.py` has data models.

**WatchDog** -- `watchdog/service.py`: Runs in a separate thread monitoring UI focus changes via UIAutomation events. Notifies the Tree service of focus changes so the accessibility tree stays current.

**Virtual Desktop Manager** -- `vdm/core.py` (~714 lines): Tracks which windows belong to which Windows virtual desktop (Win10/11). Has redundant desktop enumeration (optimization target).

**Analytics** -- `analytics.py`: Optional PostHog telemetry (disabled with `ANONYMIZED_TELEMETRY=false` env var). Tracks tool names and errors only, not arguments or outputs.

## Code Style

- Formatter/linter: **Ruff** (line length 100, double quotes, target py313)
- Naming: PEP 8 -- `snake_case` functions/variables, `PascalCase` classes, `UPPER_CASE` constants
- Type hints required on function signatures
- Google-style docstrings for public functions/classes
- Wildcard imports (`F403`, `F405`) suppressed in `uia/*.py` only

## Testing

- 140 tests, all passing (~6s runtime)
- Framework: pytest + pytest-asyncio (asyncio_mode = "auto")
- Test directory: `tests/`
- Must run via `uv run python -m pytest` (not bare `pytest`)

## Environment Variables

| Variable | Purpose | Default |
|----------|---------|---------|
| `ANONYMIZED_TELEMETRY` | Disable PostHog telemetry | `true` |
| `MODE` | `remote` for cloud VM mode | local |
| `SANDBOX_ID` | VM identifier for remote mode | - |
| `API_KEY` | API key for remote mode | - |

## Key Design Details

- Screenshots are capped to 1920x1080 for token efficiency
- `pyautogui.FAILSAFE` is disabled; `PAUSE` is set to 1.0s between actions (major perf bottleneck)
- Browser detection (Chrome, Edge, Firefox) triggers special DOM extraction mode in Snapshot
- Fuzzy string matching (`thefuzz`) is used for element name matching
- UI element fetching has retry logic (`THREAD_MAX_RETRIES=3` in tree service)
- The server supports stdio, SSE, and streamable HTTP transports

## Known Issues & Gotchas

**Critical bugs (unfixed as of v0.6.2):**
- `analytics.py` decorator captures `None` at decoration time -- telemetry silently does nothing
- COM objects (`self.dom`, `self.dom_bounding_box`) shared across thread apartment boundaries in `tree/service.py:583-584` -- undefined behavior under concurrent access
- PIL `ImageDraw` used from `ThreadPoolExecutor` in `desktop/service.py:875` -- not thread-safe
- `analytics.py:97` has a `print()` that interleaves with MCP protocol stdout
- No authentication on SSE/HTTP transport -- any network client can invoke all 19 tools

**Performance gotchas:**
- `pg.PAUSE = 1.0` adds 1 second of sleep after EVERY pyautogui call. A `type(clear=True, press_enter=True)` is 6 seconds of pure sleep
- UIA `BuildUpdatedCache` called per-node with `TreeScope_Element` + `TreeScope_Children` -- two cross-process COM round-trips per tree node instead of one `TreeScope_Subtree` call
- PowerShell subprocess (~200-500ms per call) used for 7+ operations that have Python stdlib alternatives (`winreg`, `platform`, `locale`)
- `comtypes` adds ~50-200us per COM call overhead -- 10,000 calls per Snapshot = ~1000ms waste

**Threading model:**
- Main thread: STA COM apartment (via comtypes `CoInitialize`)
- Tree traversal: `ThreadPoolExecutor` with no `max_workers` bound
- WatchDog: Separate dedicated thread with STA COM apartment
- COM objects are NOT safe to share across threads -- each thread needs its own UIA instance

## Security Context

This server has **full system access** with no sandboxing. Tools like Shell and App can perform irreversible operations. The recommended deployment target is a VM or Windows Sandbox. See `REVIEW.md` Section 3 and `SECURITY.md` for details.

Key risks: unrestricted shell execution, no path scoping on filesystem, no URL validation on Scrape (SSRF), hardcoded PostHog API key, auth client uses plain HTTP.

## Git & Fork Info

- **Upstream**: CursorTouch/Windows-MCP (original repo)
- **Origin**: David-Martel/Windows-MCP (fork)
- **Branch**: main @ b6c2a04

## Project Documentation

- `REVIEW.md` -- Comprehensive code review (7 sections: security, architecture, performance, thread safety, code quality, testing, framework comparison)
- `TODO.md` -- Prioritized action items (P0-P4.5)
- `USER_MANUAL.md` -- Full user manual with tool reference, framework comparison, performance data
- `.claude/context/` -- Multi-agent context with slices for rust, performance, python, security
