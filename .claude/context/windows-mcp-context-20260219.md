# Windows-MCP Project Context

**Date:** 2026-02-19
**Branch:** main @ 0fb7e1e
**Version:** 0.6.2

## Project Summary

Windows-MCP is a Python MCP server (FastMCP) that bridges AI agents with Windows OS for desktop automation. 19 tools exposed via MCP protocol covering UI automation, file ops, shell execution, registry access, process management, and more. Uses Windows Accessibility Tree (UIAutomation COM) for element discovery -- works on ALL Windows apps, not just browsers.

## Current State

- **Build:** Clean, 184 tests passing (5-9s)
- **Rust Extension:** `windows_mcp_core` PyO3 crate built and working (`system_info()` function)
- **Auth:** Bearer token middleware + DPAPI key storage implemented
- **Build System:** Unified `Build.ps1` with CargoTools integration
- **Git:** main @ 0fb7e1e, uncommitted changes across 18+ files (ready to commit)

## Session 3: Implementation Work (2026-02-19)

### Completed
1. **MCP Config Files** -- Created 4 client configs: `mcp-configs/claude-desktop.json`, `claude-code.json`, `codex-cli.toml`, `gemini-cli.json`
2. **Python Performance Quick Wins:**
   - `pg.PAUSE` 1.0 -> 0.05 (saves 1-6s per input)
   - ImageDraw: ThreadPoolExecutor -> sequential loop (fixes race condition)
   - analytics.py: removed print() corrupting MCP stdout
   - watchdog/event_handlers.py: 3 print() -> logger.debug()
   - tree/service.py: bounded ThreadPoolExecutor max_workers=min(8, cpu_count)
3. **PowerShell -> stdlib:** All 4 registry methods, get_windows_version, get_default_language, get_user_account_type rewritten to use winreg/locale/platform
4. **Auth System:**
   - `auth/key_manager.py`: AuthKeyManager (DPAPI encrypt/decrypt, key gen/rotation/validation)
   - `auth/middleware.py`: BearerAuthMiddleware (FastMCP Middleware pattern)
   - CLI: `--api-key`, `--generate-key`, `--rotate-key`
   - Safety: refuses non-localhost without auth
5. **PyO3 Rust Extension:**
   - `native/` scaffold: Cargo.toml, lib.rs, errors.rs, system_info.rs
   - system_info() working: OS, CPU count, memory, disks via sysinfo crate
   - OnceLock<Mutex<System>> singleton, py.allow_threads() for GIL release
6. **Test Suite:** 140 -> 184 tests (44 new), all passing
7. **Build.ps1:** Unified build script with CargoTools, Find-NativeDll, version tagging, GitHub release

### Agent Work This Session

| Agent | Task | Status | Key Output |
|-------|------|--------|-----------|
| security-auditor | MCP server security audit | Complete | Auth system requirements |
| architect-reviewer | Architecture review | Complete | Service decomposition plan |
| rust-pro | PyO3 scaffold | Complete | native/ crate with system_info |
| performance-engineer | Performance analysis | Complete | Bottleneck ranking, latency budget |
| python-pro | Code quality | Complete | Print removal, registry rewrite |
| Explore (background) | PC-AI Rust patterns | In progress | Reusable patterns for acceleration |

## Build Patterns Discovered

- **sccache port 4226 blocked** on this machine by Windows port exclusion
- **maturin + UV venv:** uv run maturin develop doesn't install into UV venv. Workaround: cargo build then copy .dll as .pyd to .venv/Lib/site-packages/
- **CargoTools shared target:** DLL at T:\RustCache\cargo-target\release\windows_mcp_core.dll
- **Build.ps1 actions:** Build, Test, Lint, Native, Check, Clean, Release, All

## Decisions Made

| Topic | Decision | Rationale |
|-------|----------|-----------|
| Auth mechanism | Bearer token + DPAPI storage | Cross-client compatible, leverages existing pywin32 |
| Registry ops | winreg stdlib | Eliminates 200-500ms PowerShell overhead per call |
| Rust Phase 1 | system_info() first | Lowest risk (no COM), validates PyO3 build pipeline |
| Build system | Build.ps1 + CargoTools | Unified orchestration of Python + Rust + Release |
| DLL install | Copy .dll as .pyd | UV venv strictness prevents maturin develop |

## Outstanding Work (from TODO.md)

### High Priority Remaining
- TreeScope_Subtree optimization (needs live UIA testing)
- LegacyIAccessiblePattern dedup (3x per element -> 1x)
- Analytics decorator binding bug (captures None)
- COM apartment threading violations (shared DOM objects)
- Shell tool sandboxing
- SSRF blocking in Scrape tool

### Rust Acceleration Next Steps
- Tree traversal module (biggest perf win: 500-5000ms -> 50-200ms)
- Screenshot via DXGI Output Duplication
- SendInput replacement for pyautogui

### New Capability Gaps
- WaitFor tool (event-driven waiting)
- Find tool (semantic element lookup)
- Invoke tool (UIA pattern actions)
- Win32 message fallback

## Tech Stack

- Python 3.13+ / UV / Hatchling / FastMCP
- Rust (PyO3 0.23 / Maturin / sysinfo 0.33 / parking_lot 0.12)
- comtypes (COM/UIAutomation) / pyautogui / pywin32
- ruff / pytest / pytest-asyncio
- CargoTools (PowerShell build framework)
