# Windows-MCP Project Context

**Date:** 2026-02-18 (Updated: Session 2)
**Branch:** main @ b6c2a04
**Version:** 0.6.2

## Project Summary

Windows-MCP is a Python MCP server (FastMCP) that bridges AI agents with Windows OS for desktop automation. 19 tools exposed via MCP protocol covering UI automation, file ops, shell execution, registry access, process management, and more. Uses Windows Accessibility Tree (UIAutomation COM) for element discovery -- works on ALL Windows apps, not just browsers.

## Current State

- **Build:** Clean, all 130 deps installed via `uv sync`
- **Tests:** 140/140 passing (5.96s)
- **Git:** Clean working tree (untracked docs only), forked to David-Martel/Windows-MCP
- **Remotes:** origin=David-Martel/Windows-MCP, upstream=CursorTouch/Windows-MCP

## Work Completed

### Session 1: Clone, Build, Initial Review
1. Cloned repo, forked to David-Martel
2. Built and verified server launches on stdio transport
3. Ran full test suite (140/140 pass)
4. Dispatched 3 specialist agents for deep analysis:
   - python-pro: Code quality, performance, thread safety
   - security-auditor: OWASP-based security audit
   - architect-reviewer: Architecture, SOLID compliance, dependency analysis
5. Created REVIEW.md, TODO.md, USER_MANUAL.md

### Session 2: Comparative Research & Optimization Roadmap
1. Conducted 15+ web searches across Windows Rust APIs, Playwright, Power Automate, FlaUI, WinAppDriver, AutoHotkey, Tauri, PyO3, MCP SDK ecosystem
2. Dispatched 3 specialist agents for comparative analysis:
   - rust-pro: Windows-rs, uiautomation-rs, PyO3 migration strategy, COM threading
   - backend-architect: Playwright/FlaUI/PAD/WinAppDriver/AutoHotkey/Tauri comparison, integration patterns
   - performance-engineer: Latency budget analysis, caching strategies, optimization roadmap
3. Synthesized findings into comprehensive comparison document
4. Updated REVIEW.md with Section 7 (Comparative Framework Analysis)
5. Updated TODO.md with P1.5/P2.5/P3.5/P4.5 priority sections
6. Updated USER_MANUAL.md with Sections 10-11 (Performance, Roadmap) and extended framework comparison
7. Saved updated context and slices

## Key Findings

### Critical Issues (from Session 1)
- No auth on SSE/HTTP transport
- Unrestricted shell execution (no sandboxing)
- Analytics decorator bug: captures None at decoration time
- COM objects shared across thread apartment boundaries
- PIL ImageDraw used from ThreadPoolExecutor (not thread-safe)

### Performance Analysis (from Session 2)
- **`pg.PAUSE = 1.0`** is the single largest performance issue (1-6s per input op)
- **UIA CacheRequest suboptimal**: per-node `BuildUpdatedCache` defeats caching purpose
- **PowerShell subprocess**: 200-500ms per call for 7+ operations with Python alternatives
- **comtypes overhead**: ~50-200us per COM call x 10,000 calls = ~1000ms waste
- **Full Snapshot latency**: 630-5440ms current -> 275-1015ms (Python optimized) -> 83-290ms (Rust)

### Framework Comparison (from Session 2)
- **Playwright**: Complementary (browser-only). Bridge integration recommended.
- **FlaUI**: Same UIA foundation, superior pattern-based interaction. Adopt InvokePattern/ValuePattern.
- **Power Automate Desktop**: Enterprise integration via Dataverse REST API.
- **WinAppDriver**: Abandoned (2021). Validates Windows-MCP's approach.
- **AutoHotkey**: Win32 message fallback + WinWait pattern should be adopted.
- **Tauri 2.0**: Architectural model for capability manifests and Rust hybrid core.

### Recommended Architecture
```
FastMCP (Python) <- Capability Manifest
    |
    +-- Playwright Bridge (browser) -> playwright-mcp
    +-- Interaction Layer (Find, Invoke, WaitFor, Coordinate Fallback, Win32 Fallback)
    +-- State Layer (Rust UIA Traversal via PyO3, DOM Mode, WatchDog Events)
    +-- System Layer (Shell, File, Registry, Process, SystemInfo, Clipboard)
    +-- Enterprise Layer (PAD Flow Trigger, Session Recording)
```

## Agent Work Registry

| Agent | Task | Status | Duration | Key Findings |
|-------|------|--------|----------|-------------|
| python-pro | Code quality + performance | Complete | ~200s | COM threading violations, PIL race, serial subprocess, unbounded pools |
| security-auditor | OWASP security audit | Complete | ~150s | 2C + 4H + 3M + 3L findings |
| architect-reviewer | Architecture review | Complete | ~128s | Desktop God Object, analytics bug, ipykernel bloat, SOLID violations |
| rust-pro | Windows Rust API comparison | Complete | ~345s | windows-rs COM overhead analysis, PyO3 hybrid architecture, migration strategy |
| backend-architect | Framework comparison | Complete | ~291s | Playwright/FlaUI/PAD/WinAppDriver/AutoHotkey/Tauri deep comparison |
| performance-engineer | Performance optimization | Complete | ~187s | Latency budget, caching strategy, 4-phase optimization roadmap |

## Recommended Next Agents

1. **python-pro**: Implement Phase 1 quick wins (pg.PAUSE, ImageDraw fix, TreeScope_Subtree, PowerShell elimination)
2. **rust-pro**: Scaffold PyO3 extension crate for tree traversal hot path
3. **test-automator**: Add integration tests for MCP tool handlers before optimization work

## Files Created/Modified

| File | Action | Purpose |
|------|--------|---------|
| REVIEW.md | Updated | Added Section 7: Comparative Framework Analysis |
| TODO.md | Updated | Added P1.5/P2.5/P3.5/P4.5 priority sections from Phase 2 |
| USER_MANUAL.md | Updated | Added Sections 10-11, extended framework comparison |
| .claude/context/windows-mcp-context-20260218.md | Updated | Full project context (this file) |
| .claude/context/slices/rust-slice.md | Created | Rust migration context |
| .claude/context/slices/performance-slice.md | Created | Performance optimization context |

## Tech Stack

- Python 3.13+ with uv package manager
- FastMCP (MCP server framework)
- comtypes (COM interop for UIAutomation)
- pyautogui + pywin32 (input simulation)
- psutil (process management)
- requests + markdownify (web scraping)
- PostHog (telemetry, opt-out available)
- Hatchling (build backend)
- ruff (linting/formatting)
- pytest + pytest-asyncio (testing)

## Planned Tech Stack Additions
- PyO3 + Maturin (Rust-Python FFI)
- windows-rs (Rust Windows COM bindings)
- rayon (Rust parallel iterators)
- image crate (Rust screenshot annotation)
- mss (Python DirectX screen capture)
