# Performance Context Slice -- Windows-MCP

## Top Bottlenecks (Ranked by Impact)

### 1. `pg.PAUSE = 1.0` (desktop/service.py:45)
- 1-second mandatory sleep after EVERY pyautogui call
- `type(clear=True, press_enter=True)` = 6 seconds of pure sleep
- Fix: Change to 0.05, add targeted delays where needed

### 2. UIA CacheRequest Suboptimal (tree/service.py)
- `BuildUpdatedCache` called per-node with `TreeScope_Element` + `TreeScope_Children`
- TWO cross-process COM round-trips per tree node
- Fix: Single `TreeScope_Subtree` on window root, walk cached result in-process
- Impact: 60-80% tree traversal reduction

### 3. PowerShell Subprocess (desktop/service.py:209)
- ~200-500ms per call (process creation + CLR init + PS engine start)
- Used for: registry, sysinfo, notifications, locale, app launch
- Fix: `winreg` module, `platform`/`locale` stdlib, WinRT for toasts

### 4. PIL ImageDraw Thread Safety Bug (desktop/service.py:875)
- `ImageDraw` shared across `ThreadPoolExecutor` without locks
- Not thread-safe -- undefined behavior
- Fix: Remove ThreadPoolExecutor (sequential drawing is <5ms)

### 5. comtypes COM Overhead (tree traversal)
- ~50-200us per COM call in Python overhead
- 10,000 calls per Snapshot = ~1000ms of pure Python/comtypes waste
- Fix: Rust PyO3 extension with windows-rs (near-zero overhead)

### 6. LegacyIAccessiblePattern (tree/service.py:439-482)
- Called up to 3 times per interactive element (live COM round-trip each time)
- Fix: Call once, store in local variable

## Latency Budget (Snapshot)

| Component | Current | Optimized Python | Rust PyO3 |
|-----------|---------|-----------------|-----------|
| Tree traversal | 500-5000ms | 200-800ms | 50-200ms |
| Screenshot | 55-150ms | 30-80ms | 12-35ms |
| VDM queries | 15-70ms | 5-15ms | 5-15ms |
| Window enum | 55-210ms | 35-110ms | 13-35ms |
| **TOTAL** | **630-5440ms** | **275-1015ms** | **83-290ms** |

## Caching Opportunities
- Per-window tree cache (WatchDog invalidation): 80-95% reduction on repeats
- Start Menu app cache (filesystem watcher): 200-500ms per launch saved
- VDM desktop list cache (5s TTL): 10-50ms per get_state
- Element coordinate cache (bounding box reuse): 5-50ms per click/type

## Phase 1 Quick Wins (Python-only, 1-2 weeks)
1. `pg.PAUSE = 0.05` (1 line, 1-6s saved per input)
2. Fix ImageDraw thread safety (30 min)
3. PowerShell -> winreg/stdlib (2-4 hours, 200-500ms per op)
4. TreeScope_Subtree cache (1 day, 60-80% tree reduction)
5. Deduplicate LegacyIAccessiblePattern calls (2 hours, 10-20% per element)
