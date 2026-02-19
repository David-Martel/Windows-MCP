# Python Context Slice -- Windows-MCP

## Key Performance Issues
- `desktop/service.py` -- Serial PowerShell subprocess per operation (~200-500ms each)
- `tree/service.py` -- ThreadPoolExecutor() without max_workers, recursive traversal (no depth guard)
- `desktop/service.py:875` -- PIL ImageDraw in ThreadPoolExecutor (NOT thread-safe)
- `pyautogui.PAUSE = 1.0` global pause after every action
- `analytics.py:97` -- print() to stdout in production (interleaves with MCP protocol)

## Critical Bugs
- Analytics decorator captures None at decoration time (telemetry does nothing)
- COM objects shared across thread apartment boundaries (tree/service.py:583-584)
- _AutomationClient singleton has TOCTOU race (uia/core.py:53-57)
- analytics.py:107 -- broken traceback (both branches produce str(error))

## Code Quality
- Desktop class: 1087 lines, God Object (6-7 responsibilities)
- 13 constants copy-pasted across uia/core.py, controls.py, patterns.py
- Boolean coercion pattern repeated 6+ times (needs utility function)
- ipykernel unused dependency in pyproject.toml

## Well-Designed Modules
- filesystem/service.py -- stateless pure functions, clean error handling
- tree/ module -- proper config/cache/views/service separation
- uia/ layer -- solid COM abstraction
