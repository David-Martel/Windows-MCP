# Test Context Slice -- Windows-MCP (2026-02-20)

## Test Infrastructure
- Framework: pytest + pytest-asyncio (asyncio_mode="auto")
- Run: `uv run python -m pytest tests/ -m "not live_desktop"`
- Full suite: 2115 tests, ~23s
- Live tests: 13 tests (test_live_com.py + test_live_patterns.py), requires running desktop

## Key Test Files

| File | Tests | Coverage Area |
|------|-------|--------------|
| test_mcp_integration.py | ~167 | MCP tool dispatch via FastMCP |
| test_main_coverage.py | ~90 | Tool edge cases, error paths |
| test_input_service.py | ~129 | InputService, ValuePattern |
| test_tree_service.py | ~58 | Tree traversal, focus change, DOM correction |
| test_get_state.py | ~60 | Desktop.get_state orchestration |
| test_analytics.py | ~178 | Telemetry, audit, rate limiting, permissions |
| test_classify_rust_tree.py | ~61 | Rust tree element classification |
| test_native_layer_gaps.py | ~43 | Native adapter imports/fallbacks |
| test_desktop_methods.py | ~80 | Desktop facade methods |
| test_process_service.py | ~51 | ProcessService list/kill/protected |
| test_vdm_service.py | ~52 | VDM desktop enumeration |
| test_window_service.py | ~82 | WindowService operations |
| test_screen_service.py | ~53 | ScreenService screenshots |

## Test Patterns

### Rust Fast-Path Isolation
When testing Python UIA fallback behavior, patch native functions to None:

**Single function (class-level @patch):**
```python
@patch("windows_mcp.native.native_find_elements", return_value=None)
class TestFindToolDispatch:
    async def test_xxx(self, _native, patched_desktop):
        ...
```

**Multiple functions (autouse fixture):**
```python
class TestInvokeToolDispatch:
    @pytest.fixture(autouse=True)
    def _disable_native_patterns(self):
        with patch.multiple("windows_mcp.native",
            native_invoke_at=MagicMock(return_value=None),
            native_toggle_at=MagicMock(return_value=None),
            ...
        ):
            yield
```

### MCP Tool Testing
```python
tools = await _get_tools()  # Returns dict of tool name -> tool object
result = await tools["Find"].fn(name="OK")
```

### Service Testing
- Patch `_state_module.desktop` (not `main_module.desktop`)
- Use `_make_desktop_state()` factories for consistent test data
- COM mocks: `SimpleNamespace` or `MagicMock` with cached properties

## Markers
- `live_desktop`: Requires running Windows desktop with UIA
- Default CI excludes live tests: `-m "not live_desktop"`
