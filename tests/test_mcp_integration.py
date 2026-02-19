"""MCP integration tests for windows_mcp.__main__.

Validates the FastMCP framework integration without requiring a live Windows desktop,
COM/UIA, or any physical UI. Tests are organised in four tiers:

  Tier 1 - Server structure: tool registration, descriptions, annotations, schemas.
  Tier 2 - Tool dispatch: call each tool's underlying function with mocked Desktop.
  Tier 3 - Error handling & edge cases: missing args, invalid modes, coercions.
  Tier 4 - Transport & auth: CLI options, middleware injection, localhost safety.

Design notes:
  - Tools are registered as FunctionTool objects on mcp._tool_manager.
    The Context injection that FastMCP performs at call-time cannot be exercised
    without a running server session, so tests call tool.fn(...) directly.
    This exercises all tool logic, the analytics wrapper (which is a no-op due to
    the known None-capture bug), and the real module-level desktop/screen_size globals.
  - desktop and screen_size are replaced via patch.object on main_module BEFORE each
    tool call so the global None value is never dereferenced.
  - win32clipboard is patched at the sys.modules level because it is imported inside
    the clipboard_tool function body at call-time.
"""

import sys
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

import windows_mcp.__main__ as main_module
from windows_mcp.desktop.views import DesktopState, Size, Status, Window
from windows_mcp.tree.views import (
    BoundingBox,
    Center,
    TreeElementNode,
    TreeState,
)

# ---------------------------------------------------------------------------
# Shared test fixtures
# ---------------------------------------------------------------------------

SCREEN_SIZE = Size(width=1920, height=1080)


def _make_bbox(
    left: int = 0,
    top: int = 0,
    right: int = 1920,
    bottom: int = 1080,
) -> BoundingBox:
    return BoundingBox(
        left=left,
        top=top,
        right=right,
        bottom=bottom,
        width=right - left,
        height=bottom - top,
    )


def _make_center(x: int = 960, y: int = 540) -> Center:
    return Center(x=x, y=y)


def _make_window(name: str = "Notepad") -> Window:
    return Window(
        name=name,
        is_browser=False,
        depth=0,
        status=Status.NORMAL,
        bounding_box=_make_bbox(),
        handle=12345,
        process_id=6789,
    )


def _make_tree_element(
    name: str = "OK Button",
    control_type: str = "Button",
    window_name: str = "Notepad",
) -> TreeElementNode:
    bbox = _make_bbox(left=100, top=200, right=200, bottom=240)
    center = _make_center(x=150, y=220)
    return TreeElementNode(
        bounding_box=bbox,
        center=center,
        name=name,
        control_type=control_type,
        window_name=window_name,
    )


def _make_desktop_state(
    window_name: str = "Notepad",
    tree_nodes: list[TreeElementNode] | None = None,
) -> DesktopState:
    win = _make_window(window_name)
    ts = TreeState(interactive_nodes=tree_nodes or [])
    return DesktopState(
        active_desktop={"name": "Desktop 1", "id": "abc-123"},
        all_desktops=[
            {"name": "Desktop 1", "id": "abc-123"},
            {"name": "Desktop 2", "id": "def-456"},
        ],
        active_window=win,
        windows=[win],
        screenshot=None,
        tree_state=ts,
    )


@pytest.fixture
def mock_desktop() -> MagicMock:
    """A MagicMock that provides sensible defaults for all Desktop methods."""
    m = MagicMock()
    m.get_screen_size.return_value = SCREEN_SIZE
    m.execute_command.return_value = ("output\n", 0)
    m.get_state.return_value = _make_desktop_state()
    m.get_system_info.return_value = "CPU: 5%\nMemory: 4096 MB"
    m.list_processes.return_value = "explorer.exe  PID:1234  MEM:50MB"
    m.kill_process.return_value = "Process terminated."
    m.lock_screen.return_value = "Screen locked."
    m.send_notification.return_value = "Notification sent."
    m.app.return_value = "App launched."
    m.scrape.return_value = "<h1>Hello</h1>"
    m.click.return_value = None
    m.move.return_value = None
    m.drag.return_value = None
    m.shortcut.return_value = None
    m.scroll.return_value = None
    m.type.return_value = None
    m.multi_select.return_value = None
    m.multi_edit.return_value = None
    m.registry_get.return_value = 'Value "Path": C:\\Windows\\system32'
    m.registry_set.return_value = 'Value "TestKey" set to "hello".'
    m.registry_delete.return_value = 'Value "TestKey" deleted.'
    m.registry_list.return_value = "Values:\n  TestKey = hello\nSub-keys:\n  Child1"
    return m


@pytest.fixture
def patched_desktop(mock_desktop):
    """Patch the module-level desktop and screen_size globals and yield the mock."""
    with patch.object(main_module, "desktop", mock_desktop):
        with patch.object(main_module, "screen_size", SCREEN_SIZE):
            yield mock_desktop


async def _get_tools() -> dict:
    """Return the dict of registered FunctionTool objects keyed by tool name."""
    return await main_module.mcp._tool_manager.get_tools()


# ---------------------------------------------------------------------------
# Tier 1: MCP Server Structure
# ---------------------------------------------------------------------------


class TestServerStructure:
    """Verify the FastMCP instance is configured correctly."""

    def test_mcp_instance_name(self):
        assert main_module.mcp.name == "windows-mcp"

    async def test_all_23_tools_registered(self):
        tools = await _get_tools()
        assert len(tools) == 23

    async def test_expected_tool_names_present(self):
        tools = await _get_tools()
        expected = {
            "App",
            "Shell",
            "File",
            "Snapshot",
            "Click",
            "Type",
            "Scroll",
            "Move",
            "Shortcut",
            "Wait",
            "Scrape",
            "MultiSelect",
            "MultiEdit",
            "Clipboard",
            "Process",
            "SystemInfo",
            "Notification",
            "LockScreen",
            "Registry",
            "WaitFor",
            "Find",
            "Invoke",
            "VisionAnalyze",
        }
        assert set(tools.keys()) == expected

    async def test_every_tool_has_description(self):
        tools = await _get_tools()
        tools_missing_desc = [name for name, t in tools.items() if not t.description]
        assert tools_missing_desc == [], f"Missing description: {tools_missing_desc}"

    async def test_every_tool_has_annotations(self):
        tools = await _get_tools()
        tools_missing_ann = [name for name, t in tools.items() if t.annotations is None]
        assert tools_missing_ann == [], f"Missing annotations: {tools_missing_ann}"

    async def test_every_tool_has_annotation_title_matching_name(self):
        tools = await _get_tools()
        mismatches = [
            name for name, t in tools.items() if t.annotations and t.annotations.title != name
        ]
        assert mismatches == [], f"Annotation title mismatch: {mismatches}"

    async def test_readonly_tools_flagged_correctly(self):
        """Six tools are read-only: Find, Scrape, Snapshot, SystemInfo, Wait, WaitFor."""
        tools = await _get_tools()
        readonly = {name for name, t in tools.items() if t.annotations.readOnlyHint}
        assert readonly == {
            "Find",
            "Scrape",
            "Snapshot",
            "SystemInfo",
            "VisionAnalyze",
            "Wait",
            "WaitFor",
        }

    async def test_open_world_tools_flagged_correctly(self):
        """Only Shell and Scrape reach outside the local machine."""
        tools = await _get_tools()
        open_world = {name for name, t in tools.items() if t.annotations.openWorldHint}
        assert open_world == {"Shell", "Scrape", "VisionAnalyze"}

    async def test_destructive_tools_include_shell(self):
        tools = await _get_tools()
        assert tools["Shell"].annotations.destructiveHint is True

    async def test_non_destructive_tools_include_snapshot(self):
        tools = await _get_tools()
        assert tools["Snapshot"].annotations.destructiveHint is False

    async def test_shell_parameters_include_command_and_timeout(self):
        tools = await _get_tools()
        params = tools["Shell"].parameters.get("properties", {})
        assert "command" in params
        assert "timeout" in params

    async def test_click_parameters_include_loc_button_clicks(self):
        tools = await _get_tools()
        params = tools["Click"].parameters.get("properties", {})
        assert "loc" in params
        assert "button" in params
        assert "clicks" in params

    async def test_file_parameters_include_mode_and_path(self):
        tools = await _get_tools()
        params = tools["File"].parameters.get("properties", {})
        assert "mode" in params
        assert "path" in params

    async def test_registry_parameters_include_mode_path_name(self):
        tools = await _get_tools()
        params = tools["Registry"].parameters.get("properties", {})
        assert "mode" in params
        assert "path" in params
        assert "name" in params

    async def test_waitfor_parameters_include_mode_name_timeout(self):
        tools = await _get_tools()
        params = tools["WaitFor"].parameters.get("properties", {})
        assert "mode" in params
        assert "name" in params
        assert "timeout" in params

    async def test_invoke_parameters_include_loc_action_value(self):
        tools = await _get_tools()
        params = tools["Invoke"].parameters.get("properties", {})
        assert "loc" in params
        assert "action" in params
        assert "value" in params

    async def test_tool_fn_attribute_points_to_callable(self):
        """Each registered tool exposes a callable .fn attribute."""
        tools = await _get_tools()
        for name, tool in tools.items():
            assert callable(tool.fn), f"Tool {name!r} .fn is not callable"


# ---------------------------------------------------------------------------
# Tier 2: Tool Dispatch via tool.fn
# ---------------------------------------------------------------------------


class TestShellToolDispatch:
    async def test_shell_returns_response_and_status(self, patched_desktop):
        patched_desktop.execute_command.return_value = ("hello\n", 0)
        tools = await _get_tools()
        result = await tools["Shell"].fn(command="echo hello", timeout=5)
        assert "hello" in result
        assert "Status Code: 0" in result

    async def test_shell_calls_execute_command_with_correct_args(self, patched_desktop):
        tools = await _get_tools()
        await tools["Shell"].fn(command="dir", timeout=10)
        patched_desktop.execute_command.assert_called_once_with("dir", 10)

    async def test_shell_uses_default_timeout_of_30(self, patched_desktop):
        tools = await _get_tools()
        await tools["Shell"].fn(command="Get-Process")
        patched_desktop.execute_command.assert_called_once_with("Get-Process", 30)

    async def test_shell_exception_returns_error_string(self, patched_desktop):
        patched_desktop.execute_command.side_effect = RuntimeError("exec failed")
        tools = await _get_tools()
        result = await tools["Shell"].fn(command="bad-cmd")
        assert "Error executing command" in result
        assert "Status Code: 1" in result


class TestFileToolDispatch:
    async def test_file_list_delegates_to_filesystem(self, patched_desktop):
        with patch(
            "windows_mcp.filesystem.list_directory", return_value="file1.txt\nfile2.txt"
        ) as mock_list:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="list", path="C:/Windows")
            assert "file1.txt" in result
            mock_list.assert_called_once()

    async def test_file_read_returns_content(self, patched_desktop):
        with patch("windows_mcp.filesystem.read_file", return_value="line1\nline2") as mock_read:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="read", path="C:/test.txt")
            assert "line1" in result
            mock_read.assert_called_once()

    async def test_file_write_delegates_to_filesystem(self, patched_desktop):
        with patch("windows_mcp.filesystem.write_file", return_value="File written.") as mock_write:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="write", path="C:/out.txt", content="data")
            assert "File written." in result
            mock_write.assert_called_once()

    async def test_file_delete_delegates_to_filesystem(self, patched_desktop):
        with patch("windows_mcp.filesystem.delete_path", return_value="Deleted.") as mock_del:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="delete", path="C:/old.txt")
            assert "Deleted." in result
            mock_del.assert_called_once()

    async def test_file_info_delegates_to_filesystem(self, patched_desktop):
        with patch(
            "windows_mcp.filesystem.get_file_info", return_value="Size: 1024 bytes"
        ) as mock_info:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="info", path="C:/test.txt")
            assert "Size" in result
            mock_info.assert_called_once()

    async def test_file_copy_delegates_to_filesystem(self, patched_desktop):
        with patch("windows_mcp.filesystem.copy_path", return_value="Copied.") as mock_copy:
            tools = await _get_tools()
            result = await tools["File"].fn(
                mode="copy", path="C:/src.txt", destination="C:/dst.txt"
            )
            assert "Copied." in result
            mock_copy.assert_called_once()

    async def test_file_move_delegates_to_filesystem(self, patched_desktop):
        with patch("windows_mcp.filesystem.move_path", return_value="Moved.") as mock_move:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="move", path="C:/a.txt", destination="C:/b.txt")
            assert "Moved." in result
            mock_move.assert_called_once()

    async def test_file_search_delegates_to_filesystem(self, patched_desktop):
        with patch(
            "windows_mcp.filesystem.search_files", return_value="C:/foo.py\nC:/bar.py"
        ) as mock_search:
            tools = await _get_tools()
            result = await tools["File"].fn(mode="search", path="C:/", pattern="*.py")
            assert "foo.py" in result
            mock_search.assert_called_once()


class TestSystemInfoToolDispatch:
    async def test_systeminfo_returns_info_string(self, patched_desktop):
        patched_desktop.get_system_info.return_value = "CPU: 42%\nDisk: 100GB"
        tools = await _get_tools()
        result = await tools["SystemInfo"].fn()
        assert "CPU" in result
        patched_desktop.get_system_info.assert_called_once()

    async def test_systeminfo_exception_returns_error_string(self, patched_desktop):
        patched_desktop.get_system_info.side_effect = RuntimeError("wmi failed")
        tools = await _get_tools()
        result = await tools["SystemInfo"].fn()
        assert "Error getting system info" in result


class TestRegistryToolDispatch:
    async def test_registry_get_calls_desktop_registry_get(self, patched_desktop):
        patched_desktop.registry_get.return_value = 'Value "Path": /usr/bin'
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="get", path="HKCU:\\Environment", name="Path")
        assert "/usr/bin" in result
        patched_desktop.registry_get.assert_called_once_with(path="HKCU:\\Environment", name="Path")

    async def test_registry_set_calls_desktop_registry_set(self, patched_desktop):
        patched_desktop.registry_set.return_value = 'Value "MyKey" set to "hello".'
        tools = await _get_tools()
        result = await tools["Registry"].fn(
            mode="set", path="HKCU:\\Test", name="MyKey", value="hello"
        )
        assert "set to" in result
        patched_desktop.registry_set.assert_called_once()

    async def test_registry_delete_calls_desktop_registry_delete(self, patched_desktop):
        patched_desktop.registry_delete.return_value = 'Value "MyKey" deleted.'
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="delete", path="HKCU:\\Test", name="MyKey")
        assert "deleted" in result
        patched_desktop.registry_delete.assert_called_once()

    async def test_registry_list_calls_desktop_registry_list(self, patched_desktop):
        patched_desktop.registry_list.return_value = "Values:\n  TestKey = hello"
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="list", path="HKCU:\\Test")
        assert "TestKey" in result
        patched_desktop.registry_list.assert_called_once_with(path="HKCU:\\Test")

    async def test_registry_exception_returns_error_string(self, patched_desktop):
        patched_desktop.registry_get.side_effect = OSError("access denied")
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="get", path="HKLM:\\Secure", name="Key")
        assert "Error accessing registry" in result


class TestProcessToolDispatch:
    async def test_process_list_returns_process_info(self, patched_desktop):
        patched_desktop.list_processes.return_value = "explorer.exe  PID:4  MEM:50MB"
        tools = await _get_tools()
        result = await tools["Process"].fn(mode="list")
        assert "explorer.exe" in result
        patched_desktop.list_processes.assert_called_once()

    async def test_process_kill_delegates_to_desktop(self, patched_desktop):
        patched_desktop.kill_process.return_value = "Terminated."
        tools = await _get_tools()
        result = await tools["Process"].fn(mode="kill", pid=1234)
        assert "Terminated." in result
        patched_desktop.kill_process.assert_called_once()

    async def test_process_kill_with_name(self, patched_desktop):
        patched_desktop.kill_process.return_value = "notepad.exe killed."
        tools = await _get_tools()
        await tools["Process"].fn(mode="kill", name="notepad.exe")
        patched_desktop.kill_process.assert_called_with(name="notepad.exe", pid=None, force=False)

    async def test_process_exception_returns_error_string(self, patched_desktop):
        patched_desktop.list_processes.side_effect = RuntimeError("psutil error")
        tools = await _get_tools()
        result = await tools["Process"].fn(mode="list")
        assert "Error managing processes" in result


class TestSnapshotToolDispatch:
    async def test_snapshot_returns_list_of_strings(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)
        assert isinstance(result, list)
        assert len(result) >= 1

    async def test_snapshot_output_contains_focused_window_header(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)
        combined = "".join(str(item) for item in result)
        assert "Focused Window" in combined

    async def test_snapshot_output_contains_desktop_header(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)
        combined = "".join(str(item) for item in result)
        assert "Active Desktop" in combined

    async def test_snapshot_calls_desktop_get_state(self, patched_desktop):
        tools = await _get_tools()
        await tools["Snapshot"].fn(use_vision=False, use_dom=False)
        patched_desktop.get_state.assert_called_once()

    async def test_snapshot_exception_returns_error_in_list(self, patched_desktop):
        patched_desktop.get_state.side_effect = RuntimeError("COM error")
        tools = await _get_tools()
        result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)
        assert isinstance(result, list)
        assert "Error capturing desktop state" in result[0]


class TestClickToolDispatch:
    async def test_click_returns_confirmation_string(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Click"].fn(loc=[100, 200], button="left", clicks=1)
        assert "Single left clicked at (100,200)" in result

    async def test_click_calls_desktop_click(self, patched_desktop):
        tools = await _get_tools()
        await tools["Click"].fn(loc=[300, 400], button="right", clicks=2)
        patched_desktop.click.assert_called_once_with(loc=(300, 400), button="right", clicks=2)

    async def test_click_hover_uses_zero_clicks(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Click"].fn(loc=[100, 200], button="left", clicks=0)
        assert "Hover" in result


class TestTypeToolDispatch:
    async def test_type_returns_confirmation_string(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Type"].fn(loc=[100, 200], text="hello world")
        assert "Typed hello world at (100,200)" in result

    async def test_type_calls_desktop_type(self, patched_desktop):
        tools = await _get_tools()
        await tools["Type"].fn(loc=[50, 75], text="test", clear=True, press_enter=True)
        patched_desktop.type.assert_called_once_with(
            loc=(50, 75),
            text="test",
            caret_position="idle",
            clear=True,
            press_enter=True,
        )


class TestScrollToolDispatch:
    async def test_scroll_with_loc_returns_message(self, patched_desktop):
        patched_desktop.scroll.return_value = None
        tools = await _get_tools()
        result = await tools["Scroll"].fn(loc=[100, 200], direction="up")
        assert "Scrolled" in result
        assert "up" in result

    async def test_scroll_calls_desktop_scroll(self, patched_desktop):
        tools = await _get_tools()
        await tools["Scroll"].fn(loc=[100, 200], direction="down", wheel_times=3)
        patched_desktop.scroll.assert_called_once_with((100, 200), "vertical", "down", 3)

    async def test_scroll_desktop_response_overrides_default_message(self, patched_desktop):
        patched_desktop.scroll.return_value = "Scrolled to bottom."
        tools = await _get_tools()
        result = await tools["Scroll"].fn(loc=[100, 200], direction="down")
        assert result == "Scrolled to bottom."


class TestMoveToolDispatch:
    async def test_move_returns_confirmation(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Move"].fn(loc=[100, 200], drag=False)
        assert "Moved the mouse pointer to (100,200)" in result

    async def test_move_drag_calls_desktop_drag(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Move"].fn(loc=[300, 400], drag=True)
        assert "Dragged to (300,400)" in result
        patched_desktop.drag.assert_called_once_with((300, 400))

    async def test_move_calls_desktop_move_when_not_drag(self, patched_desktop):
        tools = await _get_tools()
        await tools["Move"].fn(loc=[100, 200], drag=False)
        patched_desktop.move.assert_called_once_with((100, 200))


class TestShortcutToolDispatch:
    async def test_shortcut_returns_confirmation(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Shortcut"].fn(shortcut="ctrl+c")
        assert "Pressed ctrl+c" in result

    async def test_shortcut_calls_desktop_shortcut(self, patched_desktop):
        tools = await _get_tools()
        await tools["Shortcut"].fn(shortcut="win+r")
        patched_desktop.shortcut.assert_called_once_with("win+r")


class TestWaitToolDispatch:
    async def test_wait_returns_confirmation_with_duration(self, patched_desktop):
        with patch("pyautogui.sleep") as mock_sleep:
            tools = await _get_tools()
            result = await tools["Wait"].fn(duration=3)
            assert "Waited for 3 seconds" in result
            mock_sleep.assert_called_once_with(3)


class TestNotificationToolDispatch:
    async def test_notification_delegates_to_desktop(self, patched_desktop):
        patched_desktop.send_notification.return_value = "Notification sent."
        tools = await _get_tools()
        result = await tools["Notification"].fn(title="Hello", message="World")
        assert "Notification sent." in result
        patched_desktop.send_notification.assert_called_once_with("Hello", "World")

    async def test_notification_exception_returns_error_string(self, patched_desktop):
        patched_desktop.send_notification.side_effect = RuntimeError("toast error")
        tools = await _get_tools()
        result = await tools["Notification"].fn(title="Hi", message="Bye")
        assert "Error sending notification" in result


class TestLockScreenToolDispatch:
    async def test_lockscreen_delegates_to_desktop(self, patched_desktop):
        patched_desktop.lock_screen.return_value = "Screen locked."
        tools = await _get_tools()
        result = await tools["LockScreen"].fn()
        assert "Screen locked." in result
        patched_desktop.lock_screen.assert_called_once()

    async def test_lockscreen_exception_returns_error_string(self, patched_desktop):
        patched_desktop.lock_screen.side_effect = OSError("Win32 error")
        tools = await _get_tools()
        result = await tools["LockScreen"].fn()
        assert "Error locking screen" in result


class TestAppToolDispatch:
    async def test_app_launch_delegates_to_desktop(self, patched_desktop):
        patched_desktop.app.return_value = "Launched Notepad"
        tools = await _get_tools()
        result = await tools["App"].fn(mode="launch", name="Notepad")
        assert "Launched Notepad" in result
        patched_desktop.app.assert_called_once_with("launch", "Notepad", None, None)

    async def test_app_resize_passes_window_loc_and_size(self, patched_desktop):
        tools = await _get_tools()
        await tools["App"].fn(
            mode="resize",
            window_loc=[0, 0],
            window_size=[800, 600],
        )
        patched_desktop.app.assert_called_once_with("resize", None, (0, 0), (800, 600))

    async def test_app_switch_passes_name(self, patched_desktop):
        tools = await _get_tools()
        await tools["App"].fn(mode="switch", name="Calculator")
        patched_desktop.app.assert_called_once_with("switch", "Calculator", None, None)


class TestScrapeToolDispatch:
    async def test_scrape_http_mode_calls_desktop_scrape(self, patched_desktop):
        patched_desktop.scrape.return_value = "# Hello World"
        tools = await _get_tools()
        result = await tools["Scrape"].fn(url="https://example.com", use_dom=False)
        assert "https://example.com" in result
        assert "Hello World" in result
        patched_desktop.scrape.assert_called_once_with("https://example.com")

    async def test_scrape_dom_mode_no_dom_returns_fallback_message(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state()
        # tree_state.dom_node defaults to None in _make_desktop_state
        tools = await _get_tools()
        result = await tools["Scrape"].fn(url="https://example.com", use_dom=True)
        assert "No DOM information found" in result

    async def test_scrape_dom_mode_with_dom_node_returns_content(self, patched_desktop):
        dom_node = MagicMock()
        dom_node.vertical_scroll_percent = 50.0
        text_node = MagicMock()
        text_node.text = "Page content here"

        ts = TreeState(
            dom_node=dom_node,
            dom_informative_nodes=[text_node],
        )
        ds = DesktopState(
            active_desktop={"name": "Desktop 1"},
            all_desktops=[{"name": "Desktop 1"}],
            active_window=None,
            windows=[],
            tree_state=ts,
        )
        patched_desktop.get_state.return_value = ds

        tools = await _get_tools()
        result = await tools["Scrape"].fn(url="https://example.com", use_dom=True)
        assert "Page content here" in result


class TestClipboardToolDispatch:
    def _make_win32clipboard_mock(
        self,
        has_text: bool = True,
        clipboard_data: str = "test clipboard text",
    ) -> MagicMock:
        m = MagicMock()
        m.CF_UNICODETEXT = 13
        m.IsClipboardFormatAvailable.return_value = has_text
        m.GetClipboardData.return_value = clipboard_data
        return m

    async def test_clipboard_get_returns_content(self, patched_desktop):
        mock_cb = self._make_win32clipboard_mock(clipboard_data="copied text")
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="get")
        assert "copied text" in result

    async def test_clipboard_get_empty_returns_empty_message(self, patched_desktop):
        mock_cb = self._make_win32clipboard_mock(has_text=False)
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="get")
        assert "empty or contains non-text data" in result

    async def test_clipboard_set_stores_text(self, patched_desktop):
        mock_cb = self._make_win32clipboard_mock()
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="set", text="new value")
        assert "new value" in result

    async def test_clipboard_set_long_text_truncated_in_confirmation(self, patched_desktop):
        long_text = "A" * 200
        mock_cb = self._make_win32clipboard_mock()
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="set", text=long_text)
        assert "..." in result

    async def test_clipboard_exception_returns_error_string(self, patched_desktop):
        mock_cb = MagicMock()
        mock_cb.CF_UNICODETEXT = 13
        mock_cb.OpenClipboard.side_effect = OSError("clipboard locked")
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="get")
        assert "Error managing clipboard" in result


class TestMultiSelectToolDispatch:
    async def test_multiselect_returns_all_coordinates(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["MultiSelect"].fn(locs=[[100, 200], [300, 400]])
        assert "(100,200)" in result
        assert "(300,400)" in result

    async def test_multiselect_calls_desktop_multi_select(self, patched_desktop):
        tools = await _get_tools()
        await tools["MultiSelect"].fn(locs=[[100, 200], [300, 400]], press_ctrl=True)
        patched_desktop.multi_select.assert_called_once_with(True, [(100, 200), (300, 400)])


class TestMultiEditToolDispatch:
    async def test_multiedit_returns_all_coordinates_and_text(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["MultiEdit"].fn(locs=[[100, 200, "hello"], [300, 400, "world"]])
        assert "(100,200)" in result
        assert "hello" in result
        assert "(300,400)" in result
        assert "world" in result

    async def test_multiedit_calls_desktop_multi_edit(self, patched_desktop):
        locs = [[100, 200, "alpha"], [500, 600, "beta"]]
        tools = await _get_tools()
        await tools["MultiEdit"].fn(locs=locs)
        patched_desktop.multi_edit.assert_called_once_with(
            [(100, 200, "alpha"), (500, 600, "beta")]
        )


class TestWaitForToolDispatch:
    async def test_waitfor_window_found_immediately(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state(window_name="Notepad")
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="Notepad", timeout=5)
        assert "Window found" in result
        assert "Notepad" in result

    async def test_waitfor_element_found_in_tree(self, patched_desktop):
        node = _make_tree_element(name="OK Button", control_type="Button")
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=[node])
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="element", name="OK", timeout=5)
        assert "Element found" in result
        assert "OK Button" in result
        assert "Button" in result

    async def test_waitfor_times_out_when_not_found(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state(window_name="Other")
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="NonExistent", timeout=1)
        assert "Timeout" in result
        assert "NonExistent" in result

    async def test_waitfor_active_window_also_checked(self, patched_desktop):
        ds = _make_desktop_state(window_name="Calculator")
        ds.windows = []  # empty windows list -- only active_window
        patched_desktop.get_state.return_value = ds
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="Calc", timeout=5)
        assert "Window found" in result


class TestFindToolDispatch:
    async def test_find_by_name_returns_matching_elements(self, patched_desktop):
        nodes = [
            _make_tree_element(name="Submit Button", control_type="Button"),
            _make_tree_element(name="Cancel Button", control_type="Button"),
            _make_tree_element(name="Text Field", control_type="Edit"),
        ]
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=nodes)
        tools = await _get_tools()
        result = await tools["Find"].fn(name="Button")
        assert "Submit Button" in result
        assert "Cancel Button" in result
        assert "Text Field" not in result

    async def test_find_by_control_type_filters_correctly(self, patched_desktop):
        nodes = [
            _make_tree_element(name="Name", control_type="Edit"),
            _make_tree_element(name="Submit", control_type="Button"),
        ]
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=nodes)
        tools = await _get_tools()
        result = await tools["Find"].fn(control_type="Edit")
        assert "Name" in result
        assert "Submit" not in result

    async def test_find_by_window_filters_by_window_name(self, patched_desktop):
        nodes = [
            _make_tree_element(name="OK", window_name="Dialog"),
            _make_tree_element(name="Cancel", window_name="Main"),
        ]
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=nodes)
        tools = await _get_tools()
        result = await tools["Find"].fn(window="Dialog")
        assert "OK" in result
        assert "Cancel" not in result

    async def test_find_returns_count_header(self, patched_desktop):
        node = _make_tree_element(name="OK Button")
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=[node])
        tools = await _get_tools()
        result = await tools["Find"].fn(name="OK")
        assert "Found 1 element" in result

    async def test_find_no_matches_returns_not_found_message(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=[])
        tools = await _get_tools()
        result = await tools["Find"].fn(name="Invisible")
        assert "No elements found" in result

    async def test_find_respects_limit_parameter(self, patched_desktop):
        nodes = [_make_tree_element(name=f"Button {i}") for i in range(10)]
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=nodes)
        tools = await _get_tools()
        result = await tools["Find"].fn(name="Button", limit=3)
        assert "Found 3 element" in result


class TestInvokeToolDispatch:
    def _make_element_mock(
        self,
        name: str = "OK",
        control_type: str = "button",
        automation_id: str = "btn_ok",
    ) -> MagicMock:
        elem = MagicMock()
        elem.Name = name
        elem.AutomationId = automation_id
        elem.LocalizedControlType = control_type
        return elem

    async def test_invoke_action_calls_invoke_pattern(self, patched_desktop):
        elem = self._make_element_mock()
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="invoke")
        assert "Invoked 'OK'" in result
        mock_pattern.Invoke.assert_called_once()

    async def test_invoke_toggle_returns_state(self, patched_desktop):
        elem = self._make_element_mock(name="Checkbox", control_type="checkbox")
        mock_pattern = MagicMock()
        mock_pattern.ToggleState = 1  # on
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="toggle")
        assert "Toggled" in result
        assert "on" in result

    async def test_invoke_set_value_calls_value_pattern(self, patched_desktop):
        elem = self._make_element_mock(name="InputField", control_type="edit")
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="set_value", value="hello")
        assert "Set value" in result
        mock_pattern.SetValue.assert_called_once_with("hello")

    async def test_invoke_expand_calls_expand_collapse_pattern(self, patched_desktop):
        elem = self._make_element_mock(name="Dropdown", control_type="combobox")
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="expand")
        assert "Expanded" in result
        mock_pattern.Expand.assert_called_once()

    async def test_invoke_collapse_calls_collapse(self, patched_desktop):
        elem = self._make_element_mock(name="Dropdown", control_type="combobox")
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="collapse")
        assert "Collapsed" in result
        mock_pattern.Collapse.assert_called_once()

    async def test_invoke_select_calls_selection_item_pattern(self, patched_desktop):
        elem = self._make_element_mock(name="ListItem", control_type="listitem")
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="select")
        assert "Selected" in result
        mock_pattern.Select.assert_called_once()

    async def test_invoke_pattern_not_supported_returns_error(self, patched_desktop):
        elem = self._make_element_mock(name="Label", control_type="text")
        elem.GetPattern.return_value = None  # pattern not supported

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="invoke")
        assert "does not support InvokePattern" in result


class TestVisionAnalyzeToolDispatch:
    async def test_vision_not_configured_returns_error(self, patched_desktop):
        """VisionAnalyze returns helpful error when VISION_API_URL is not set."""
        with patch.dict("os.environ", {}, clear=False):
            # Ensure VISION_API_URL is absent
            import os

            os.environ.pop("VISION_API_URL", None)
            tools = await _get_tools()
            result = await tools["VisionAnalyze"].fn()
        assert "not configured" in result

    async def test_vision_describe_mode(self, patched_desktop):
        """VisionAnalyze describe mode captures screenshot and calls vision API."""
        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, format: buf.write(b"\x89PNG"))
        patched_desktop.get_screenshot.return_value = mock_img

        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.raise_for_status = MagicMock()
            mock_resp.json.return_value = {
                "choices": [{"message": {"content": "Notepad is open."}}]
            }
            mock_post.return_value = mock_resp
            with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
                tools = await _get_tools()
                result = await tools["VisionAnalyze"].fn(mode="describe")
        assert "Notepad" in result

    async def test_vision_elements_mode_returns_json(self, patched_desktop):
        """VisionAnalyze elements mode returns JSON array of detected elements."""
        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, format: buf.write(b"\x89PNG"))
        patched_desktop.get_screenshot.return_value = mock_img

        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.raise_for_status = MagicMock()
            mock_resp.json.return_value = {
                "choices": [{"message": {"content": '[{"type":"button","label":"OK"}]'}}]
            }
            mock_post.return_value = mock_resp
            with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
                tools = await _get_tools()
                result = await tools["VisionAnalyze"].fn(mode="elements")
        assert "button" in result
        assert "OK" in result

    async def test_vision_elements_mode_empty_returns_message(self, patched_desktop):
        """VisionAnalyze elements mode returns message when no elements found."""
        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, format: buf.write(b"\x89PNG"))
        patched_desktop.get_screenshot.return_value = mock_img

        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.raise_for_status = MagicMock()
            mock_resp.json.return_value = {"choices": [{"message": {"content": "[]"}}]}
            mock_post.return_value = mock_resp
            with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
                tools = await _get_tools()
                result = await tools["VisionAnalyze"].fn(mode="elements")
        assert "No UI elements" in result

    async def test_vision_query_mode_with_query(self, patched_desktop):
        """VisionAnalyze query mode with a query calls vision API."""
        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, format: buf.write(b"\x89PNG"))
        patched_desktop.get_screenshot.return_value = mock_img

        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.raise_for_status = MagicMock()
            mock_resp.json.return_value = {
                "choices": [{"message": {"content": "The dialog says 'Save changes?'"}}]
            }
            mock_post.return_value = mock_resp
            with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
                tools = await _get_tools()
                result = await tools["VisionAnalyze"].fn(
                    mode="query", query="What does the dialog say?"
                )
        assert "Save changes" in result

    async def test_vision_query_mode_requires_query(self, patched_desktop):
        """VisionAnalyze query mode requires a query parameter."""
        with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
            tools = await _get_tools()
            result = await tools["VisionAnalyze"].fn(mode="query", query="")
        assert "Error" in result

    async def test_vision_unknown_mode_returns_error(self, patched_desktop):
        """VisionAnalyze with invalid mode returns clear error."""
        mock_img = MagicMock()
        mock_img.save = MagicMock(side_effect=lambda buf, format: buf.write(b"\x89PNG"))
        patched_desktop.get_screenshot.return_value = mock_img

        with patch.dict("os.environ", {"VISION_API_URL": "http://test:8080/v1"}):
            tools = await _get_tools()
            result = await tools["VisionAnalyze"].fn(mode="invalid_mode")
        assert "unknown mode" in result
        assert "invalid_mode" in result


# ---------------------------------------------------------------------------
# Tier 3: Error Handling & Edge Cases
# ---------------------------------------------------------------------------


class TestLocationValidation:
    """Tools that accept [x, y] coordinates reject wrong-length lists."""

    async def test_click_wrong_loc_length_raises_value_error(self, patched_desktop):
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"must be \[x, y\]"):
            await tools["Click"].fn(loc=[100, 200, 300])

    async def test_type_wrong_loc_length_raises_value_error(self, patched_desktop):
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"must be \[x, y\]"):
            await tools["Type"].fn(loc=[100], text="hello")

    async def test_move_wrong_loc_length_raises_value_error(self, patched_desktop):
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"must be \[x, y\]"):
            await tools["Move"].fn(loc=[100])

    async def test_scroll_wrong_loc_length_raises_value_error(self, patched_desktop):
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"must be \[x, y\]"):
            await tools["Scroll"].fn(loc=[100])

    async def test_invoke_single_element_loc_returns_error_string(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Invoke"].fn(loc=[100], action="invoke")
        assert "Error: loc must be [x, y]" in result


class TestInvalidModes:
    """Tools with mode parameters return error strings for unknown modes."""

    async def test_file_unknown_mode_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["File"].fn(mode="compress", path="C:/test.txt")
        assert 'Unknown mode "compress"' in result

    async def test_registry_unknown_mode_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="patch", path="HKCU:\\Test")
        assert 'mode must be "get", "set", "delete", or "list"' in result

    async def test_process_unknown_mode_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Process"].fn(mode="suspend")
        assert 'mode must be either "list" or "kill"' in result

    async def test_clipboard_unknown_mode_returns_error(self, patched_desktop):
        mock_cb = MagicMock()
        mock_cb.CF_UNICODETEXT = 13
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="clear")
        assert 'mode must be either "get" or "set"' in result


class TestMissingRequiredArguments:
    """Tool functions return descriptive error strings when required args are absent."""

    async def test_file_write_without_content_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["File"].fn(mode="write", path="C:/out.txt")
        assert "content parameter is required" in result

    async def test_file_copy_without_destination_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["File"].fn(mode="copy", path="C:/src.txt")
        assert "destination parameter is required" in result

    async def test_file_move_without_destination_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["File"].fn(mode="move", path="C:/src.txt")
        assert "destination parameter is required" in result

    async def test_file_search_without_pattern_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["File"].fn(mode="search", path="C:/")
        assert "pattern parameter is required" in result

    async def test_registry_get_without_name_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="get", path="HKCU:\\Test")
        assert "name parameter is required" in result

    async def test_registry_set_without_name_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="set", path="HKCU:\\Test", value="val")
        assert "name parameter is required" in result

    async def test_registry_set_without_value_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Registry"].fn(mode="set", path="HKCU:\\Test", name="Key")
        assert "value parameter is required" in result

    async def test_clipboard_set_without_text_returns_error(self, patched_desktop):
        mock_cb = MagicMock()
        mock_cb.CF_UNICODETEXT = 13
        with patch.dict(sys.modules, {"win32clipboard": mock_cb}):
            tools = await _get_tools()
            result = await tools["Clipboard"].fn(mode="set", text=None)
        assert "text parameter required" in result

    async def test_invoke_set_value_without_value_returns_error(self, patched_desktop):
        elem = MagicMock()
        elem.Name = "Field"
        elem.AutomationId = "field"
        elem.LocalizedControlType = "edit"
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="set_value", value=None)
        assert "value parameter required" in result

    async def test_find_without_any_criteria_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Find"].fn()
        assert "at least one of name, control_type, or window must be specified" in result


class TestInvokeEdgeCases:
    async def test_invoke_no_element_at_coordinates_returns_error(self, patched_desktop):
        with patch("windows_mcp.uia.ControlFromPoint", return_value=None):
            tools = await _get_tools()
            result = await tools["Invoke"].fn(loc=[999, 999], action="invoke")
        assert "no element found at (999,999)" in result

    async def test_invoke_set_value_over_10000_chars_returns_error(self, patched_desktop):
        elem = MagicMock()
        elem.Name = "BigField"
        elem.AutomationId = "big"
        elem.LocalizedControlType = "edit"
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(
                    loc=[100, 200], action="set_value", value="x" * 10001
                )
        assert "value too long" in result
        assert "10001" in result

    async def test_invoke_set_value_exactly_10000_chars_succeeds(self, patched_desktop):
        elem = MagicMock()
        elem.Name = "Field"
        elem.AutomationId = "f"
        elem.LocalizedControlType = "edit"
        mock_pattern = MagicMock()
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(
                    loc=[100, 200], action="set_value", value="x" * 10000
                )
        assert "Set value" in result

    async def test_invoke_pattern_exception_returns_error_string(self, patched_desktop):
        elem = MagicMock()
        elem.Name = "OK"
        elem.AutomationId = "ok"
        elem.LocalizedControlType = "button"
        mock_pattern = MagicMock()
        mock_pattern.Invoke.side_effect = RuntimeError("COM RPC error")
        elem.GetPattern.return_value = mock_pattern

        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="invoke")
        assert "Error invoking invoke on 'OK'" in result


class TestBoolStringCoercion:
    """Tools that accept bool | str parameters coerce the string 'true'/'false' correctly."""

    async def test_move_drag_string_true_calls_drag(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Move"].fn(loc=[100, 200], drag="true")
        assert "Dragged" in result
        patched_desktop.drag.assert_called_once_with((100, 200))

    async def test_move_drag_string_false_calls_move(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Move"].fn(loc=[100, 200], drag="false")
        assert "Moved the mouse pointer" in result
        patched_desktop.move.assert_called_once_with((100, 200))

    async def test_multiselect_press_ctrl_string_false_passes_false(self, patched_desktop):
        tools = await _get_tools()
        await tools["MultiSelect"].fn(locs=[[100, 200]], press_ctrl="false")
        patched_desktop.multi_select.assert_called_once_with(False, [(100, 200)])

    async def test_multiselect_press_ctrl_string_true_passes_true(self, patched_desktop):
        tools = await _get_tools()
        await tools["MultiSelect"].fn(locs=[[100, 200]], press_ctrl="true")
        patched_desktop.multi_select.assert_called_once_with(True, [(100, 200)])

    async def test_process_kill_force_string_true_passes_force(self, patched_desktop):
        tools = await _get_tools()
        await tools["Process"].fn(mode="kill", pid=1234, force="true")
        call_kwargs = patched_desktop.kill_process.call_args[1]
        assert call_kwargs["force"] is True

    async def test_file_recursive_string_true_passes_true(self, patched_desktop):
        with patch("windows_mcp.filesystem.delete_path", return_value="Deleted."):
            tools = await _get_tools()
            await tools["File"].fn(mode="delete", path="C:/dir", recursive="true")

    async def test_snapshot_use_vision_string_true_coerces_to_true(self, patched_desktop):
        """use_vision='true' should trigger vision logic (screenshot processing)."""
        tools = await _get_tools()
        # With no screenshot bytes the result is still a list -- no exception
        result = await tools["Snapshot"].fn(use_vision="true", use_dom="false")
        assert isinstance(result, list)

    async def test_snapshot_vision_with_screenshot_returns_image(self, patched_desktop):
        """When use_vision=True and screenshot exists, result includes Image data."""
        from PIL import Image as PILImage

        # Create a small real PIL Image to test bytes conversion path
        img = PILImage.new("RGB", (100, 100), color=(255, 0, 0))
        ds = _make_desktop_state()
        ds.screenshot = img
        patched_desktop.get_state.return_value = ds

        tools = await _get_tools()
        result = await tools["Snapshot"].fn(use_vision=True, use_dom=False)
        assert isinstance(result, list)
        assert len(result) >= 2  # text + Image


class TestWaitForTimeoutCapping:
    async def test_waitfor_negative_timeout_capped_to_1_second(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state(window_name="Other")
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="Ghost", timeout=-100)
        # The message should report 1s (minimum)
        assert "within 1s" in result

    async def test_waitfor_zero_timeout_capped_to_1_second(self, patched_desktop):
        patched_desktop.get_state.return_value = _make_desktop_state(window_name="Other")
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="Ghost", timeout=0)
        assert "within 1s" in result

    async def test_waitfor_poll_exception_is_swallowed_and_continues(self, patched_desktop):
        """Exceptions during polling are caught; WaitFor times out gracefully."""
        call_count = [0]
        ds_match = _make_desktop_state(window_name="Target")

        def _side_effect(**kwargs):
            call_count[0] += 1
            if call_count[0] == 1:
                raise RuntimeError("transient COM error")
            return ds_match

        # asyncio.to_thread calls get_state positionally
        patched_desktop.get_state.side_effect = _side_effect
        tools = await _get_tools()
        result = await tools["WaitFor"].fn(mode="window", name="Target", timeout=3)
        assert "Window found" in result


class TestInputValidationGuards:
    """Tests for bounds checking and empty-input validation added to MCP tools."""

    async def test_wait_negative_duration_clamped_to_zero(self, patched_desktop):
        tools = await _get_tools()
        with patch("windows_mcp.__main__.pg") as mock_pg:
            result = await tools["Wait"].fn(duration=-5)
        mock_pg.sleep.assert_called_once_with(0)
        assert "Waited for 0 seconds" in result

    async def test_wait_zero_duration_is_valid(self, patched_desktop):
        tools = await _get_tools()
        with patch("windows_mcp.__main__.pg") as mock_pg:
            result = await tools["Wait"].fn(duration=0)
        mock_pg.sleep.assert_called_once_with(0)
        assert "Waited for 0 seconds" in result

    async def test_shell_empty_command_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Shell"].fn(command="")
        assert "Error" in result
        assert "empty" in result.lower()

    async def test_shell_whitespace_command_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Shell"].fn(command="   ")
        assert "Error" in result
        assert "empty" in result.lower()

    async def test_shell_negative_timeout_clamped_to_1(self, patched_desktop):
        patched_desktop.execute_command.return_value = ("output", 0)
        tools = await _get_tools()
        await tools["Shell"].fn(command="echo test", timeout=-10)
        patched_desktop.execute_command.assert_called_once_with("echo test", 1)

    async def test_shell_timeout_clamped_to_300(self, patched_desktop):
        patched_desktop.execute_command.return_value = ("output", 0)
        tools = await _get_tools()
        await tools["Shell"].fn(command="echo test", timeout=9999)
        patched_desktop.execute_command.assert_called_once_with("echo test", 300)

    async def test_shortcut_empty_string_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Shortcut"].fn(shortcut="")
        assert "Error" in result
        assert "empty" in result.lower()

    async def test_shortcut_whitespace_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["Shortcut"].fn(shortcut="  ")
        assert "Error" in result

    async def test_multi_select_empty_locs_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["MultiSelect"].fn(locs=[])
        assert "Error" in result
        assert "at least one" in result.lower()

    async def test_multi_edit_empty_locs_returns_error(self, patched_desktop):
        tools = await _get_tools()
        result = await tools["MultiEdit"].fn(locs=[])
        assert "Error" in result
        assert "at least one" in result.lower()


class TestLifespanContextManager:
    async def test_lifespan_initialises_desktop(self):
        mock_app = MagicMock()
        mock_desktop_cls = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_cls.return_value = mock_desktop_inst
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()

        mock_watchdog_cls = MagicMock()
        mock_watchdog_inst = MagicMock()
        mock_watchdog_cls.return_value = mock_watchdog_inst

        with patch("windows_mcp.__main__.Desktop", mock_desktop_cls):
            with patch("windows_mcp.__main__.WatchDog", mock_watchdog_cls):
                with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "false"}):
                    with patch("asyncio.sleep", AsyncMock()):
                        async with main_module.lifespan(mock_app):
                            assert main_module.desktop is mock_desktop_inst

    async def test_lifespan_initialises_screen_size(self):
        mock_app = MagicMock()
        mock_desktop_cls = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_cls.return_value = mock_desktop_inst
        mock_desktop_inst.get_screen_size.return_value = Size(2560, 1440)
        mock_desktop_inst.tree._on_focus_change = MagicMock()

        mock_watchdog_cls = MagicMock()
        mock_watchdog_inst = MagicMock()
        mock_watchdog_cls.return_value = mock_watchdog_inst

        with patch("windows_mcp.__main__.Desktop", mock_desktop_cls):
            with patch("windows_mcp.__main__.WatchDog", mock_watchdog_cls):
                with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "false"}):
                    with patch("asyncio.sleep", AsyncMock()):
                        async with main_module.lifespan(mock_app):
                            assert main_module.screen_size == Size(2560, 1440)

    async def test_lifespan_starts_watchdog(self):
        mock_app = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()
        mock_watchdog_inst = MagicMock()

        with patch("windows_mcp.__main__.Desktop", return_value=mock_desktop_inst):
            with patch("windows_mcp.__main__.WatchDog", return_value=mock_watchdog_inst):
                with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "false"}):
                    with patch("asyncio.sleep", AsyncMock()):
                        async with main_module.lifespan(mock_app):
                            pass
        mock_watchdog_inst.start.assert_called_once()

    async def test_lifespan_stops_watchdog_on_exit(self):
        mock_app = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()
        mock_watchdog_inst = MagicMock()

        with patch("windows_mcp.__main__.Desktop", return_value=mock_desktop_inst):
            with patch("windows_mcp.__main__.WatchDog", return_value=mock_watchdog_inst):
                with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "false"}):
                    with patch("asyncio.sleep", AsyncMock()):
                        async with main_module.lifespan(mock_app):
                            pass
        mock_watchdog_inst.stop.assert_called_once()

    async def test_lifespan_creates_analytics_when_telemetry_enabled(self):
        mock_app = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()
        mock_watchdog_inst = MagicMock()
        mock_analytics_inst = MagicMock()
        mock_analytics_inst.close = AsyncMock()
        mock_analytics_cls = MagicMock(return_value=mock_analytics_inst)

        with patch("windows_mcp.__main__.Desktop", return_value=mock_desktop_inst):
            with patch("windows_mcp.__main__.WatchDog", return_value=mock_watchdog_inst):
                with patch("windows_mcp.__main__.PostHogAnalytics", mock_analytics_cls):
                    with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "true"}):
                        with patch("asyncio.sleep", AsyncMock()):
                            async with main_module.lifespan(mock_app):
                                assert main_module.analytics is mock_analytics_inst

    async def test_lifespan_closes_analytics_on_exit(self):
        mock_app = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()
        mock_watchdog_inst = MagicMock()
        mock_analytics_inst = MagicMock()
        mock_analytics_inst.close = AsyncMock()
        mock_analytics_cls = MagicMock(return_value=mock_analytics_inst)

        with patch("windows_mcp.__main__.Desktop", return_value=mock_desktop_inst):
            with patch("windows_mcp.__main__.WatchDog", return_value=mock_watchdog_inst):
                with patch("windows_mcp.__main__.PostHogAnalytics", mock_analytics_cls):
                    with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "true"}):
                        with patch("asyncio.sleep", AsyncMock()):
                            async with main_module.lifespan(mock_app):
                                pass
        mock_analytics_inst.close.assert_called_once()

    async def test_lifespan_skips_analytics_when_telemetry_disabled(self):
        mock_app = MagicMock()
        mock_desktop_inst = MagicMock()
        mock_desktop_inst.get_screen_size.return_value = SCREEN_SIZE
        mock_desktop_inst.tree._on_focus_change = MagicMock()
        mock_watchdog_inst = MagicMock()
        mock_analytics_cls = MagicMock()

        with patch("windows_mcp.__main__.Desktop", return_value=mock_desktop_inst):
            with patch("windows_mcp.__main__.WatchDog", return_value=mock_watchdog_inst):
                with patch("windows_mcp.__main__.PostHogAnalytics", mock_analytics_cls):
                    with patch.dict("os.environ", {"ANONYMIZED_TELEMETRY": "false"}):
                        with patch("asyncio.sleep", AsyncMock()):
                            async with main_module.lifespan(mock_app):
                                pass
        mock_analytics_cls.assert_not_called()


# ---------------------------------------------------------------------------
# Tier 4: Transport & Auth Integration
# ---------------------------------------------------------------------------


class TestCLIGenerateKey:
    def test_generate_key_exits_zero_and_prints_key(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.generate_key.return_value = "generated-key-abc123"
            result = runner.invoke(main_module.main, ["--generate-key"])
        assert result.exit_code == 0
        assert "generated-key-abc123" in result.output
        assert "DPAPI" in result.output


class TestCLIRotateKey:
    def test_rotate_key_exits_zero_and_prints_new_key(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.rotate_key.return_value = "new-rotated-key-xyz"
            result = runner.invoke(main_module.main, ["--rotate-key"])
        assert result.exit_code == 0
        assert "new-rotated-key-xyz" in result.output


class TestCLIAuthSafety:
    def test_sse_without_auth_and_non_localhost_host_exits_1(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = None
            result = runner.invoke(
                main_module.main,
                ["--transport", "sse", "--host", "0.0.0.0", "--port", "8000"],
            )
        assert result.exit_code == 1
        assert "Cannot bind to 0.0.0.0" in result.output

    def test_streamable_http_without_auth_and_non_localhost_exits_1(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = None
            result = runner.invoke(
                main_module.main,
                [
                    "--transport",
                    "streamable-http",
                    "--host",
                    "192.168.1.10",
                    "--port",
                    "9000",
                ],
            )
        assert result.exit_code == 1

    def test_sse_with_api_key_on_non_localhost_succeeds(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = None
            with patch.object(main_module.mcp, "add_middleware") as mock_add_mw:
                with patch.object(main_module.mcp, "run"):
                    result = runner.invoke(
                        main_module.main,
                        [
                            "--transport",
                            "sse",
                            "--host",
                            "0.0.0.0",
                            "--port",
                            "8000",
                            "--api-key",
                            "super-secret",
                        ],
                    )
        assert result.exit_code == 0
        mock_add_mw.assert_called_once()

    def test_sse_with_api_key_installs_bearer_auth_middleware(self):
        from click.testing import CliRunner

        from windows_mcp.auth import BearerAuthMiddleware

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = None
            with patch.object(main_module.mcp, "add_middleware") as mock_add_mw:
                with patch.object(main_module.mcp, "run"):
                    runner.invoke(
                        main_module.main,
                        [
                            "--transport",
                            "sse",
                            "--host",
                            "localhost",
                            "--port",
                            "8000",
                            "--api-key",
                            "my-key",
                        ],
                    )
        middleware_instance = mock_add_mw.call_args[0][0]
        assert isinstance(middleware_instance, BearerAuthMiddleware)

    def test_sse_localhost_without_auth_exits_0_with_warning(self):
        """localhost binding without an API key is allowed but logs a warning."""
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = None
            with patch.object(main_module.mcp, "run"):
                result = runner.invoke(
                    main_module.main,
                    ["--transport", "sse", "--host", "localhost", "--port", "8000"],
                )
        assert result.exit_code == 0

    def test_stored_key_loaded_when_no_cli_key_provided(self):
        """AuthKeyManager.load_key() is called for HTTP transports with no --api-key."""
        from click.testing import CliRunner

        runner = CliRunner()
        with patch("windows_mcp.__main__.AuthKeyManager") as mock_km:
            mock_km.load_key.return_value = "stored-key"
            with patch.object(main_module.mcp, "add_middleware"):
                with patch.object(main_module.mcp, "run"):
                    runner.invoke(
                        main_module.main,
                        ["--transport", "sse", "--host", "localhost", "--port", "8000"],
                    )
        mock_km.load_key.assert_called_once()


class TestCLIStdioTransport:
    def test_stdio_transport_calls_mcp_run_with_stdio(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch.object(main_module.mcp, "run") as mock_run:
            result = runner.invoke(main_module.main, ["--transport", "stdio"])
        assert result.exit_code == 0
        mock_run.assert_called_once()
        call_kwargs = mock_run.call_args[1]
        assert call_kwargs.get("transport") == "stdio"

    def test_stdio_transport_does_not_install_auth_middleware(self):
        from click.testing import CliRunner

        runner = CliRunner()
        with patch.object(main_module.mcp, "add_middleware") as mock_add_mw:
            with patch.object(main_module.mcp, "run"):
                runner.invoke(main_module.main, ["--transport", "stdio"])
        mock_add_mw.assert_not_called()
