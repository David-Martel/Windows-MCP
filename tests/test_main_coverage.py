"""Branch-coverage tests for windows_mcp.__main__ gaps.

Organised by the cluster of source branches each group targets:

  Cluster 1  - file_tool relative-path resolution and exception handling
  Cluster 2  - state_tool (Snapshot) null screen_size and error branch
  Cluster 3  - scroll_tool loc=None message branch and string return path
  Cluster 4  - multi_select_tool per-entry validation
  Cluster 5  - multi_edit_tool per-entry validation
  Cluster 6  - find_tool / waitfor_tool null tree_state branches
  Cluster 7  - invoke_tool: every action's "pattern not supported" branch,
               length guard, no-element, and invalid loc
  Cluster 8  - Transport / Mode enum __str__ values
  Cluster 9  - scrape_tool DOM mode with null tree_state

Design follows the patterns established in test_mcp_integration.py:
  - patch.object(main_module, "desktop", mock) + "screen_size" as required
  - tools retrieved via await main_module.mcp._tool_manager.get_tools()
  - tool.fn(...) called directly to bypass FastMCP context injection
  - asyncio_mode = "auto" in pyproject.toml so bare async def test_* works
"""

import os
from unittest.mock import MagicMock, patch

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
# _coerce_bool unit tests
# ---------------------------------------------------------------------------


class TestCoerceBool:
    """Tests for _coerce_bool() utility that normalises bool|str MCP params."""

    def test_true_bool_returns_true(self):
        assert main_module._coerce_bool(True) is True

    def test_false_bool_returns_false(self):
        assert main_module._coerce_bool(False) is False

    def test_string_true_returns_true(self):
        assert main_module._coerce_bool("true") is True

    def test_string_false_returns_false(self):
        assert main_module._coerce_bool("false") is False

    def test_string_TRUE_case_insensitive(self):
        assert main_module._coerce_bool("TRUE") is True

    def test_string_True_mixed_case(self):
        assert main_module._coerce_bool("True") is True

    def test_string_FALSE_case_insensitive(self):
        assert main_module._coerce_bool("FALSE") is False

    def test_empty_string_returns_false(self):
        assert main_module._coerce_bool("") is False

    def test_arbitrary_string_returns_false(self):
        assert main_module._coerce_bool("yes") is False

    def test_none_returns_default_false(self):
        assert main_module._coerce_bool(None) is False

    def test_none_returns_custom_default_true(self):
        assert main_module._coerce_bool(None, default=True) is True

    def test_int_returns_default(self):
        assert main_module._coerce_bool(1) is False

    def test_int_returns_custom_default(self):
        assert main_module._coerce_bool(1, default=True) is True


# ---------------------------------------------------------------------------
# Shared helpers (local copies so this file is self-contained)
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
    tree_state: TreeState | None = None,
) -> DesktopState:
    """Build a minimal DesktopState.  Pass tree_state=None to leave it absent."""
    win = _make_window(window_name)
    if tree_state is None and tree_nodes is not None:
        tree_state = TreeState(interactive_nodes=tree_nodes)
    elif tree_state is None:
        tree_state = TreeState(interactive_nodes=[])
    return DesktopState(
        active_desktop={"name": "Desktop 1", "id": "abc-123"},
        all_desktops=[{"name": "Desktop 1", "id": "abc-123"}],
        active_window=win,
        windows=[win],
        screenshot=None,
        tree_state=tree_state,
    )


def _make_desktop_state_no_tree(window_name: str = "Notepad") -> DesktopState:
    """Build a DesktopState with tree_state explicitly set to None."""
    win = _make_window(window_name)
    return DesktopState(
        active_desktop={"name": "Desktop 1", "id": "abc-123"},
        all_desktops=[{"name": "Desktop 1", "id": "abc-123"}],
        active_window=win,
        windows=[win],
        screenshot=None,
        tree_state=None,
    )


@pytest.fixture
def mock_desktop() -> MagicMock:
    m = MagicMock()
    m.get_screen_size.return_value = SCREEN_SIZE
    m.get_state.return_value = _make_desktop_state()
    m.scroll.return_value = None
    m.multi_select.return_value = None
    m.multi_edit.return_value = None
    m.scrape.return_value = "# Hello"
    return m


@pytest.fixture
def patched_desktop(mock_desktop):
    with patch.object(main_module, "desktop", mock_desktop):
        with patch.object(main_module, "screen_size", SCREEN_SIZE):
            yield mock_desktop


async def _get_tools() -> dict:
    return await main_module.mcp._tool_manager.get_tools()


# ---------------------------------------------------------------------------
# Cluster 1 - file_tool relative path resolution and exception branch
# ---------------------------------------------------------------------------


class TestFileToolRelativePaths:
    async def test_file_read_relative_path_resolves_to_desktop_dir(self, patched_desktop):
        """A relative path is joined onto the user_desktop_dir result."""
        with patch("platformdirs.user_desktop_dir", return_value="C:/Users/test/Desktop"):
            with patch("windows_mcp.filesystem.read_file", return_value="content") as mock_read:
                tools = await _get_tools()
                result = await tools["File"].fn(mode="read", path="notes.txt")
        assert "content" in result
        # The path argument passed to read_file must be the joined absolute path
        call_args = mock_read.call_args[0]
        passed_path = call_args[0]
        assert passed_path == os.path.join("C:/Users/test/Desktop", "notes.txt")

    async def test_file_copy_relative_destination_resolves(self, patched_desktop):
        """A relative destination is also joined onto the user_desktop_dir result."""
        with patch("platformdirs.user_desktop_dir", return_value="C:/Users/test/Desktop"):
            with patch("windows_mcp.filesystem.copy_path", return_value="Copied.") as mock_copy:
                tools = await _get_tools()
                await tools["File"].fn(
                    mode="copy",
                    path="C:/src.txt",
                    destination="backup.txt",
                )
        call_kwargs = mock_copy.call_args
        dest_arg = call_kwargs[0][1]
        assert dest_arg == os.path.join("C:/Users/test/Desktop", "backup.txt")

    async def test_file_tool_exception_returns_error_string(self, patched_desktop):
        """Any exception raised inside file_tool is caught and returned as a string."""
        with patch("platformdirs.user_desktop_dir", side_effect=OSError("no desktop")):
            tools = await _get_tools()
            result = await tools["File"].fn(mode="read", path="anything.txt")
        assert "Error in File tool" in result
        assert "no desktop" in result


# ---------------------------------------------------------------------------
# Cluster 2 - state_tool (Snapshot) null screen_size and exception path
# ---------------------------------------------------------------------------


class TestSnapshotNullScreenSize:
    async def test_snapshot_uses_scale_1_when_screen_size_none(self, mock_desktop):
        """When screen_size is None the scale factor must default to 1.0."""
        mock_desktop.get_state.return_value = _make_desktop_state()
        with patch.object(main_module, "desktop", mock_desktop):
            with patch.object(main_module, "screen_size", None):
                tools = await _get_tools()
                result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)

        assert isinstance(result, list)
        # Verify get_state was called with scale=1.0 (not some computed fraction)
        call_kwargs = mock_desktop.get_state.call_args[1]
        assert call_kwargs["scale"] == 1.0

    async def test_snapshot_error_returns_error_list(self, mock_desktop):
        """An exception from desktop.get_state returns a single-element list with the error."""
        mock_desktop.get_state.side_effect = RuntimeError("UIA crash")
        with patch.object(main_module, "desktop", mock_desktop):
            with patch.object(main_module, "screen_size", SCREEN_SIZE):
                tools = await _get_tools()
                result = await tools["Snapshot"].fn(use_vision=False, use_dom=False)

        assert isinstance(result, list)
        assert len(result) == 1
        assert "Error capturing desktop state" in result[0]
        assert "UIA crash" in result[0]


# ---------------------------------------------------------------------------
# Cluster 3 - scroll_tool loc=None branch and desktop string passthrough
# ---------------------------------------------------------------------------


class TestScrollToolNullLoc:
    async def test_scroll_without_loc_message_has_no_coordinates(self, patched_desktop):
        """When loc is None the result must not contain 'at (' coordinate text."""
        patched_desktop.scroll.return_value = None
        tools = await _get_tools()
        result = await tools["Scroll"].fn(loc=None, direction="down")
        assert "Scrolled" in result
        assert "at (" not in result

    async def test_scroll_with_desktop_response_returns_it(self, patched_desktop):
        """A non-None return value from desktop.scroll is passed through as-is."""
        patched_desktop.scroll.return_value = "Reached bottom of list."
        tools = await _get_tools()
        result = await tools["Scroll"].fn(loc=[100, 200], direction="down")
        assert result == "Reached bottom of list."

    async def test_scroll_without_loc_passes_none_to_desktop(self, patched_desktop):
        """loc=None is forwarded unchanged to desktop.scroll."""
        patched_desktop.scroll.return_value = None
        tools = await _get_tools()
        await tools["Scroll"].fn(loc=None, direction="up", wheel_times=2)
        patched_desktop.scroll.assert_called_once_with(None, "vertical", "up", 2)


# ---------------------------------------------------------------------------
# Cluster 4 - multi_select_tool per-entry validation
# ---------------------------------------------------------------------------


class TestMultiSelectValidation:
    async def test_multiselect_invalid_loc_entry_raises(self, patched_desktop):
        """An entry with only one element (not two) raises ValueError."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[0\] must be \[x, y\]"):
            await tools["MultiSelect"].fn(locs=[[100]])

    async def test_multiselect_non_list_entry_raises(self, patched_desktop):
        """A string entry in the locs list raises ValueError."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[0\] must be \[x, y\]"):
            await tools["MultiSelect"].fn(locs=["100,200"])

    async def test_multiselect_second_invalid_entry_raises_with_correct_index(
        self, patched_desktop
    ):
        """Validation error message includes the correct index of the bad entry."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[1\] must be \[x, y\]"):
            await tools["MultiSelect"].fn(locs=[[100, 200], [300]])


# ---------------------------------------------------------------------------
# Cluster 5 - multi_edit_tool per-entry validation
# ---------------------------------------------------------------------------


class TestMultiEditValidation:
    async def test_multiedit_short_entry_raises(self, patched_desktop):
        """An entry with only [x, y] (missing text) raises ValueError."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[0\] must be \[x, y, text\]"):
            await tools["MultiEdit"].fn(locs=[[100, 200]])

    async def test_multiedit_non_list_entry_raises(self, patched_desktop):
        """A string entry raises ValueError."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[0\] must be \[x, y, text\]"):
            await tools["MultiEdit"].fn(locs=["100,200,hello"])

    async def test_multiedit_second_invalid_entry_raises_with_correct_index(self, patched_desktop):
        """Validation error message includes the correct index of the bad entry."""
        tools = await _get_tools()
        with pytest.raises(ValueError, match=r"locs\[1\] must be \[x, y, text\]"):
            await tools["MultiEdit"].fn(locs=[[100, 200, "ok"], [300, 400]])


# ---------------------------------------------------------------------------
# Cluster 6 - find_tool and waitfor_tool null tree_state branches
# ---------------------------------------------------------------------------


class TestFindToolNullTreeState:
    async def test_find_null_tree_state_returns_error(self, patched_desktop):
        """When desktop_state.tree_state is None find_tool returns an error string."""
        patched_desktop.get_state.return_value = _make_desktop_state_no_tree()
        tools = await _get_tools()
        result = await tools["Find"].fn(name="OK")
        assert "Error" in result
        assert "could not capture desktop state" in result

    async def test_find_no_matches_returns_criteria_in_message(self, patched_desktop):
        """When there are zero matches the error message includes the search criteria."""
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=[])
        tools = await _get_tools()
        result = await tools["Find"].fn(name="GhostButton", control_type="Button")
        assert "No elements found" in result
        assert "GhostButton" in result
        assert "Button" in result

    async def test_find_no_matches_with_window_criteria_shows_window(self, patched_desktop):
        """When window criterion is given it appears in the no-match message."""
        patched_desktop.get_state.return_value = _make_desktop_state(tree_nodes=[])
        tools = await _get_tools()
        result = await tools["Find"].fn(window="Calculator")
        assert "Calculator" in result


class TestWaitForNullTreeState:
    async def test_waitfor_element_null_tree_state_times_out(self, mock_desktop):
        """When tree_state is None the element branch never matches and WaitFor times out."""
        mock_desktop.get_state.return_value = _make_desktop_state_no_tree()
        with patch.object(main_module, "desktop", mock_desktop):
            with patch.object(main_module, "screen_size", SCREEN_SIZE):
                tools = await _get_tools()
                result = await tools["WaitFor"].fn(mode="element", name="SomeButton", timeout=1)
        assert "Timeout" in result
        assert "SomeButton" in result

    async def test_waitfor_element_skips_null_tree_state_then_finds_element(self, mock_desktop):
        """A first poll returning null tree_state is skipped; the second poll succeeds."""
        node = _make_tree_element(name="Submit Button", control_type="Button")
        ds_with_tree = _make_desktop_state(tree_nodes=[node])
        ds_no_tree = _make_desktop_state_no_tree()

        call_count = [0]

        def _side_effect(**_kwargs):
            call_count[0] += 1
            return ds_no_tree if call_count[0] == 1 else ds_with_tree

        mock_desktop.get_state.side_effect = _side_effect
        with patch.object(main_module, "desktop", mock_desktop):
            with patch.object(main_module, "screen_size", SCREEN_SIZE):
                tools = await _get_tools()
                result = await tools["WaitFor"].fn(mode="element", name="Submit", timeout=5)
        assert "Element found" in result
        assert "Submit Button" in result


# ---------------------------------------------------------------------------
# Cluster 7 - invoke_tool: "pattern not supported" for each action variant,
#              value-length guard, no-element, and invalid loc
# ---------------------------------------------------------------------------


class TestInvokeToolPatternVariants:
    """Each action has its own 'pattern not supported' branch that must be hit."""

    def _make_elem(self, name: str = "Widget", control_type: str = "button") -> MagicMock:
        elem = MagicMock()
        elem.Name = name
        elem.AutomationId = "w_id"
        elem.LocalizedControlType = control_type
        elem.GetPattern.return_value = None  # simulate unsupported pattern
        return elem

    async def test_invoke_toggle_no_pattern_returns_error(self, patched_desktop):
        elem = self._make_elem(name="MyCheckbox", control_type="checkbox")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="toggle")
        assert "does not support TogglePattern" in result
        assert "MyCheckbox" in result

    async def test_invoke_expand_no_pattern_returns_error(self, patched_desktop):
        elem = self._make_elem(name="MyCombo", control_type="combobox")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="expand")
        assert "does not support ExpandCollapsePattern" in result
        assert "MyCombo" in result

    async def test_invoke_collapse_no_pattern_returns_error(self, patched_desktop):
        elem = self._make_elem(name="MyTree", control_type="treeitem")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="collapse")
        assert "does not support ExpandCollapsePattern" in result
        assert "MyTree" in result

    async def test_invoke_select_no_pattern_returns_error(self, patched_desktop):
        elem = self._make_elem(name="ListItem", control_type="listitem")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="select")
        assert "does not support SelectionItemPattern" in result
        assert "ListItem" in result

    async def test_invoke_set_value_no_pattern_returns_error(self, patched_desktop):
        elem = self._make_elem(name="ReadOnlyField", control_type="text")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="set_value", value="hello")
        assert "does not support ValuePattern" in result
        assert "ReadOnlyField" in result

    async def test_invoke_set_value_too_long_returns_error(self, patched_desktop):
        """Values exceeding 10 000 characters are rejected before any UIA call."""
        elem = self._make_elem(name="BigInput", control_type="edit")
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(
                    loc=[100, 200], action="set_value", value="x" * 10001
                )
        assert "value too long" in result
        assert "10001" in result

    async def test_invoke_no_element_at_coords_returns_error(self, patched_desktop):
        """ControlFromPoint returning None triggers the 'no element' error path."""
        with patch("windows_mcp.uia.ControlFromPoint", return_value=None):
            tools = await _get_tools()
            result = await tools["Invoke"].fn(loc=[1, 1], action="invoke")
        assert "no element found at (1,1)" in result

    async def test_invoke_invalid_loc_returns_error(self, patched_desktop):
        """A loc with length != 2 returns the error string (not a ValueError)."""
        tools = await _get_tools()
        result = await tools["Invoke"].fn(loc=[100], action="invoke")
        assert "Error: loc must be [x, y]" in result

    async def test_invoke_unknown_action_returns_error(self, patched_desktop):
        """An action string not in the known set returns the unknown-action error."""
        elem = self._make_elem()
        elem.GetPattern.return_value = MagicMock()  # pattern is present
        with patch("windows_mcp.uia.ControlFromPoint", return_value=elem):
            with patch("windows_mcp.uia.PatternId"):
                tools = await _get_tools()
                result = await tools["Invoke"].fn(loc=[100, 200], action="fly")
        assert "unknown action" in result
        assert "fly" in result


# ---------------------------------------------------------------------------
# Cluster 8 - Transport and Mode enum __str__ values
# ---------------------------------------------------------------------------


class TestTransportEnum:
    def test_transport_stdio_str(self):
        assert str(main_module.Transport.STDIO) == "stdio"

    def test_transport_sse_str(self):
        assert str(main_module.Transport.SSE) == "sse"

    def test_transport_streamable_http_str(self):
        assert str(main_module.Transport.STREAMABLE_HTTP) == "streamable-http"

    def test_transport_values_are_correct(self):
        assert main_module.Transport.STDIO.value == "stdio"
        assert main_module.Transport.SSE.value == "sse"
        assert main_module.Transport.STREAMABLE_HTTP.value == "streamable-http"


class TestModeEnum:
    def test_mode_local_str(self):
        assert str(main_module.Mode.LOCAL) == "local"

    def test_mode_remote_str(self):
        assert str(main_module.Mode.REMOTE) == "remote"

    def test_mode_values_are_correct(self):
        assert main_module.Mode.LOCAL.value == "local"
        assert main_module.Mode.REMOTE.value == "remote"


# ---------------------------------------------------------------------------
# Cluster 9 - scrape_tool DOM mode with null tree_state
# ---------------------------------------------------------------------------


class TestScrapeToolNullTreeState:
    async def test_scrape_dom_mode_null_tree_state_returns_error(self, patched_desktop):
        """use_dom=True when tree_state is None returns the 'No DOM information' message."""
        patched_desktop.get_state.return_value = _make_desktop_state_no_tree()
        tools = await _get_tools()
        result = await tools["Scrape"].fn(url="https://example.com", use_dom=True)
        assert "No DOM information found" in result
        assert "https://example.com" in result

    async def test_scrape_dom_mode_tree_state_without_dom_node_returns_error(self, patched_desktop):
        """use_dom=True when tree_state.dom_node is None also returns the 'No DOM' message."""
        ts = TreeState(dom_node=None, interactive_nodes=[])
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
        assert "No DOM information found" in result

    async def test_scrape_http_mode_ignores_tree_state(self, patched_desktop):
        """When use_dom=False desktop.scrape is called regardless of tree_state."""
        patched_desktop.scrape.return_value = "# Page"
        patched_desktop.get_state.return_value = _make_desktop_state_no_tree()
        tools = await _get_tools()
        result = await tools["Scrape"].fn(url="https://example.com", use_dom=False)
        assert "https://example.com" in result
        patched_desktop.scrape.assert_called_once_with("https://example.com")


# ---------------------------------------------------------------------------
# Cluster 10 - main() mode/transport validation paths
# ---------------------------------------------------------------------------


class TestMainFunctionValidation:
    """Cover lines 1088-1089, 1104-1105, 1106-1107 in main()."""

    def test_invalid_transport_local_mode(self):
        """Invalid transport in local mode raises ValueError (lines 1088-1089)."""
        with patch.dict(os.environ, {"MODE": "local"}, clear=False):
            with patch.object(main_module, "mcp"):
                # Simulate calling main with invalid transport
                with pytest.raises(ValueError, match="Invalid transport"):
                    main_module.main.callback(
                        transport="invalid-transport",
                        host="localhost",
                        port=8000,
                        api_key=None,
                        generate_key=False,
                        rotate_key=False,
                    )

    def test_invalid_mode_raises(self):
        """Invalid mode raises ValueError (lines 1106-1107)."""
        with patch.dict(os.environ, {"MODE": "bogus-mode"}, clear=False):
            with pytest.raises(ValueError, match="Invalid mode"):
                main_module.main.callback(
                    transport="stdio",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=False,
                )

    def test_remote_mode_missing_sandbox_id(self):
        """Remote mode without SANDBOX_ID raises ValueError (line 1092)."""
        with patch.dict(
            os.environ, {"MODE": "remote", "SANDBOX_ID": "", "API_KEY": "k"}, clear=False
        ):
            with pytest.raises(ValueError, match="SANDBOX_ID is required"):
                main_module.main.callback(
                    transport="stdio",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=False,
                )

    def test_remote_mode_missing_api_key(self):
        """Remote mode without API_KEY raises ValueError (lines 1093-1094)."""
        with patch.dict(
            os.environ, {"MODE": "remote", "SANDBOX_ID": "sb123", "API_KEY": ""}, clear=False
        ):
            with pytest.raises(ValueError, match="API_KEY is required"):
                main_module.main.callback(
                    transport="stdio",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=False,
                )

    def test_generate_key_exits(self):
        """--generate-key flag generates key and exits (lines 1038-1043)."""
        with patch.object(main_module.AuthKeyManager, "generate_key", return_value="test-key-123"):
            with pytest.raises(SystemExit) as exc_info:
                main_module.main.callback(
                    transport="stdio",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=True,
                    rotate_key=False,
                )
            assert exc_info.value.code == 0

    def test_rotate_key_exits(self):
        """--rotate-key flag rotates key and exits (lines 1045-1049)."""
        with patch.object(main_module.AuthKeyManager, "rotate_key", return_value="new-key-456"):
            with pytest.raises(SystemExit) as exc_info:
                main_module.main.callback(
                    transport="stdio",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=True,
                )
            assert exc_info.value.code == 0

    def test_sse_without_auth_non_localhost_exits(self):
        """SSE on non-localhost without auth exits with code 1 (lines 1062-1073)."""
        with patch.object(main_module.AuthKeyManager, "load_key", return_value=None):
            with pytest.raises(SystemExit) as exc_info:
                main_module.main.callback(
                    transport="sse",
                    host="0.0.0.0",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=False,
                )
            assert exc_info.value.code == 1

    def test_sse_localhost_no_auth_warns_but_runs(self):
        """SSE on localhost without auth logs warning but starts server (line 1074)."""
        with patch.object(main_module.AuthKeyManager, "load_key", return_value=None):
            with patch.object(main_module, "mcp") as mock_mcp:
                main_module.main.callback(
                    transport="sse",
                    host="localhost",
                    port=8000,
                    api_key=None,
                    generate_key=False,
                    rotate_key=False,
                )
                mock_mcp.run.assert_called_once()

    def test_sse_with_api_key_adds_middleware(self):
        """SSE with --api-key adds BearerAuthMiddleware (lines 1057-1059)."""
        with patch.object(main_module, "mcp") as mock_mcp:
            main_module.main.callback(
                transport="sse",
                host="localhost",
                port=8000,
                api_key="my-secret-key",
                generate_key=False,
                rotate_key=False,
            )
            mock_mcp.add_middleware.assert_called_once()
            mock_mcp.run.assert_called_once()
