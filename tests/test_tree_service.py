from time import time
from types import SimpleNamespace
from unittest.mock import MagicMock, create_autospec, patch

import pytest

from windows_mcp.desktop.views import Size
from windows_mcp.tree.service import Tree
from windows_mcp.tree.views import BoundingBox, TreeElementNode
from windows_mcp.uia.controls import WindowControl


@pytest.fixture
def tree_instance():
    mock_desktop = MagicMock()
    mock_desktop.get_screen_size.return_value = Size(width=1920, height=1080)
    return Tree(mock_desktop)


@pytest.fixture
def tree_with_desktop():
    """Return (tree, mock_desktop) so tests can keep the desktop alive through the proxy."""
    mock_desktop = MagicMock()
    mock_desktop.get_screen_size.return_value = Size(width=1920, height=1080)
    tree = Tree(mock_desktop)
    return tree, mock_desktop


class TestIouBoundingBox:
    def test_full_overlap(self, tree_instance):
        window = SimpleNamespace(left=0, top=0, right=500, bottom=500)
        element = SimpleNamespace(left=100, top=100, right=200, bottom=200)
        result = tree_instance.iou_bounding_box(window, element)
        assert result.left == 100
        assert result.top == 100
        assert result.right == 200
        assert result.bottom == 200
        assert result.width == 100
        assert result.height == 100

    def test_partial_overlap(self, tree_instance):
        window = SimpleNamespace(left=0, top=0, right=150, bottom=150)
        element = SimpleNamespace(left=100, top=100, right=200, bottom=200)
        result = tree_instance.iou_bounding_box(window, element)
        assert result.left == 100
        assert result.top == 100
        assert result.right == 150
        assert result.bottom == 150
        assert result.width == 50
        assert result.height == 50

    def test_no_overlap(self, tree_instance):
        window = SimpleNamespace(left=0, top=0, right=50, bottom=50)
        element = SimpleNamespace(left=100, top=100, right=200, bottom=200)
        result = tree_instance.iou_bounding_box(window, element)
        assert result.width == 0
        assert result.height == 0

    def test_screen_clamping(self, tree_instance):
        # Element extends beyond screen (1920x1080)
        window = SimpleNamespace(left=0, top=0, right=2000, bottom=2000)
        element = SimpleNamespace(left=1900, top=1060, right=2000, bottom=1200)
        result = tree_instance.iou_bounding_box(window, element)
        assert result.left == 1900
        assert result.top == 1060
        assert result.right == 1920
        assert result.bottom == 1080
        assert result.width == 20
        assert result.height == 20


# ---------------------------------------------------------------------------
# Helpers for tree_traversal tests
# ---------------------------------------------------------------------------


def _make_rect(left=0, top=0, right=200, bottom=200):
    """Build a SimpleNamespace that behaves like a UIA BoundingRectangle."""
    r = SimpleNamespace(left=left, top=top, right=right, bottom=bottom)
    r.width = lambda: right - left
    r.height = lambda: bottom - top
    return r


def _make_uia_node(
    control_type_name="ButtonControl",
    localized_control_type="button",
    name="OK",
    is_offscreen=False,
    is_enabled=True,
    is_control_element=True,
    is_keyboard_focusable=True,
    has_keyboard_focus=False,
    accelerator_key="",
    automation_id="",
    bounding_rect=None,
    children=None,
    legacy_role_name="PushButton",
    legacy_value="",
    legacy_default_action="",
    window_pattern_is_modal=False,
    scroll_pattern=None,
    automation_element_id="RootWebArea_NO",
):
    """
    Build a MagicMock that looks like a comtypes Control with cached properties.

    The mock already has _is_cached set so tree_traversal skips BuildUpdatedCache,
    and GetCachedChildren returns the children list directly.
    """
    rect = bounding_rect or _make_rect(10, 10, 200, 200)
    node = MagicMock()

    # Mark as already cached so the traversal skips BuildUpdatedCache
    node._is_cached = True

    # Cached property attributes
    node.CachedControlTypeName = control_type_name
    node.CachedLocalizedControlType = localized_control_type
    node.CachedName = name
    node.CachedIsOffscreen = is_offscreen
    node.CachedIsEnabled = is_enabled
    node.CachedIsControlElement = is_control_element
    node.CachedIsKeyboardFocusable = is_keyboard_focusable
    node.CachedHasKeyboardFocus = has_keyboard_focus
    node.CachedAcceleratorKey = accelerator_key
    node.CachedAutomationId = automation_id
    node.CachedBoundingRectangle = rect

    # Live properties (used for scroll pattern / window check)
    node.BoundingRectangle = rect

    # Automation element ID used to detect RootWebArea
    node.CachedAutomationId = automation_element_id

    # Children
    kids = children or []
    node.GetCachedChildren.return_value = kids

    # Legacy accessible pattern
    legacy = MagicMock()
    legacy.Role = 0  # Will be looked up via AccessibleRoleNames
    legacy.Value = legacy_value
    legacy.DefaultAction = legacy_default_action
    node.GetLegacyIAccessiblePattern.return_value = legacy

    # Scroll pattern (default: not scrollable)
    if scroll_pattern is None:
        node.GetPattern.return_value = None
    else:
        node.GetPattern.return_value = scroll_pattern

    # Window pattern for modal detection
    wp = MagicMock()
    wp.IsModal = window_pattern_is_modal
    node.GetWindowPattern.return_value = wp

    # DOM detection attributes
    node.ControlTypeName = control_type_name
    node.LocalizedControlType = localized_control_type
    node.Name = name
    node.AcceleratorKey = accelerator_key
    node.HasKeyboardFocus = has_keyboard_focus

    return node


def _make_scroll_pattern(
    vertically_scrollable=True,
    vertical_scroll_percent=25.0,
    horizontally_scrollable=False,
    horizontal_scroll_percent=0.0,
):
    sp = MagicMock()
    sp.VerticallyScrollable = vertically_scrollable
    sp.VerticalScrollPercent = vertical_scroll_percent
    sp.HorizontallyScrollable = horizontally_scrollable
    sp.HorizontalScrollPercent = horizontal_scroll_percent
    return sp


def _make_window_bounding_box():
    """Return a Rect-compatible object covering the whole test area."""
    return _make_rect(0, 0, 1920, 1080)


def _run_traversal(tree, root_node, is_browser=False, children=None):
    """
    Run tree_traversal in subtree_cached mode (no live COM cache calls).

    Returns (interactive_nodes, scrollable_nodes, dom_interactive_nodes, dom_informative_nodes).
    """
    if children is not None:
        root_node.GetCachedChildren.return_value = children

    interactive_nodes = []
    scrollable_nodes = []
    dom_interactive_nodes = []
    dom_informative_nodes = []

    tree.tree_traversal(
        node=root_node,
        window_bounding_box=_make_window_bounding_box(),
        window_name="TestApp",
        is_browser=is_browser,
        interactive_nodes=interactive_nodes,
        scrollable_nodes=scrollable_nodes,
        dom_interactive_nodes=dom_interactive_nodes,
        dom_informative_nodes=dom_informative_nodes,
        is_dom=False,
        is_dialog=False,
        element_cache_req=None,
        children_cache_req=None,
        subtree_cached=True,
    )
    return interactive_nodes, scrollable_nodes, dom_interactive_nodes, dom_informative_nodes


# ---------------------------------------------------------------------------
# TestTreeTraversal
# ---------------------------------------------------------------------------


class TestTreeTraversal:
    """Tests for Tree.tree_traversal() element classification logic."""

    # --- Interactive classification ---

    def test_button_classified_as_interactive(self, tree_instance):
        """ButtonControl is in INTERACTIVE_CONTROL_TYPE_NAMES and is added to interactive list."""
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""

        # Map role 0 -> "PushButton" in INTERACTIVE_ROLES via AccessibleRoleNames patch
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Submit",
        )
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "PushButton"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "Submit"
        assert interactive[0].control_type == "Button"

    def test_edit_classified_as_interactive(self, tree_instance):
        """EditControl gets is_keyboard_focusable=True by control-type fast path."""
        node = _make_uia_node(
            control_type_name="EditControl",
            localized_control_type="edit",
            name="Username",
            is_keyboard_focusable=False,  # Should be overridden by the fast-path set
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = "user@example.com"
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "Text"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "Username"

    def test_checkbox_classified_as_interactive(self, tree_instance):
        node = _make_uia_node(
            control_type_name="CheckBoxControl",
            localized_control_type="check box",
            name="Accept Terms",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "CheckButton"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "Accept Terms"

    def test_listitem_classified_as_interactive(self, tree_instance):
        node = _make_uia_node(
            control_type_name="ListItemControl",
            localized_control_type="list item",
            name="Item 1",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "ListItem"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1

    def test_menuitem_classified_as_interactive(self, tree_instance):
        node = _make_uia_node(
            control_type_name="MenuItemControl",
            localized_control_type="menu item",
            name="File",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "MenuItem"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "File"

    def test_hyperlink_classified_as_interactive(self, tree_instance):
        node = _make_uia_node(
            control_type_name="HyperlinkControl",
            localized_control_type="hyperlink",
            name="Click here",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch(
            "windows_mcp.tree.service.AccessibleRoleNames",
            {0: "Link"},
        ):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "Click here"

    # --- Non-interactive classification ---

    def test_pane_not_classified_as_interactive(self, tree_instance):
        """PaneControl is structural and should not appear in interactive nodes."""
        node = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="Main Panel",
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    def test_text_control_not_classified_as_interactive(self, tree_instance):
        """TextControl is informative only, not interactive."""
        node = _make_uia_node(
            control_type_name="TextControl",
            localized_control_type="text",
            name="Label text",
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    # --- Offscreen filtering ---

    def test_offscreen_button_is_skipped(self, tree_instance):
        """Offscreen non-Edit controls are excluded from interactive nodes."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Hidden Button",
            is_offscreen=True,
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    def test_offscreen_edit_is_kept(self, tree_instance):
        """EditControl is included even when offscreen (e.g. autocomplete dropdowns)."""
        node = _make_uia_node(
            control_type_name="EditControl",
            localized_control_type="edit",
            name="Search",
            is_offscreen=True,
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "Text"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].name == "Search"

    # --- Disabled element handling ---

    def test_disabled_button_is_excluded(self, tree_instance):
        """Disabled elements (CachedIsEnabled=False) are not classified as interactive."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Disabled",
            is_enabled=False,
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    def test_disabled_element_children_are_not_traversed_further_but_no_crash(
        self, tree_instance
    ):
        """
        Disabled parents skip the visible check but children are still pushed to the stack
        because tree_traversal pushes children unconditionally.
        A disabled parent with enabled children yields the enabled child.
        """
        child = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Child Button",
        )
        child.GetCachedChildren.return_value = []

        parent = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="Panel",
            is_enabled=False,
            children=[child],
        )

        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        child.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, parent)

        # Child is enabled; it should be classified
        assert len(interactive) == 1
        assert interactive[0].name == "Child Button"

    # --- Zero-area element filtering ---

    def test_zero_area_element_is_excluded(self, tree_instance):
        """Elements with zero-area bounding rectangles are invisible and skipped."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="ZeroArea",
            bounding_rect=_make_rect(100, 100, 100, 100),  # width=0, height=0
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    def test_non_control_element_is_excluded(self, tree_instance):
        """CachedIsControlElement=False means the element is a raw element, not a control."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Raw",
            is_control_element=False,
        )
        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 0

    # --- Max depth limit ---

    def test_max_depth_limit_stops_traversal(self, tree_instance):
        """When depth == MAX_TREE_DEPTH the node is skipped without error."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Deep Button",
        )

        interactive_nodes = []
        scrollable_nodes = []

        # Call with depth at MAX_TREE_DEPTH -- should skip this frame
        tree_instance.tree_traversal(
            node=node,
            window_bounding_box=_make_window_bounding_box(),
            window_name="TestApp",
            is_browser=False,
            interactive_nodes=interactive_nodes,
            scrollable_nodes=scrollable_nodes,
            dom_interactive_nodes=[],
            dom_informative_nodes=[],
            is_dom=False,
            is_dialog=False,
            subtree_cached=True,
            depth=Tree.MAX_TREE_DEPTH,  # At the limit
        )

        assert len(interactive_nodes) == 0

    def test_depth_just_below_limit_is_processed(self, tree_instance):
        """Depth == MAX_TREE_DEPTH - 1 should still be processed."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Near Limit",
        )
        node.GetCachedChildren.return_value = []

        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        interactive_nodes = []

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            tree_instance.tree_traversal(
                node=node,
                window_bounding_box=_make_window_bounding_box(),
                window_name="TestApp",
                is_browser=False,
                interactive_nodes=interactive_nodes,
                scrollable_nodes=[],
                dom_interactive_nodes=[],
                dom_informative_nodes=[],
                is_dom=False,
                is_dialog=False,
                subtree_cached=True,
                depth=Tree.MAX_TREE_DEPTH - 1,
            )

        assert len(interactive_nodes) == 1

    # --- Dialog detection clears prior nodes ---

    def test_modal_dialog_child_clears_interactive_nodes(self, tree_instance):
        """
        When a WindowControl child reports IsModal=True, the clear_int flag is set,
        which causes interactive_nodes to be cleared before that child is processed.
        This simulates how a modal dialog replaces the parent window's elements.

        Uses create_autospec(WindowControl) so that isinstance(child, WindowControl)
        returns True without needing a real COM object.
        """
        # Build a spec-based mock that satisfies isinstance(obj, WindowControl)
        dialog_child = create_autospec(WindowControl, instance=True)
        dialog_child._is_cached = True
        dialog_child.CachedControlTypeName = "WindowControl"
        dialog_child.CachedLocalizedControlType = "window"
        dialog_child.CachedName = "Alert Dialog"
        dialog_child.CachedIsOffscreen = False
        dialog_child.CachedIsEnabled = True
        dialog_child.CachedIsControlElement = True
        dialog_child.CachedIsKeyboardFocusable = False
        dialog_child.CachedHasKeyboardFocus = False
        dialog_child.CachedAutomationId = ""
        dialog_child.CachedBoundingRectangle = _make_rect(200, 200, 600, 500)
        dialog_child.GetCachedChildren.return_value = []

        wp = MagicMock()
        wp.IsModal = True
        dialog_child.GetWindowPattern.return_value = wp

        # Root window
        root = _make_uia_node(
            control_type_name="WindowControl",
            localized_control_type="window",
            name="Main Window",
        )
        root.GetCachedChildren.return_value = [dialog_child]

        interactive_nodes = [
            TreeElementNode(
                name="Pre-existing",
                control_type="Button",
                bounding_box=BoundingBox(0, 0, 100, 100, 100, 100),
                center=BoundingBox(0, 0, 100, 100, 100, 100).get_center(),
                window_name="Main Window",
                value="",
                shortcut="",
                xpath="",
                is_focused=False,
            )
        ]

        tree_instance.tree_traversal(
            node=root,
            window_bounding_box=_make_window_bounding_box(),
            window_name="TestApp",
            is_browser=False,
            interactive_nodes=interactive_nodes,
            scrollable_nodes=[],
            dom_interactive_nodes=[],
            dom_informative_nodes=[],
            is_dom=False,
            is_dialog=False,
            subtree_cached=True,
        )

        # The modal dialog caused clear_int=True, so interactive_nodes was cleared
        assert len(interactive_nodes) == 0

    # --- Scrollable element detection ---

    def test_scroll_pattern_detected_on_non_interactive_control(self, tree_instance):
        """PaneControl with a VerticallyScrollable scroll pattern is added to scrollable_nodes."""
        scroll_pattern = _make_scroll_pattern(
            vertically_scrollable=True, vertical_scroll_percent=50.0
        )
        # random_point_within_bounding_box uses node.BoundingRectangle with width/height callables
        rect = _make_rect(0, 0, 300, 400)
        node = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="ScrollPane",
            bounding_rect=rect,
            scroll_pattern=scroll_pattern,
        )
        node.GetCachedChildren.return_value = []
        node.BoundingRectangle = rect

        with patch(
            "windows_mcp.tree.service.random_point_within_bounding_box",
            return_value=(150, 200),
        ):
            _, scrollable, _, _ = _run_traversal(tree_instance, node)

        assert len(scrollable) == 1
        assert scrollable[0].name == "ScrollPane"
        assert scrollable[0].vertical_scrollable is True
        assert scrollable[0].vertical_scroll_percent == 50.0

    def test_non_scrollable_pattern_not_added(self, tree_instance):
        """A PaneControl where VerticallyScrollable=False is NOT added to scrollable_nodes."""
        scroll_pattern = _make_scroll_pattern(vertically_scrollable=False)
        node = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="StaticPane",
            scroll_pattern=scroll_pattern,
        )
        node.GetCachedChildren.return_value = []

        with patch(
            "windows_mcp.tree.service.random_point_within_bounding_box",
            return_value=(100, 100),
        ):
            _, scrollable, _, _ = _run_traversal(tree_instance, node)

        assert len(scrollable) == 0

    def test_interactive_control_scroll_pattern_skipped(self, tree_instance):
        """
        Interactive controls (ButtonControl) are excluded from scroll detection
        because the code only checks scroll pattern when the control type is NOT
        in INTERACTIVE_CONTROL_TYPE_NAMES | INFORMATIVE_CONTROL_TYPE_NAMES.
        """
        scroll_pattern = _make_scroll_pattern(vertically_scrollable=True)
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Scrollable Button",
            scroll_pattern=scroll_pattern,
        )
        node.GetCachedChildren.return_value = []

        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            with patch(
                "windows_mcp.tree.service.random_point_within_bounding_box",
                return_value=(100, 100),
            ):
                interactive, scrollable, _, _ = _run_traversal(tree_instance, node)

        # Button is interactive (added), but NOT in scrollable
        assert len(interactive) == 1
        assert len(scrollable) == 0

    def test_scroll_pattern_exception_does_not_crash(self, tree_instance):
        """If GetPattern raises, traversal continues without crashing."""
        node = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="Panel",
        )
        node.GetPattern.side_effect = Exception("COM error")
        node.GetCachedChildren.return_value = []

        _, scrollable, _, _ = _run_traversal(tree_instance, node)

        # No scrollable node added, no exception raised
        assert len(scrollable) == 0

    # --- Shortcut / accelerator key ---

    def test_accelerator_key_propagated_to_node(self, tree_instance):
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Save",
            accelerator_key="Ctrl+S",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert interactive[0].shortcut == "Ctrl+S"

    # --- Keyboard focus ---

    def test_has_keyboard_focus_set_on_node(self, tree_instance):
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Focused",
            has_keyboard_focus=True,
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert interactive[0].is_focused is True

    # --- Informative node collection ---

    def test_text_control_in_browser_dom_adds_informative_node(self, tree_instance):
        """TextControl inside a DOM subtree (is_browser=True, is_dom=True) is informative."""
        text_node = _make_uia_node(
            control_type_name="TextControl",
            localized_control_type="text",
            name="Some Label",
            is_keyboard_focusable=False,
        )
        text_node.GetCachedChildren.return_value = []

        dom_bb = BoundingBox(left=0, top=0, right=1920, bottom=1080, width=1920, height=1080)

        interactive_nodes = []
        dom_informative_nodes = []

        tree_instance.tree_traversal(
            node=text_node,
            window_bounding_box=_make_window_bounding_box(),
            window_name="Browser",
            is_browser=True,
            interactive_nodes=interactive_nodes,
            scrollable_nodes=[],
            dom_interactive_nodes=[],
            dom_informative_nodes=dom_informative_nodes,
            is_dom=True,  # Inside a DOM subtree
            is_dialog=False,
            subtree_cached=True,
            dom_bounding_box=dom_bb,
        )

        assert len(dom_informative_nodes) == 1
        assert dom_informative_nodes[0].text == "Some Label"

    def test_text_control_outside_dom_not_informative(self, tree_instance):
        """TextControl outside a DOM subtree is NOT added to dom_informative_nodes."""
        text_node = _make_uia_node(
            control_type_name="TextControl",
            localized_control_type="text",
            name="Outside DOM",
        )
        text_node.GetCachedChildren.return_value = []

        _, _, _, dom_informative = _run_traversal(tree_instance, text_node, is_browser=True)

        # is_dom=False (default), so no informative nodes collected
        assert len(dom_informative) == 0

    # --- Multi-node tree traversal ---

    def test_multiple_interactive_children_all_collected(self, tree_instance):
        """All interactive children at the same level are collected."""
        btn1 = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Btn1",
        )
        btn1.GetCachedChildren.return_value = []
        legacy1 = MagicMock()
        legacy1.Role = 0
        legacy1.Value = ""
        btn1.GetLegacyIAccessiblePattern.return_value = legacy1

        btn2 = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Btn2",
        )
        btn2.GetCachedChildren.return_value = []
        legacy2 = MagicMock()
        legacy2.Role = 0
        legacy2.Value = ""
        btn2.GetLegacyIAccessiblePattern.return_value = legacy2

        parent = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="Container",
            children=[btn1, btn2],
        )

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, parent)

        assert len(interactive) == 2
        names = {n.name for n in interactive}
        assert names == {"Btn1", "Btn2"}

    def test_nested_button_inside_pane_collected(self, tree_instance):
        """Interactive element nested inside a non-interactive pane is still collected."""
        button = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Nested",
        )
        button.GetCachedChildren.return_value = []
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        button.GetLegacyIAccessiblePattern.return_value = legacy

        pane = _make_uia_node(
            control_type_name="PaneControl",
            localized_control_type="pane",
            name="Outer",
            children=[button],
        )

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, pane)

        assert len(interactive) == 1
        assert interactive[0].name == "Nested"

    # --- Bounding box clipping ---

    def test_element_clipped_to_window_bounds(self, tree_instance):
        """Elements extending beyond the window are clipped to window/screen intersection."""
        # Button at (1900, 1060) to (2000, 1200), screen is 1920x1080
        rect = _make_rect(1900, 1060, 2000, 1200)
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="EdgeBtn",
            bounding_rect=rect,
        )
        node.GetCachedChildren.return_value = []
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert len(interactive) == 1
        assert interactive[0].bounding_box.right == 1920
        assert interactive[0].bounding_box.bottom == 1080

    # --- Value from LegacyIAccessiblePattern ---

    def test_interactive_node_value_from_legacy_pattern(self, tree_instance):
        node = _make_uia_node(
            control_type_name="EditControl",
            localized_control_type="edit",
            name="Field",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = "  hello world  "
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "Text"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert interactive[0].value == "hello world"

    def test_legacy_pattern_none_value_becomes_empty_string(self, tree_instance):
        """If legacy_pattern.Value is None, value is set to empty string."""
        node = _make_uia_node(
            control_type_name="ButtonControl",
            localized_control_type="button",
            name="Btn",
        )
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = None  # Simulate None value
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node)

        assert interactive[0].value == ""

    # --- ImageControl special casing ---

    def test_focusable_image_in_non_browser_is_interactive(self, tree_instance):
        """ImageControl with CachedIsKeyboardFocusable=True in non-browser is interactive."""
        node = _make_uia_node(
            control_type_name="ImageControl",
            localized_control_type="image",
            name="Logo",
            is_keyboard_focusable=True,
        )
        node.GetCachedChildren.return_value = []
        legacy = MagicMock()
        legacy.Role = 0
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {0: "PushButton"}):
            interactive, _, _, _ = _run_traversal(tree_instance, node, is_browser=False)

        assert len(interactive) == 1
        assert interactive[0].name == "Logo"

    def test_non_focusable_image_not_interactive(self, tree_instance):
        """ImageControl with is_keyboard_focusable=False is not interactive."""
        node = _make_uia_node(
            control_type_name="ImageControl",
            localized_control_type="image",
            name="Decorative",
            is_keyboard_focusable=False,
        )
        node.GetCachedChildren.return_value = []

        with patch("windows_mcp.tree.service.AccessibleRoleNames", {}):
            interactive, _, _, _ = _run_traversal(tree_instance, node, is_browser=False)

        assert len(interactive) == 0


# ---------------------------------------------------------------------------
# TestOnFocusChange
# ---------------------------------------------------------------------------


class TestOnFocusChange:
    """Tests for Tree._on_focus_change() debounce logic."""

    def test_first_event_is_not_debounced(self, tree_instance):
        """First focus event with no prior history is always processed."""
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [1, 2, 3]
        mock_element.Name = "Button A"
        mock_element.ControlTypeName = "ButtonControl"

        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            result = tree_instance._on_focus_change(sender)

        # No debounce on first event; result is None (function returns None after logging)
        assert result is None
        # The event should have been recorded
        assert tree_instance._last_focus_event is not None

    def test_same_event_within_one_second_is_debounced(self, tree_instance):
        """Duplicate focus event for the same element within 1s returns None immediately."""
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [5, 6, 7]
        mock_element.Name = "CheckBox"
        mock_element.ControlTypeName = "CheckBoxControl"

        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            # First call sets the state
            tree_instance._on_focus_change(sender)

            # Record the event key and time that was stored
            key_after_first, time_after_first = tree_instance._last_focus_event

            # Second call immediately (within debounce window) -- should be debounced
            result = tree_instance._on_focus_change(sender)

        # Debounced: returns None without updating _last_focus_event time
        assert result is None
        # _last_focus_event should NOT have been updated by the debounced call
        assert tree_instance._last_focus_event[1] == time_after_first

    def test_same_event_after_one_second_is_not_debounced(self, tree_instance):
        """The same element re-focused after >1 second should update _last_focus_event."""
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [10, 11, 12]
        mock_element.Name = "Edit"
        mock_element.ControlTypeName = "EditControl"

        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            # Seed _last_focus_event with a timestamp 2 seconds in the past
            tree_instance._last_focus_event = ((10, 11, 12), time() - 2.0)
            old_time = tree_instance._last_focus_event[1]

            tree_instance._on_focus_change(sender)

            new_time = tree_instance._last_focus_event[1]

        # The event should be processed and the timestamp updated
        assert new_time > old_time

    def test_different_element_is_not_debounced(self, tree_instance):
        """A different element always passes through the debounce check."""
        element_a = MagicMock()
        element_a.GetRuntimeId.return_value = [1, 2, 3]
        element_a.Name = "Button A"
        element_a.ControlTypeName = "ButtonControl"

        element_b = MagicMock()
        element_b.GetRuntimeId.return_value = [4, 5, 6]
        element_b.Name = "Button B"
        element_b.ControlTypeName = "ButtonControl"

        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            side_effect=[element_a, element_b],
        ):
            # Focus element A
            tree_instance._on_focus_change(sender)
            time_after_a = tree_instance._last_focus_event[1]

            # Focus element B immediately -- different key, NOT debounced
            tree_instance._on_focus_change(sender)
            time_after_b = tree_instance._last_focus_event[1]

        # _last_focus_event should have been updated for element B
        assert tree_instance._last_focus_event[0] == (4, 5, 6)
        assert time_after_b >= time_after_a

    def test_focus_change_records_correct_event_key(self, tree_instance):
        """The debounce key is a tuple of the element's runtime ID."""
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [99, 100]
        mock_element.Name = "Tab"
        mock_element.ControlTypeName = "TabItemControl"

        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            tree_instance._on_focus_change(sender)

        event_key, _ = tree_instance._last_focus_event
        assert event_key == (99, 100)

    def test_focus_change_handles_exception_gracefully(self, tree_instance):
        """If CreateControlFromElement raises, the exception propagates (no silent swallow)."""
        sender = MagicMock()

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            side_effect=Exception("COM error"),
        ):
            with pytest.raises(Exception, match="COM error"):
                tree_instance._on_focus_change(sender)


# ---------------------------------------------------------------------------
# TestDomCorrection
# ---------------------------------------------------------------------------


class TestDomCorrection:
    """Tests for Tree._dom_correction() structural corrections."""

    def _make_dom_bb(self):
        return BoundingBox(left=0, top=0, right=1920, bottom=1080, width=1920, height=1080)

    # --- list-item-with-link correction ---

    def test_list_item_with_link_child_pops_last_node(self, tree_instance):
        """A list item whose first child is a link causes the last dom_interactive_node to be removed."""
        link_child = MagicMock()
        link_child.LocalizedControlType = "link"

        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = link_child
        node.ControlTypeName = "ListItemControl"  # Not "GroupControl"

        # Seed the interactive nodes list with a placeholder
        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        assert len(dom_nodes) == 0

    def test_item_with_link_child_also_pops(self, tree_instance):
        """'item' localized type with a link child is treated the same as list item."""
        link_child = MagicMock()
        link_child.LocalizedControlType = "link"

        node = MagicMock()
        node.LocalizedControlType = "item"
        node.GetFirstChildControl.return_value = link_child
        node.ControlTypeName = "CustomControl"

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        assert len(dom_nodes) == 0

    def test_list_item_without_link_child_no_pop(self, tree_instance):
        """A list item whose first child is NOT a link should not trigger this correction."""
        text_child = MagicMock()
        text_child.LocalizedControlType = "text"

        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = text_child
        node.ControlTypeName = "ListItemControl"

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # Node should not have been popped
        assert len(dom_nodes) == 1

    def test_list_item_with_no_child_no_pop(self, tree_instance):
        """A list item with no first child does not trigger the correction."""
        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = None
        node.ControlTypeName = "ListItemControl"

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        assert len(dom_nodes) == 1

    def test_pop_on_empty_list_does_not_crash(self, tree_instance):
        """Correction with empty dom_interactive_nodes list should not crash."""
        link_child = MagicMock()
        link_child.LocalizedControlType = "link"

        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = link_child
        node.ControlTypeName = "ListItemControl"

        dom_nodes = []  # Empty -- pop should be guarded

        # Should not raise IndexError
        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        assert len(dom_nodes) == 0

    # --- GroupControl correction ---

    def test_group_control_non_focusable_just_pops(self, tree_instance):
        """GroupControl that is not keyboard-focusable pops the last node and does not add."""
        node = MagicMock()
        node.LocalizedControlType = "group"
        node.GetFirstChildControl.return_value = None
        node.ControlTypeName = "GroupControl"
        node.CachedControlTypeName = "GroupControl"  # Not in the fast-path set
        node.CachedIsKeyboardFocusable = False

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # Popped but not replaced (not focusable, so exits early)
        assert len(dom_nodes) == 0

    def test_group_control_focusable_with_text_leaf_adds_node(self, tree_instance):
        """
        A keyboard-focusable GroupControl whose deepest child is a TextControl
        should have a new TreeElementNode appended to dom_interactive_nodes.
        """
        text_leaf = MagicMock()
        text_leaf.ControlTypeName = "TextControl"
        text_leaf.GetFirstChildControl.return_value = None
        text_leaf.Name = "  Click Me  "
        text_leaf.LocalizedControlType = "text"

        node = MagicMock()
        node.LocalizedControlType = "group"
        # GetFirstChildControl: first call returns text_leaf, which then has no children
        node.GetFirstChildControl.return_value = text_leaf
        node.ControlTypeName = "GroupControl"
        node.CachedControlTypeName = "GroupControl"
        node.CachedIsKeyboardFocusable = True
        node.AcceleratorKey = ""
        node.HasKeyboardFocus = False

        # Legacy pattern for value
        legacy = MagicMock()
        legacy.Value = "some value"
        node.GetLegacyIAccessiblePattern.return_value = legacy

        # Bounding rectangle
        rect = _make_rect(50, 50, 200, 100)
        node.BoundingRectangle = rect

        dom_nodes = [MagicMock()]  # The node that will be popped

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # The old node was popped and a new one appended
        assert len(dom_nodes) == 1
        assert dom_nodes[0].name == "Click Me"

    def test_group_control_with_interactive_child_returns_none(self, tree_instance):
        """
        If a GroupControl child is itself an interactive control type, the correction
        returns None without adding anything (the child handles its own classification).
        """
        interactive_child = MagicMock()
        interactive_child.ControlTypeName = "ButtonControl"
        interactive_child.GetFirstChildControl.return_value = None

        node = MagicMock()
        node.LocalizedControlType = "group"
        node.GetFirstChildControl.return_value = interactive_child
        node.ControlTypeName = "GroupControl"
        node.CachedControlTypeName = "GroupControl"
        node.CachedIsKeyboardFocusable = True

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # Popped but NOT re-added (early return when child is ButtonControl)
        assert len(dom_nodes) == 0

    def test_group_control_edit_fast_path_is_focusable(self, tree_instance):
        """GroupControl with CachedControlTypeName='EditControl' is treated as focusable."""
        text_leaf = MagicMock()
        text_leaf.ControlTypeName = "TextControl"
        text_leaf.GetFirstChildControl.return_value = None
        text_leaf.Name = "Input Label"
        text_leaf.LocalizedControlType = "text"

        node = MagicMock()
        node.LocalizedControlType = "group"
        node.GetFirstChildControl.return_value = text_leaf
        node.ControlTypeName = "GroupControl"
        node.CachedControlTypeName = "EditControl"  # Fast-path to is_kb_focusable=True
        node.AcceleratorKey = ""
        node.HasKeyboardFocus = False

        legacy = MagicMock()
        legacy.Value = ""
        node.GetLegacyIAccessiblePattern.return_value = legacy

        rect = _make_rect(0, 0, 100, 50)
        node.BoundingRectangle = rect

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        assert len(dom_nodes) == 1
        assert dom_nodes[0].name == "Input Label"

    # --- link-with-heading correction ---

    def test_link_with_heading_child_replaces_node(self, tree_instance):
        """A link element whose first child is a heading gets added as a 'link' type node."""
        heading_child = MagicMock()
        heading_child.LocalizedControlType = "heading"
        heading_child.Name = "  Section Title  "
        heading_child.AcceleratorKey = ""
        heading_child.HasKeyboardFocus = False

        link_legacy = MagicMock()
        link_legacy.Value = "https://example.com"
        heading_child.GetLegacyIAccessiblePattern.return_value = link_legacy

        rect = _make_rect(20, 20, 400, 60)
        heading_child.BoundingRectangle = rect

        node = MagicMock()
        node.LocalizedControlType = "link"
        node.GetFirstChildControl.return_value = heading_child
        node.ControlTypeName = "HyperlinkControl"

        dom_nodes = [MagicMock()]  # Will be popped

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # The old node is replaced by the heading-link node
        assert len(dom_nodes) == 1
        added = dom_nodes[0]
        assert added.name == "Section Title"
        assert added.control_type == "link"
        assert added.value == "Section Title"

    def test_link_without_heading_child_no_correction(self, tree_instance):
        """A link whose first child is not a heading should not trigger this correction."""
        text_child = MagicMock()
        text_child.LocalizedControlType = "text"

        node = MagicMock()
        node.LocalizedControlType = "link"
        node.GetFirstChildControl.return_value = text_child
        node.ControlTypeName = "HyperlinkControl"

        dom_nodes = [MagicMock()]

        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # No change; 'link' with 'text' child doesn't match any correction pattern
        assert len(dom_nodes) == 1

    def test_link_heading_pop_on_empty_list_does_not_crash(self, tree_instance):
        """Pop on empty dom_interactive_nodes for link-heading correction should not crash."""
        heading_child = MagicMock()
        heading_child.LocalizedControlType = "heading"
        heading_child.Name = "Title"
        heading_child.AcceleratorKey = ""
        heading_child.HasKeyboardFocus = False

        link_legacy = MagicMock()
        link_legacy.Value = ""
        heading_child.GetLegacyIAccessiblePattern.return_value = link_legacy

        rect = _make_rect(0, 0, 200, 50)
        heading_child.BoundingRectangle = rect

        node = MagicMock()
        node.LocalizedControlType = "link"
        node.GetFirstChildControl.return_value = heading_child
        node.ControlTypeName = "HyperlinkControl"

        dom_nodes = []  # Empty

        # Should not raise
        tree_instance._dom_correction(node, dom_nodes, "Browser", self._make_dom_bb())

        # One node added (the heading-link), nothing crashed
        assert len(dom_nodes) == 1

    # --- element_has_child_element ---

    def test_element_has_child_element_match(self, tree_instance):
        """element_has_child_element returns True when types match."""
        child = MagicMock()
        child.LocalizedControlType = "link"

        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = child

        result = tree_instance.element_has_child_element(node, "list item", "link")

        assert result is True

    def test_element_has_child_element_no_child(self, tree_instance):
        """element_has_child_element returns False when there is no first child."""
        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = None

        result = tree_instance.element_has_child_element(node, "list item", "link")

        assert result is False

    def test_element_has_child_element_wrong_parent_type(self, tree_instance):
        """element_has_child_element returns None when parent type does not match."""
        node = MagicMock()
        node.LocalizedControlType = "button"

        result = tree_instance.element_has_child_element(node, "list item", "link")

        # Condition is False, function returns None implicitly
        assert result is None

    def test_element_has_child_element_wrong_child_type(self, tree_instance):
        """element_has_child_element returns False when child type does not match."""
        child = MagicMock()
        child.LocalizedControlType = "text"

        node = MagicMock()
        node.LocalizedControlType = "list item"
        node.GetFirstChildControl.return_value = child

        result = tree_instance.element_has_child_element(node, "list item", "link")

        assert result is False


# ---------------------------------------------------------------------------
# Helpers shared by window-cache test classes
# ---------------------------------------------------------------------------


def _make_cached_entry(interactive=None, scrollable=None, informative=None, age=0.0):
    """Build a _CachedWindowNodes with a configurable timestamp age (seconds old)."""
    from windows_mcp.tree.service import _CachedWindowNodes

    return _CachedWindowNodes(
        interactive=interactive or [],
        scrollable=scrollable or [],
        informative=informative or [],
        dom_node=None,
        timestamp=time() - age,
    )


def _make_cached_interactive(name="CachedBtn", age=0.0):
    """Return a single-element interactive list and a matching cache entry."""
    node = TreeElementNode(
        name=name,
        control_type="Button",
        bounding_box=BoundingBox(0, 0, 100, 100, 100, 100),
        center=BoundingBox(0, 0, 100, 100, 100, 100).get_center(),
        window_name="TestApp",
        value="",
        shortcut="",
        xpath="",
        is_focused=False,
    )
    entry = _make_cached_entry(interactive=[node], age=age)
    return [node], entry


# ---------------------------------------------------------------------------
# TestWindowCacheHits
# ---------------------------------------------------------------------------


class TestWindowCacheHits:
    """Tests for per-window cache serving background windows from cache."""

    def test_single_background_window_served_from_cache(self, tree_with_desktop):
        """
        A background window with a fresh, non-dirty cache entry is served
        without any traversal.  native_capture_tree must not be called.
        """
        tree, desktop = tree_with_desktop
        hwnd_active = 100
        hwnd_bg = 200
        nodes, entry = _make_cached_interactive("CachedBg")
        tree._window_cache[hwnd_bg] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            # Active window: ControlFromHandle returns a non-browser mock
            active_ctrl = MagicMock()
            active_ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = active_ctrl
            desktop.is_window_browser.return_value = False
            # Rust returns None so active window falls to task_inputs but no workers start
            mock_rust.return_value = None

            interactive, _, _, _ = tree.get_window_wise_nodes(
                windows_handles=[hwnd_active, hwnd_bg],
                active_window_flag=True,
            )

        # CachedBg must appear in results (served from cache)
        assert any(n.name == "CachedBg" for n in interactive)
        # native_capture_tree is called only for the active window handle
        call_args = mock_rust.call_args
        assert hwnd_bg not in call_args[0][0]

    def test_multiple_background_windows_all_served_from_cache(self, tree_with_desktop):
        """
        Multiple background windows with valid cache entries are all served
        from cache and native_capture_tree receives only the active handle.
        """
        tree, desktop = tree_with_desktop
        hwnd_active = 1
        hwnd_bg1 = 2
        hwnd_bg2 = 3

        _, entry1 = _make_cached_interactive("Bg1Btn")
        _, entry2 = _make_cached_interactive("Bg2Btn")
        tree._window_cache[hwnd_bg1] = entry1
        tree._window_cache[hwnd_bg2] = entry2

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            active_ctrl = MagicMock()
            active_ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = active_ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            interactive, _, _, _ = tree.get_window_wise_nodes(
                windows_handles=[hwnd_active, hwnd_bg1, hwnd_bg2],
                active_window_flag=True,
            )

        names = {n.name for n in interactive}
        assert "Bg1Btn" in names
        assert "Bg2Btn" in names
        rust_handles_arg = mock_rust.call_args[0][0]
        assert hwnd_bg1 not in rust_handles_arg
        assert hwnd_bg2 not in rust_handles_arg

    def test_all_cache_hit_returns_early_without_traversal(self, tree_with_desktop):
        """
        When every handle has a valid cache entry (no active_window_flag),
        get_window_wise_nodes returns early and neither native_capture_tree
        nor ControlFromHandle is called at all.
        """
        tree, desktop = tree_with_desktop
        hwnd1 = 10
        hwnd2 = 20

        _, entry1 = _make_cached_interactive("W1")
        _, entry2 = _make_cached_interactive("W2")
        tree._window_cache[hwnd1] = entry1
        tree._window_cache[hwnd2] = entry2

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            interactive, scrollable, informative, dom_node = (
                tree.get_window_wise_nodes(
                    windows_handles=[hwnd1, hwnd2],
                    active_window_flag=False,
                )
            )

        mock_rust.assert_not_called()
        mock_cfh.assert_not_called()
        assert len(interactive) == 2

    def test_active_window_always_re_traversed_even_with_valid_cache(self, tree_with_desktop):
        """
        The first handle when active_window_flag=True is always sent for
        traversal regardless of a warm, non-dirty cache entry.
        """
        tree, desktop = tree_with_desktop
        hwnd_active = 999
        nodes, entry = _make_cached_interactive("ShouldNotReturn")
        tree._window_cache[hwnd_active] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            active_ctrl = MagicMock()
            active_ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = active_ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None  # No Rust results

            tree.get_window_wise_nodes(
                windows_handles=[hwnd_active],
                active_window_flag=True,
            )

        # ControlFromHandle must have been called for the active handle
        mock_cfh.assert_called()
        called_handles = [call.args[0] for call in mock_cfh.call_args_list]
        assert hwnd_active in called_handles

    def test_cache_dom_node_merged_into_result(self, tree_with_desktop):
        """
        A cached dom_node from a background window is propagated to the
        caller's dom_node return value.
        """
        from windows_mcp.tree.service import _CachedWindowNodes
        from windows_mcp.tree.views import ScrollElementNode

        tree, desktop = tree_with_desktop
        hwnd_bg = 55
        mock_dom = MagicMock(spec=ScrollElementNode)
        entry = _CachedWindowNodes(
            interactive=[],
            scrollable=[],
            informative=[],
            dom_node=mock_dom,
            timestamp=time(),
        )
        tree._window_cache[hwnd_bg] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree"),
            patch("windows_mcp.tree.service.ControlFromHandle"),
        ):
            _, _, _, dom_node = tree.get_window_wise_nodes(
                windows_handles=[hwnd_bg],
                active_window_flag=False,
            )

        assert dom_node is mock_dom


# ---------------------------------------------------------------------------
# TestWindowCacheInvalidation
# ---------------------------------------------------------------------------


class TestWindowCacheInvalidation:
    """Tests for dirty-flag invalidation and cache re-traversal."""

    def test_dirty_window_is_re_traversed(self, tree_with_desktop):
        """
        A window marked dirty via _invalidate_window_cache is excluded from
        cache and sent to traversal even though a cache entry exists.
        """
        tree, desktop = tree_with_desktop
        hwnd = 300

        _, entry = _make_cached_interactive("OldData")
        tree._window_cache[hwnd] = entry
        tree._invalidate_window_cache(hwnd)

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None  # Force Python path

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        # Must reach ControlFromHandle, confirming traversal was triggered
        mock_cfh.assert_called_with(hwnd)

    def test_dirty_set_cleared_after_call(self, tree_with_desktop):
        """
        After get_window_wise_nodes completes, _dirty_windows is empty
        regardless of how many handles were dirty.
        """
        tree, desktop = tree_with_desktop
        hwnd_a = 400
        hwnd_b = 401
        tree._invalidate_window_cache(hwnd_a)
        tree._invalidate_window_cache(hwnd_b)

        assert len(tree._dirty_windows) == 2

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            tree.get_window_wise_nodes(
                windows_handles=[hwnd_a, hwnd_b],
                active_window_flag=False,
            )

        assert len(tree._dirty_windows) == 0

    def test_invalidate_window_cache_zero_hwnd_is_no_op(self, tree_instance):
        """_invalidate_window_cache(0) must not add 0 to the dirty set."""
        tree_instance._invalidate_window_cache(0)

        assert 0 not in tree_instance._dirty_windows
        assert len(tree_instance._dirty_windows) == 0

    def test_non_dirty_window_beside_dirty_window_still_cached(self, tree_with_desktop):
        """
        When one background window is dirty and another is clean, only the
        dirty window is re-traversed; the clean window is served from cache.
        """
        tree, desktop = tree_with_desktop
        hwnd_dirty = 500
        hwnd_clean = 501

        _, entry_dirty = _make_cached_interactive("Dirty")
        _, entry_clean = _make_cached_interactive("Clean")
        tree._window_cache[hwnd_dirty] = entry_dirty
        tree._window_cache[hwnd_clean] = entry_clean
        tree._invalidate_window_cache(hwnd_dirty)

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            interactive, _, _, _ = tree.get_window_wise_nodes(
                windows_handles=[hwnd_dirty, hwnd_clean],
                active_window_flag=False,
            )

        # ControlFromHandle called only for the dirty handle, not the clean one
        called_handles = [call.args[0] for call in mock_cfh.call_args_list]
        assert hwnd_dirty in called_handles
        assert hwnd_clean not in called_handles
        # Clean window's cached node appears in results
        assert any(n.name == "Clean" for n in interactive)


# ---------------------------------------------------------------------------
# TestWindowCacheTTL
# ---------------------------------------------------------------------------


class TestWindowCacheTTL:
    """Tests for TTL-based cache expiry (30-second window)."""

    def test_expired_cache_entry_triggers_re_traversal(self, tree_with_desktop):
        """
        A cache entry older than _WINDOW_CACHE_TTL (30 s) must be bypassed
        and the window re-traversed.
        """
        tree, desktop = tree_with_desktop
        hwnd = 600

        # Age = 31 seconds -- past the 30-second TTL
        _, entry = _make_cached_interactive("Stale", age=31.0)
        tree._window_cache[hwnd] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        mock_cfh.assert_called_with(hwnd)

    def test_fresh_cache_entry_is_not_re_traversed(self, tree_with_desktop):
        """
        A cache entry younger than _WINDOW_CACHE_TTL must be served from
        cache without any traversal.
        """
        tree, desktop = tree_with_desktop
        hwnd = 700

        # Age = 5 seconds -- well within the 30-second TTL
        _, entry = _make_cached_interactive("Fresh", age=5.0)
        tree._window_cache[hwnd] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            interactive, _, _, _ = tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        mock_rust.assert_not_called()
        mock_cfh.assert_not_called()
        assert any(n.name == "Fresh" for n in interactive)

    def test_entry_at_exact_ttl_boundary_is_re_traversed(self, tree_with_desktop):
        """
        An entry whose age equals _WINDOW_CACHE_TTL exactly does NOT satisfy
        the strict less-than check and must be re-traversed.
        """
        from windows_mcp.tree.service import _WINDOW_CACHE_TTL, _CachedWindowNodes

        tree, desktop = tree_with_desktop
        hwnd = 800

        # Patch time() so we can set age precisely to the TTL value
        base_time = 10_000.0
        entry_timestamp = base_time - _WINDOW_CACHE_TTL  # exactly at the boundary

        entry = _CachedWindowNodes(timestamp=entry_timestamp)
        tree._window_cache[hwnd] = entry

        with (
            patch("windows_mcp.tree.service.time", return_value=base_time),
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        mock_cfh.assert_called_with(hwnd)

    def test_just_under_ttl_boundary_is_served_from_cache(self, tree_with_desktop):
        """
        An entry one millisecond under _WINDOW_CACHE_TTL still satisfies the
        strict less-than check and is served from cache.
        """
        from windows_mcp.tree.service import _WINDOW_CACHE_TTL, _CachedWindowNodes

        tree, desktop = tree_with_desktop
        hwnd = 900
        base_time = 10_000.0
        age = _WINDOW_CACHE_TTL - 0.001
        entry_timestamp = base_time - age

        nodes, _ = _make_cached_interactive("JustUnder")

        entry = _CachedWindowNodes(interactive=nodes, timestamp=entry_timestamp)
        tree._window_cache[hwnd] = entry

        with (
            patch("windows_mcp.tree.service.time", return_value=base_time),
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            interactive, _, _, _ = tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        mock_rust.assert_not_called()
        mock_cfh.assert_not_called()
        assert any(n.name == "JustUnder" for n in interactive)


# ---------------------------------------------------------------------------
# TestWindowCachePruning
# ---------------------------------------------------------------------------


class TestWindowCachePruning:
    """Tests for stale-entry pruning when windows disappear."""

    def test_vanished_window_removed_from_cache(self, tree_with_desktop):
        """
        A handle that appears in _window_cache but is absent from the current
        windows_handles list must be pruned from the cache after the call.
        """
        tree, desktop = tree_with_desktop
        hwnd_present = 1000
        hwnd_gone = 1001

        _, entry_present = _make_cached_interactive("Present")
        _, entry_gone = _make_cached_interactive("Gone")
        tree._window_cache[hwnd_present] = entry_present
        tree._window_cache[hwnd_gone] = entry_gone

        with (
            patch("windows_mcp.tree.service.native_capture_tree"),
            patch("windows_mcp.tree.service.ControlFromHandle"),
        ):
            tree.get_window_wise_nodes(
                windows_handles=[hwnd_present],
                active_window_flag=False,
            )

        assert hwnd_gone not in tree._window_cache
        assert hwnd_present in tree._window_cache

    def test_multiple_vanished_windows_all_pruned(self, tree_with_desktop):
        """All stale cache keys not in the current handle list are pruned."""
        tree, desktop = tree_with_desktop
        hwnd_active = 2000
        stale_handles = [2001, 2002, 2003]

        for h in stale_handles:
            _, entry = _make_cached_interactive(f"Stale{h}")
            tree._window_cache[h] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            tree.get_window_wise_nodes(
                windows_handles=[hwnd_active],
                active_window_flag=True,
            )

        for h in stale_handles:
            assert h not in tree._window_cache

    def test_empty_handle_list_clears_entire_cache(self, tree_with_desktop):
        """
        Passing an empty windows_handles list causes all cache entries to be
        pruned because no handle appears in the current set.
        """
        tree, desktop = tree_with_desktop
        for h in [3000, 3001, 3002]:
            _, entry = _make_cached_interactive(f"W{h}")
            tree._window_cache[h] = entry

        with (
            patch("windows_mcp.tree.service.native_capture_tree"),
            patch("windows_mcp.tree.service.ControlFromHandle"),
        ):
            tree.get_window_wise_nodes(
                windows_handles=[],
                active_window_flag=False,
            )

        assert len(tree._window_cache) == 0


# ---------------------------------------------------------------------------
# TestWindowCachePopulation
# ---------------------------------------------------------------------------


class TestWindowCachePopulation:
    """Tests that traversal results are stored back into _window_cache."""

    def test_rust_fast_path_populates_cache(self, tree_with_desktop):
        """
        When native_capture_tree returns a snapshot, the result is stored in
        _window_cache under the corresponding HWND.
        """
        tree, desktop = tree_with_desktop
        hwnd = 4000

        rust_snapshot = {
            "name": "TestApp",
            "bounding_rect": [0.0, 0.0, 800.0, 600.0],
            "control_type": "Window",
            "is_control_element": True,
            "is_enabled": True,
            "is_offscreen": False,
            "is_keyboard_focusable": False,
            "children": [
                {
                    "name": "RustButton",
                    "control_type": "Button",
                    "bounding_rect": [10.0, 10.0, 110.0, 50.0],
                    "is_control_element": True,
                    "is_enabled": True,
                    "is_offscreen": False,
                    "is_keyboard_focusable": True,
                    "children": [],
                }
            ],
        }

        with (
            patch(
                "windows_mcp.tree.service.native_capture_tree",
                return_value=[rust_snapshot],
            ),
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            # ControlFromHandle called for browser check; returns non-browser control
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        assert hwnd in tree._window_cache
        cached = tree._window_cache[hwnd]
        assert cached.timestamp > 0
        # Interactive list populated with the RustButton element
        assert any(n.name == "RustButton" for n in cached.interactive)

    def test_cache_timestamp_is_set_after_rust_traversal(self, tree_with_desktop):
        """
        The cache entry timestamp is set to approximately the time of the call,
        not zero, after a Rust fast-path traversal.
        """
        tree, desktop = tree_with_desktop
        hwnd = 4100
        before = time()

        rust_snapshot = {
            "name": "App",
            "bounding_rect": [0.0, 0.0, 800.0, 600.0],
            "control_type": "Window",
            "is_control_element": True,
            "is_enabled": True,
            "is_offscreen": False,
            "is_keyboard_focusable": False,
            "children": [],
        }

        with (
            patch(
                "windows_mcp.tree.service.native_capture_tree",
                return_value=[rust_snapshot],
            ),
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        after = time()
        cached = tree._window_cache[hwnd]
        assert before <= cached.timestamp <= after


# ---------------------------------------------------------------------------
# TestWindowCacheEventInvalidation
# ---------------------------------------------------------------------------


class TestWindowCacheEventInvalidation:
    """Tests for cache invalidation triggered by WatchDog events."""

    def test_on_focus_change_marks_window_dirty(self, tree_instance):
        """
        _on_focus_change populates _dirty_windows with the HWND of the
        element that received focus.
        """
        hwnd = 5000

        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [1, 2, 3]
        mock_element.NativeWindowHandle = hwnd
        mock_element.Name = "SomeButton"
        mock_element.ControlTypeName = "ButtonControl"

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            tree_instance._on_focus_change(MagicMock())

        assert hwnd in tree_instance._dirty_windows

    def test_on_focus_change_multiple_windows_all_dirty(self, tree_instance):
        """
        Successive _on_focus_change calls for elements in different windows
        accumulate all their HWNDs in _dirty_windows.
        """
        hwnd_a = 5100
        hwnd_b = 5101

        elem_a = MagicMock()
        elem_a.GetRuntimeId.return_value = [10]
        elem_a.NativeWindowHandle = hwnd_a
        elem_a.Name = "A"
        elem_a.ControlTypeName = "ButtonControl"

        elem_b = MagicMock()
        elem_b.GetRuntimeId.return_value = [20]
        elem_b.NativeWindowHandle = hwnd_b
        elem_b.Name = "B"
        elem_b.ControlTypeName = "EditControl"

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            side_effect=[elem_a, elem_b],
        ):
            tree_instance._on_focus_change(MagicMock())
            # Second call has a different runtime ID so it is not debounced
            tree_instance._on_focus_change(MagicMock())

        assert hwnd_a in tree_instance._dirty_windows
        assert hwnd_b in tree_instance._dirty_windows

    def test_on_focus_change_nativewindowhandle_exception_does_not_propagate(
        self, tree_instance
    ):
        """
        If accessing NativeWindowHandle raises, the exception is swallowed
        and _dirty_windows remains empty.
        """
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [99]
        mock_element.Name = "X"
        mock_element.ControlTypeName = "ButtonControl"
        type(mock_element).NativeWindowHandle = property(
            lambda self: (_ for _ in ()).throw(Exception("COM error"))
        )

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            # Should not raise
            tree_instance._on_focus_change(MagicMock())

        assert len(tree_instance._dirty_windows) == 0

    def test_on_property_change_marks_window_dirty(self, tree_instance):
        """
        _on_property_change marks the element's HWND dirty in _dirty_windows.
        """
        hwnd = 6000

        mock_element = MagicMock()
        mock_element.NativeWindowHandle = hwnd
        mock_element.Name = "Checkbox"
        mock_element.ControlTypeName = "CheckBoxControl"

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            tree_instance._on_property_change(MagicMock(), propertyId=30003, newValue=True)

        assert hwnd in tree_instance._dirty_windows

    def test_on_property_change_exception_is_swallowed(self, tree_instance):
        """
        If CreateControlFromElement raises inside _on_property_change, the
        exception is caught and _dirty_windows remains unchanged.
        """
        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            side_effect=Exception("COM failure"),
        ):
            # Must not propagate
            tree_instance._on_property_change(MagicMock(), propertyId=30003, newValue=None)

        assert len(tree_instance._dirty_windows) == 0

    def test_focus_then_traversal_dirty_cleared(self, tree_with_desktop):
        """
        A window dirtied by _on_focus_change has its dirty flag cleared after
        the next get_window_wise_nodes call completes.
        """
        tree, desktop = tree_with_desktop
        hwnd = 7000

        _, entry = _make_cached_interactive("OldNode")
        tree._window_cache[hwnd] = entry

        # Mark the window dirty via the focus event path
        mock_element = MagicMock()
        mock_element.GetRuntimeId.return_value = [77]
        mock_element.NativeWindowHandle = hwnd
        mock_element.Name = "Focused"
        mock_element.ControlTypeName = "EditControl"

        with patch(
            "windows_mcp.tree.service.Control.CreateControlFromElement",
            return_value=mock_element,
        ):
            tree._on_focus_change(MagicMock())

        assert hwnd in tree._dirty_windows

        # Now traverse; dirty set should be cleared afterwards
        with (
            patch("windows_mcp.tree.service.native_capture_tree") as mock_rust,
            patch("windows_mcp.tree.service.ControlFromHandle") as mock_cfh,
        ):
            ctrl = MagicMock()
            ctrl.ClassName = "NotProgman"
            mock_cfh.return_value = ctrl
            desktop.is_window_browser.return_value = False
            mock_rust.return_value = None

            tree.get_window_wise_nodes(
                windows_handles=[hwnd],
                active_window_flag=False,
            )

        assert hwnd not in tree._dirty_windows
