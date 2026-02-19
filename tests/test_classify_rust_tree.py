"""Tests for Tree._classify_rust_tree() -- Rust fast-path element classifier.

Validates classification of interactive and scrollable elements from
Rust TreeElementSnapshot dicts (cached properties only, no live COM).
"""

from unittest.mock import MagicMock

import pytest

from windows_mcp.desktop.views import Size
from windows_mcp.tree.service import Tree


@pytest.fixture
def tree():
    mock_desktop = MagicMock()
    mock_desktop.get_screen_size.return_value = Size(width=1920, height=1080)
    return Tree(mock_desktop)


def _make_elem(
    name="Test",
    control_type="Button",
    localized_control_type="button",
    bounding_rect=None,
    is_control_element=True,
    is_enabled=True,
    is_offscreen=False,
    is_keyboard_focusable=True,
    has_keyboard_focus=False,
    accelerator_key="",
    automation_id="",
    class_name="",
    children=None,
    depth=0,
):
    """Build a minimal Rust TreeElementSnapshot dict."""
    return {
        "name": name,
        "automation_id": automation_id,
        "control_type": control_type,
        "localized_control_type": localized_control_type,
        "class_name": class_name,
        "bounding_rect": bounding_rect or [100, 100, 200, 200],
        "is_offscreen": is_offscreen,
        "is_enabled": is_enabled,
        "is_control_element": is_control_element,
        "has_keyboard_focus": has_keyboard_focus,
        "is_keyboard_focusable": is_keyboard_focusable,
        "accelerator_key": accelerator_key,
        "depth": depth,
        "children": children or [],
    }


class TestInteractiveClassification:
    """Test interactive element detection from Rust snapshots."""

    def test_button_is_interactive(self, tree):
        snap = _make_elem(name="OK", control_type="Button")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "OK"
        assert nodes[0].control_type == "Button"

    def test_edit_is_interactive(self, tree):
        snap = _make_elem(name="Username", control_type="Edit", localized_control_type="edit")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "Username"

    def test_checkbox_is_interactive(self, tree):
        snap = _make_elem(control_type="CheckBox", localized_control_type="check box")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_combobox_is_interactive(self, tree):
        snap = _make_elem(control_type="ComboBox", localized_control_type="combo box")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_hyperlink_is_interactive(self, tree):
        snap = _make_elem(control_type="Hyperlink", localized_control_type="hyperlink")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_document_is_interactive(self, tree):
        snap = _make_elem(control_type="Document", localized_control_type="document")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_listitem_is_interactive(self, tree):
        snap = _make_elem(control_type="ListItem", localized_control_type="list item")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_tabitem_is_interactive(self, tree):
        snap = _make_elem(control_type="TabItem", localized_control_type="tab item")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_treeitem_is_interactive(self, tree):
        snap = _make_elem(control_type="TreeItem", localized_control_type="tree item")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_focusable_image_is_interactive(self, tree):
        snap = _make_elem(
            control_type="Image",
            localized_control_type="image",
            is_keyboard_focusable=True,
        )
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_non_focusable_image_is_not_interactive(self, tree):
        snap = _make_elem(
            control_type="Image",
            localized_control_type="image",
            is_keyboard_focusable=False,
        )
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0


class TestNonInteractiveElements:
    """Test elements that should NOT be classified as interactive."""

    def test_text_is_not_interactive(self, tree):
        snap = _make_elem(control_type="Text", localized_control_type="text")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_pane_is_not_interactive(self, tree):
        snap = _make_elem(control_type="Pane", localized_control_type="pane")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_group_is_not_interactive(self, tree):
        snap = _make_elem(control_type="Group", localized_control_type="group")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_custom_is_not_interactive(self, tree):
        snap = _make_elem(control_type="Custom", localized_control_type="custom")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_statusbar_is_not_interactive(self, tree):
        snap = _make_elem(control_type="StatusBar", localized_control_type="status bar")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_unknown_type_is_not_interactive(self, tree):
        snap = _make_elem(control_type="Unknown")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0


class TestFilteringLogic:
    """Test element filtering: offscreen, disabled, non-control, zero area."""

    def test_disabled_element_skipped(self, tree):
        snap = _make_elem(control_type="Button", is_enabled=False)
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_non_control_element_skipped(self, tree):
        snap = _make_elem(control_type="Button", is_control_element=False)
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_offscreen_element_skipped(self, tree):
        snap = _make_elem(control_type="Button", is_offscreen=True)
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_offscreen_edit_is_kept(self, tree):
        """Edit controls are kept even when offscreen (auto-complete dropdowns, etc)."""
        snap = _make_elem(
            control_type="Edit",
            localized_control_type="edit",
            is_offscreen=True,
        )
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1

    def test_zero_area_element_skipped(self, tree):
        snap = _make_elem(control_type="Button", bounding_rect=[100, 100, 100, 100])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_negative_area_element_skipped(self, tree):
        """Inverted bounding rect (right < left)."""
        snap = _make_elem(control_type="Button", bounding_rect=[200, 200, 100, 100])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0


class TestBoundingBoxIntersection:
    """Test intersection with window and screen boundaries."""

    def test_element_fully_inside_window(self, tree):
        snap = _make_elem(bounding_rect=[100, 100, 200, 200])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 500, 500], nodes)
        assert len(nodes) == 1
        bb = nodes[0].bounding_box
        assert bb.left == 100
        assert bb.top == 100
        assert bb.right == 200
        assert bb.bottom == 200

    def test_element_clipped_by_window(self, tree):
        snap = _make_elem(bounding_rect=[400, 400, 600, 600])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 500, 500], nodes)
        assert len(nodes) == 1
        bb = nodes[0].bounding_box
        assert bb.right == 500
        assert bb.bottom == 500

    def test_element_clipped_by_screen(self, tree):
        """Element extends beyond the 1920x1080 screen."""
        snap = _make_elem(bounding_rect=[1900, 1060, 2000, 1200])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 2000, 1200], nodes)
        assert len(nodes) == 1
        bb = nodes[0].bounding_box
        assert bb.right == 1920
        assert bb.bottom == 1080

    def test_element_fully_outside_window(self, tree):
        snap = _make_elem(bounding_rect=[600, 600, 700, 700])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 500, 500], nodes)
        assert len(nodes) == 0

    def test_element_fully_outside_screen(self, tree):
        snap = _make_elem(bounding_rect=[2000, 2000, 2100, 2100])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 3000, 3000], nodes)
        assert len(nodes) == 0


class TestNodeAttributes:
    """Test that output TreeElementNode attributes are set correctly."""

    def test_name_stripped(self, tree):
        snap = _make_elem(name="  OK Button  ")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].name == "OK Button"

    def test_window_name_preserved(self, tree):
        snap = _make_elem()
        nodes = []
        tree._classify_rust_tree(snap, "My Window", [0, 0, 1920, 1080], nodes)
        assert nodes[0].window_name == "My Window"

    def test_control_type_uses_localized_title(self, tree):
        snap = _make_elem(control_type="Button", localized_control_type="push button")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].control_type == "Push Button"

    def test_accelerator_key_preserved(self, tree):
        snap = _make_elem(accelerator_key="Ctrl+S")
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].shortcut == "Ctrl+S"

    def test_keyboard_focus(self, tree):
        snap = _make_elem(has_keyboard_focus=True)
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].is_focused is True

    def test_center_computed_correctly(self, tree):
        snap = _make_elem(bounding_rect=[100, 200, 300, 400])
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].center.x == 200
        assert nodes[0].center.y == 300

    def test_value_is_empty_string(self, tree):
        """Rust path doesn't have LegacyIAccessiblePattern value."""
        snap = _make_elem()
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].value == ""

    def test_xpath_is_empty_string(self, tree):
        snap = _make_elem()
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].xpath == ""


class TestChildTraversal:
    """Test stack-based tree walking."""

    def test_children_are_traversed(self, tree):
        child1 = _make_elem(name="Child1", bounding_rect=[10, 10, 50, 50])
        child2 = _make_elem(name="Child2", bounding_rect=[60, 10, 100, 50])
        parent = _make_elem(
            name="Parent",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 200, 200],
            children=[child1, child2],
        )
        nodes = []
        tree._classify_rust_tree(parent, "TestApp", [0, 0, 1920, 1080], nodes)
        # Parent (Pane) is not interactive, but children (Button) are
        assert len(nodes) == 2
        names = {n.name for n in nodes}
        assert names == {"Child1", "Child2"}

    def test_deeply_nested_children(self, tree):
        """Verify iterative traversal handles deep nesting."""
        # Build 20-deep nested structure
        inner = _make_elem(name="Deepest", bounding_rect=[10, 10, 50, 50])
        for i in range(20):
            inner = _make_elem(
                name=f"Level{i}",
                control_type="Group",
                localized_control_type="group",
                bounding_rect=[0, 0, 200, 200],
                children=[inner],
            )
        nodes = []
        tree._classify_rust_tree(inner, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "Deepest"

    def test_non_control_parent_children_still_traversed(self, tree):
        """Non-control elements are skipped but their children should still be walked."""
        child = _make_elem(name="NestedButton", bounding_rect=[10, 10, 50, 50])
        parent = _make_elem(
            control_type="Custom",
            is_control_element=False,
            children=[child],
        )
        nodes = []
        tree._classify_rust_tree(parent, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "NestedButton"

    def test_disabled_parent_children_still_traversed(self, tree):
        """Disabled elements are skipped but their children should still be walked."""
        child = _make_elem(name="ChildBtn", bounding_rect=[10, 10, 50, 50])
        parent = _make_elem(
            control_type="Group",
            is_enabled=False,
            children=[child],
        )
        nodes = []
        tree._classify_rust_tree(parent, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "ChildBtn"


class TestScrollHeuristic:
    """Test heuristic scroll detection for structural containers."""

    def test_pane_with_overflow_detected_as_scrollable(self, tree):
        # Child extends below parent bottom (300 > 200) -> overflow
        child = _make_elem(
            name="Content",
            bounding_rect=[10, 10, 190, 300],
        )
        pane = _make_elem(
            name="ScrollArea",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        interactive = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], interactive, scrollable)
        assert len(scrollable) == 1
        assert scrollable[0].name == "ScrollArea"
        assert scrollable[0].vertical_scrollable is True
        assert scrollable[0].vertical_scroll_percent == -1  # Unknown

    def test_pane_without_overflow_not_scrollable(self, tree):
        child = _make_elem(
            name="Content",
            bounding_rect=[10, 10, 190, 190],  # Fits within parent
        )
        pane = _make_elem(
            name="Container",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        interactive = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], interactive, scrollable)
        assert len(scrollable) == 0

    def test_pane_with_no_children_not_scrollable(self, tree):
        pane = _make_elem(
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 200, 200],
            children=[],
        )
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert len(scrollable) == 0

    def test_list_with_overflow_detected_as_scrollable(self, tree):
        child = _make_elem(bounding_rect=[10, 10, 190, 300])
        lst = _make_elem(
            name="ItemList",
            control_type="List",
            localized_control_type="list",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(lst, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert len(scrollable) == 1

    def test_tree_control_with_overflow_detected(self, tree):
        child = _make_elem(bounding_rect=[10, 10, 190, 500])
        tree_ctrl = _make_elem(
            name="FileTree",
            control_type="Tree",
            localized_control_type="tree",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(tree_ctrl, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert len(scrollable) == 1

    def test_button_never_scrollable(self, tree):
        """Interactive control types should not be detected as scrollable."""
        child = _make_elem(bounding_rect=[0, 0, 100, 500])
        btn = _make_elem(
            control_type="Button",
            bounding_rect=[0, 0, 100, 100],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(btn, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert len(scrollable) == 0

    def test_scrollable_nodes_none_disables_detection(self, tree):
        """When scrollable_nodes is None, skip scroll detection entirely."""
        child = _make_elem(bounding_rect=[0, 0, 200, 500])
        pane = _make_elem(
            control_type="Pane",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        interactive = []
        # scrollable_nodes=None (default) -- no scroll detection
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], interactive)
        # Should not raise, just skip scroll detection

    def test_scrollable_clipped_to_screen(self, tree):
        """Scrollable area should be clipped to screen bounds."""
        child = _make_elem(bounding_rect=[1800, 900, 2000, 1200])
        pane = _make_elem(
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[1800, 900, 2000, 1100],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 2000, 1200], [], scrollable)
        assert len(scrollable) == 1
        assert scrollable[0].bounding_box.right == 1920
        assert scrollable[0].bounding_box.bottom == 1080


class TestMissingFields:
    """Test handling of malformed/incomplete snapshot dicts."""

    def test_missing_name_defaults_empty(self, tree):
        snap = _make_elem()
        del snap["name"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1
        assert nodes[0].name == ""

    def test_missing_control_type_defaults_unknown(self, tree):
        snap = _make_elem()
        del snap["control_type"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0  # "Unknown" is not interactive

    def test_missing_bounding_rect_defaults_zero(self, tree):
        snap = _make_elem()
        del snap["bounding_rect"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0  # Zero area

    def test_missing_children_key(self, tree):
        snap = _make_elem()
        del snap["children"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 1  # Element itself is still classified

    def test_missing_is_control_element_defaults_false(self, tree):
        snap = _make_elem()
        del snap["is_control_element"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0  # Defaults to False -> skipped

    def test_missing_is_enabled_defaults_false(self, tree):
        snap = _make_elem()
        del snap["is_enabled"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0  # Defaults to False -> skipped

    def test_empty_snapshot(self, tree):
        """Completely empty dict should not crash."""
        nodes = []
        tree._classify_rust_tree({}, "TestApp", [0, 0, 1920, 1080], nodes)
        assert len(nodes) == 0

    def test_missing_accelerator_key(self, tree):
        snap = _make_elem()
        del snap["accelerator_key"]
        nodes = []
        tree._classify_rust_tree(snap, "TestApp", [0, 0, 1920, 1080], nodes)
        assert nodes[0].shortcut == ""


class TestScrollableNameFallback:
    """Test scroll node name resolution: name -> automation_id -> localized_type."""

    def test_scroll_name_from_name(self, tree):
        child = _make_elem(bounding_rect=[0, 0, 200, 500])
        pane = _make_elem(
            name="MyScroller",
            control_type="Pane",
            localized_control_type="pane",
            automation_id="scroller1",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert scrollable[0].name == "MyScroller"

    def test_scroll_name_fallback_to_automation_id(self, tree):
        child = _make_elem(bounding_rect=[0, 0, 200, 500])
        pane = _make_elem(
            name="",
            control_type="Pane",
            localized_control_type="pane",
            automation_id="scroller1",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert scrollable[0].name == "scroller1"

    def test_scroll_name_fallback_to_localized_type(self, tree):
        child = _make_elem(bounding_rect=[0, 0, 200, 500])
        pane = _make_elem(
            name="",
            control_type="Pane",
            localized_control_type="pane",
            automation_id="",
            bounding_rect=[0, 0, 200, 200],
            children=[child],
        )
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert scrollable[0].name == "Pane"


class TestMixedInteractiveAndScrollable:
    """Test trees with both interactive elements and scrollable containers."""

    def test_tree_with_buttons_inside_scrollable_pane(self, tree):
        btn1 = _make_elem(name="Btn1", bounding_rect=[10, 10, 50, 50])
        btn2 = _make_elem(name="Btn2", bounding_rect=[10, 200, 50, 400])  # Overflows pane
        pane = _make_elem(
            name="ScrollPane",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 200, 200],
            children=[btn1, btn2],
        )
        interactive = []
        scrollable = []
        tree._classify_rust_tree(pane, "TestApp", [0, 0, 1920, 1080], interactive, scrollable)
        assert len(scrollable) == 1  # Pane is scrollable
        assert len(interactive) == 2  # Both buttons are interactive

    def test_multiple_scrollable_areas(self, tree):
        child1 = _make_elem(bounding_rect=[0, 0, 100, 500])
        child2 = _make_elem(bounding_rect=[200, 0, 300, 500])
        pane1 = _make_elem(
            name="Left",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 100, 200],
            children=[child1],
        )
        pane2 = _make_elem(
            name="Right",
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[200, 0, 300, 200],
            children=[child2],
        )
        root = _make_elem(
            control_type="Pane",
            localized_control_type="pane",
            bounding_rect=[0, 0, 400, 300],
            children=[pane1, pane2],
        )
        scrollable = []
        tree._classify_rust_tree(root, "TestApp", [0, 0, 1920, 1080], [], scrollable)
        assert len(scrollable) == 2
        names = {s.name for s in scrollable}
        assert names == {"Left", "Right"}
