"""Unit tests for Desktop XPath resolution methods.

Tests get_xpath_from_element and get_element_from_xpath with mocked UIA controls.
All UIA interactions are mocked so the suite runs headless.
"""

from unittest.mock import MagicMock, patch

import pytest

from tests.desktop_helpers import make_bare_desktop

_UIA = "windows_mcp.desktop.service.uia"


# ---------------------------------------------------------------------------
# Helpers for building mock UIA control trees
# ---------------------------------------------------------------------------


def _make_control(
    name="Ctrl",
    control_type=50000,
    control_type_name="Button",
    runtime_id=None,
    parent=None,
    children=None,
):
    """Create a mock UIA control with the given properties."""
    ctrl = MagicMock()
    ctrl.Name = name
    ctrl.ControlType = control_type
    ctrl.ControlTypeName = control_type_name
    ctrl.GetRuntimeId.return_value = runtime_id or [42, 1]
    ctrl.GetParentControl.return_value = parent
    ctrl.GetChildren.return_value = children or []
    return ctrl


# ===========================================================================
# get_xpath_from_element
# ===========================================================================


class TestGetXpathFromElement:
    """Tests for building an XPath string from a UIA element."""

    def test_returns_empty_string_for_none_element(self):
        d = make_bare_desktop()
        assert d.get_xpath_from_element(None) == ""

    def test_root_element_returns_control_type_name(self):
        """An element with no parent is the root node."""
        d = make_bare_desktop()
        root = _make_control(control_type_name="Pane", parent=None)
        result = d.get_xpath_from_element(root)
        assert result == "Pane"

    def test_single_child_element(self):
        """Parent -> Child produces 'ParentType/ChildType[1]'."""
        d = make_bare_desktop()
        parent = _make_control(control_type_name="Window", parent=None, runtime_id=[42, 1])
        child = _make_control(
            control_type_name="Button", control_type=50000, parent=parent, runtime_id=[42, 2]
        )
        # Parent has one child of type Button
        parent.GetChildren.return_value = [child]
        result = d.get_xpath_from_element(child)
        assert result == "Window/Button[1]"

    def test_multi_level_path(self):
        """Root -> Pane -> Button produces correct path."""
        d = make_bare_desktop()
        root = _make_control(control_type_name="Pane", parent=None, runtime_id=[42, 0])
        pane = _make_control(
            control_type_name="Pane", control_type=50033, parent=root, runtime_id=[42, 1]
        )
        button = _make_control(
            control_type_name="Button", control_type=50000, parent=pane, runtime_id=[42, 2]
        )
        root.GetChildren.return_value = [pane]
        pane.GetChildren.return_value = [button]

        result = d.get_xpath_from_element(button)
        assert result == "Pane/Pane[1]/Button[1]"

    def test_sibling_indexing(self):
        """Second sibling of same type gets index [2]."""
        d = make_bare_desktop()
        parent = _make_control(control_type_name="Window", parent=None, runtime_id=[42, 0])
        btn1 = _make_control(
            control_type_name="Button", control_type=50000, parent=parent, runtime_id=[42, 1]
        )
        btn2 = _make_control(
            control_type_name="Button", control_type=50000, parent=parent, runtime_id=[42, 2]
        )
        parent.GetChildren.return_value = [btn1, btn2]

        result = d.get_xpath_from_element(btn2)
        assert result == "Window/Button[2]"

    def test_runtime_id_not_found_defaults_to_index_zero(self):
        """When RuntimeId doesn't match any sibling, index defaults to 0 -> [1]."""
        d = make_bare_desktop()
        parent = _make_control(control_type_name="Window", parent=None, runtime_id=[42, 0])
        child = _make_control(
            control_type_name="Edit", control_type=50004, parent=parent, runtime_id=[42, 99]
        )
        # Sibling has a different RuntimeId
        sibling = _make_control(
            control_type_name="Edit", control_type=50004, parent=parent, runtime_id=[42, 1]
        )
        parent.GetChildren.return_value = [sibling]

        result = d.get_xpath_from_element(child)
        # child's RuntimeId [42, 99] -> "42-99" not in ["42-1"] -> ValueError -> index=0 -> [1]
        assert "Edit[1]" in result


# ===========================================================================
# get_element_from_xpath
# ===========================================================================


class TestGetElementFromXpath:
    """Tests for resolving a UIA element from an XPath string."""

    def test_single_level_xpath(self):
        d = make_bare_desktop()
        btn = _make_control(control_type_name="Button")
        root = _make_control(control_type_name="Pane", children=[btn])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            result = d.get_element_from_xpath("Pane/Button[1]")

        assert result is btn

    def test_xpath_without_index_selects_first(self):
        """When no [index] is specified, first matching child is selected."""
        d = make_bare_desktop()
        btn1 = _make_control(control_type_name="Button")
        btn2 = _make_control(control_type_name="Button")
        root = _make_control(control_type_name="Pane", children=[btn1, btn2])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            result = d.get_element_from_xpath("Pane/Button")

        assert result is btn1

    def test_xpath_with_index_selects_correct_sibling(self):
        d = make_bare_desktop()
        btn1 = _make_control(control_type_name="Button")
        btn2 = _make_control(control_type_name="Button")
        root = _make_control(control_type_name="Pane", children=[btn1, btn2])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            result = d.get_element_from_xpath("Pane/Button[2]")

        assert result is btn2

    def test_multi_level_resolution(self):
        d = make_bare_desktop()
        edit = _make_control(control_type_name="Edit")
        pane = _make_control(control_type_name="Pane", children=[edit])
        root = _make_control(control_type_name="Window", children=[pane])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            result = d.get_element_from_xpath("Window/Pane[1]/Edit[1]")

        assert result is edit

    def test_raises_for_missing_control_type(self):
        d = make_bare_desktop()
        root = _make_control(control_type_name="Pane", children=[])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            with pytest.raises(ValueError, match="no children of type"):
                d.get_element_from_xpath("Pane/Button[1]")

    def test_raises_for_index_out_of_range(self):
        d = make_bare_desktop()
        btn = _make_control(control_type_name="Button")
        root = _make_control(control_type_name="Pane", children=[btn])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            with pytest.raises(ValueError, match="index 5 out of range"):
                d.get_element_from_xpath("Pane/Button[5]")

    def test_raises_for_index_zero(self):
        """Index 0 is invalid (1-based indexing)."""
        d = make_bare_desktop()
        btn = _make_control(control_type_name="Button")
        root = _make_control(control_type_name="Pane", children=[btn])

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            with pytest.raises(ValueError, match="index 0 out of range"):
                d.get_element_from_xpath("Pane/Button[0]")

    def test_skips_non_matching_xpath_parts(self):
        """Parts that don't match the regex pattern are skipped."""
        d = make_bare_desktop()
        root = _make_control(control_type_name="Pane")

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            # "!!!" won't match \w+ pattern, so it gets skipped
            result = d.get_element_from_xpath("Pane/!!!")

        assert result is root

    def test_returns_root_for_root_only_xpath(self):
        d = make_bare_desktop()
        root = _make_control(control_type_name="Pane")

        with patch(f"{_UIA}.GetRootControl", return_value=root):
            result = d.get_element_from_xpath("Pane")

        assert result is root
