"""Unit tests for tree utility functions and pure methods.

Tests app_name_correction (module-level), random_point_within_bounding_box,
and Tree.iou_bounding_box (pure math, mocked screen_box).
"""

from unittest.mock import MagicMock

from windows_mcp.tree.utils import app_name_correction, random_point_within_bounding_box
from windows_mcp.tree.views import BoundingBox

# ---------------------------------------------------------------------------
# app_name_correction -- all branches
# ---------------------------------------------------------------------------


class TestAppNameCorrection:
    """Pure function: maps internal class names to user-friendly display names."""

    def test_progman_becomes_desktop(self):
        assert app_name_correction("Progman") == "Desktop"

    def test_shell_traywnd_becomes_taskbar(self):
        assert app_name_correction("Shell_TrayWnd") == "Taskbar"

    def test_shell_secondary_traywnd_becomes_taskbar(self):
        assert app_name_correction("Shell_SecondaryTrayWnd") == "Taskbar"

    def test_popup_window_site_bridge_becomes_context_menu(self):
        assert app_name_correction("Microsoft.UI.Content.PopupWindowSiteBridge") == "Context Menu"

    def test_regular_name_passes_through(self):
        assert app_name_correction("Notepad") == "Notepad"

    def test_empty_string_passes_through(self):
        assert app_name_correction("") == ""

    def test_case_sensitive_no_match(self):
        # "progman" (lowercase) should NOT match "Progman"
        assert app_name_correction("progman") == "progman"

    def test_partial_match_does_not_correct(self):
        assert app_name_correction("Shell_TrayWnd_Extra") == "Shell_TrayWnd_Extra"


# ---------------------------------------------------------------------------
# random_point_within_bounding_box
# ---------------------------------------------------------------------------


class TestRandomPointWithinBoundingBox:
    """Verifies random point generation stays within bounds."""

    def _make_node(self, left=100, top=50, right=300, bottom=200):
        node = MagicMock()
        box = MagicMock()
        box.left = left
        box.top = top
        box.right = right
        box.bottom = bottom
        box.width.return_value = right - left
        box.height.return_value = bottom - top
        node.BoundingRectangle = box
        return node

    def test_point_within_bounds_default_scale(self):
        node = self._make_node(100, 50, 300, 200)
        for _ in range(50):
            x, y = random_point_within_bounding_box(node)
            assert 100 <= x <= 300
            assert 50 <= y <= 200

    def test_point_with_smaller_scale(self):
        node = self._make_node(0, 0, 100, 100)
        for _ in range(50):
            x, y = random_point_within_bounding_box(node, scale_factor=0.5)
            # Scaled to 50x50 centered in 100x100 -> range [25..75]
            assert 25 <= x <= 75
            assert 25 <= y <= 75

    def test_single_pixel_box(self):
        node = self._make_node(10, 20, 10, 20)
        x, y = random_point_within_bounding_box(node)
        assert x == 10
        assert y == 20


# ---------------------------------------------------------------------------
# Tree.iou_bounding_box -- intersection + screen clamping
# ---------------------------------------------------------------------------


class TestIouBoundingBox:
    """Tests the bounding box intersection-of-union with screen clamping."""

    def _make_tree(self, screen_left=0, screen_top=0, screen_right=1920, screen_bottom=1080):
        """Create a minimal Tree-like object with just screen_box set."""
        from windows_mcp.tree.service import Tree

        tree = Tree.__new__(Tree)
        tree.screen_box = BoundingBox(
            left=screen_left,
            top=screen_top,
            right=screen_right,
            bottom=screen_bottom,
            width=screen_right - screen_left,
            height=screen_bottom - screen_top,
        )
        return tree

    def _make_rect(self, left, top, right, bottom):
        rect = MagicMock()
        rect.left = left
        rect.top = top
        rect.right = right
        rect.bottom = bottom
        return rect

    def test_full_overlap(self):
        tree = self._make_tree()
        window = self._make_rect(0, 0, 800, 600)
        element = self._make_rect(100, 100, 400, 300)
        bb = tree.iou_bounding_box(window, element)
        assert bb.left == 100
        assert bb.top == 100
        assert bb.right == 400
        assert bb.bottom == 300
        assert bb.width == 300
        assert bb.height == 200

    def test_partial_overlap(self):
        tree = self._make_tree()
        window = self._make_rect(0, 0, 200, 200)
        element = self._make_rect(100, 100, 400, 400)
        bb = tree.iou_bounding_box(window, element)
        assert bb.left == 100
        assert bb.top == 100
        assert bb.right == 200
        assert bb.bottom == 200

    def test_no_overlap_returns_zero_box(self):
        tree = self._make_tree()
        window = self._make_rect(0, 0, 100, 100)
        element = self._make_rect(200, 200, 300, 300)
        bb = tree.iou_bounding_box(window, element)
        assert bb.left == 0
        assert bb.width == 0
        assert bb.height == 0

    def test_screen_clamping(self):
        """Element extends beyond screen -- result is clamped."""
        tree = self._make_tree(0, 0, 1920, 1080)
        window = self._make_rect(-100, -50, 2000, 1200)
        element = self._make_rect(-50, -30, 2500, 1500)
        bb = tree.iou_bounding_box(window, element)
        # Intersection of window and element is (-50, -30, 2000, 1200)
        # Clamped to screen: (0, 0, 1920, 1080)
        assert bb.left == 0
        assert bb.top == 0
        assert bb.right == 1920
        assert bb.bottom == 1080

    def test_element_fully_outside_screen(self):
        tree = self._make_tree(0, 0, 1920, 1080)
        window = self._make_rect(0, 0, 1920, 1080)
        element = self._make_rect(2000, 2000, 2100, 2100)
        bb = tree.iou_bounding_box(window, element)
        assert bb.width == 0
        assert bb.height == 0

    def test_touching_edges_returns_zero_box(self):
        """Adjacent but non-overlapping boxes: right==left edge."""
        tree = self._make_tree()
        window = self._make_rect(0, 0, 100, 100)
        element = self._make_rect(100, 0, 200, 100)
        bb = tree.iou_bounding_box(window, element)
        # right == left means no overlap (> not >=)
        assert bb.width == 0

    def test_small_screen(self):
        """Screen smaller than window/element."""
        tree = self._make_tree(0, 0, 100, 100)
        window = self._make_rect(0, 0, 500, 500)
        element = self._make_rect(0, 0, 500, 500)
        bb = tree.iou_bounding_box(window, element)
        assert bb.right == 100
        assert bb.bottom == 100
        assert bb.width == 100
        assert bb.height == 100
