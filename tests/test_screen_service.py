"""Unit tests for ScreenService.

All UIA, pyautogui, ImageGrab, and ctypes calls are mocked so no live desktop
interaction occurs.  Covers happy paths, fallback paths, edge cases, and
dimension arithmetic for get_annotated_screenshot.
"""

from unittest.mock import MagicMock, patch

import pytest
from PIL import Image

from windows_mcp.desktop.views import Size
from windows_mcp.screen.service import ScreenService
from windows_mcp.tree.views import BoundingBox, Center, TreeElementNode

# ---------------------------------------------------------------------------
# Patch target constants
# ---------------------------------------------------------------------------

_UIA = "windows_mcp.screen.service.uia"
_PG = "windows_mcp.screen.service.pg"
_IMAGE_GRAB = "windows_mcp.screen.service.ImageGrab"
_CTYPES_WINDLL = "ctypes.windll.user32.GetDpiForSystem"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_rgb_image(width: int = 100, height: int = 80) -> Image.Image:
    """Return a real RGB PIL image of the given dimensions."""
    return Image.new("RGB", (width, height), color=(128, 128, 128))


def _make_node(left: int, top: int, right: int, bottom: int) -> TreeElementNode:
    """Return a minimal TreeElementNode with a bounding box."""
    box = BoundingBox(
        left=left,
        top=top,
        right=right,
        bottom=bottom,
        width=right - left,
        height=bottom - top,
    )
    center = Center(x=(left + right) // 2, y=(top + bottom) // 2)
    return TreeElementNode(
        bounding_box=box,
        center=center,
        name="Button",
        control_type="Button",
        window_name="TestApp",
    )


# ---------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------


@pytest.fixture
def svc() -> ScreenService:
    return ScreenService()


# ===========================================================================
# get_screen_size
# ===========================================================================


class TestGetScreenSize:
    """Tests for ScreenService.get_screen_size()."""

    def test_returns_size_dataclass(self, svc: ScreenService):
        """Return value must be a Size instance."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (1920, 1080)
            result = svc.get_screen_size()
        assert isinstance(result, Size)

    def test_width_and_height_populated_from_uia(self, svc: ScreenService):
        """Width and height fields match the tuple returned by GetVirtualScreenSize."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (1920, 1080)
            result = svc.get_screen_size()
        assert result.width == 1920
        assert result.height == 1080

    def test_non_standard_resolution(self, svc: ScreenService):
        """3440x1440 ultrawide resolution is returned verbatim."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (3440, 1440)
            result = svc.get_screen_size()
        assert result.width == 3440
        assert result.height == 1440

    def test_small_resolution(self, svc: ScreenService):
        """800x600 low-res display is returned without modification."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (800, 600)
            result = svc.get_screen_size()
        assert result.width == 800
        assert result.height == 600

    def test_multi_monitor_virtual_screen(self, svc: ScreenService):
        """A wide virtual screen spanning two monitors is returned correctly."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (3840, 1080)
            result = svc.get_screen_size()
        assert result.width == 3840
        assert result.height == 1080

    def test_uia_called_exactly_once(self, svc: ScreenService):
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (1280, 720)
            svc.get_screen_size()
        mock_uia.GetVirtualScreenSize.assert_called_once_with()

    def test_zero_dimensions_forwarded(self, svc: ScreenService):
        """Zero dimensions from UIA are forwarded without special-casing."""
        with patch(_UIA) as mock_uia:
            mock_uia.GetVirtualScreenSize.return_value = (0, 0)
            result = svc.get_screen_size()
        assert result.width == 0
        assert result.height == 0


# ===========================================================================
# get_dpi_scaling
# ===========================================================================


class TestGetDpiScaling:
    """Tests for ScreenService.get_dpi_scaling()."""

    def test_96_dpi_returns_1_0(self, svc: ScreenService):
        """96 DPI (standard) must produce exactly 1.0."""
        with patch(_CTYPES_WINDLL, return_value=96):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.0)

    def test_144_dpi_returns_1_5(self, svc: ScreenService):
        """144 DPI (150 % scaling) must produce 1.5."""
        with patch(_CTYPES_WINDLL, return_value=144):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.5)

    def test_192_dpi_returns_2_0(self, svc: ScreenService):
        """192 DPI (200 % scaling) must produce 2.0."""
        with patch(_CTYPES_WINDLL, return_value=192):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(2.0)

    def test_120_dpi_returns_1_25(self, svc: ScreenService):
        """120 DPI (125 % scaling) must produce 1.25."""
        with patch(_CTYPES_WINDLL, return_value=120):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.25)

    def test_zero_dpi_returns_1_0(self, svc: ScreenService):
        """When GetDpiForSystem returns 0 the guard (dpi > 0) must produce 1.0."""
        with patch(_CTYPES_WINDLL, return_value=0):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.0)

    def test_negative_dpi_returns_1_0(self, svc: ScreenService):
        """Negative DPI (abnormal return value) must trigger the guard and return 1.0."""
        with patch(_CTYPES_WINDLL, return_value=-1):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.0)

    def test_exception_from_windll_returns_1_0(self, svc: ScreenService):
        """Any exception raised by GetDpiForSystem must be swallowed and return 1.0."""
        with patch(_CTYPES_WINDLL, side_effect=OSError("no windll")):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.0)

    def test_attribute_error_returns_1_0(self, svc: ScreenService):
        """AttributeError (e.g. windll absent) must also be swallowed."""
        with patch(_CTYPES_WINDLL, side_effect=AttributeError("windll missing")):
            result = svc.get_dpi_scaling()
        assert result == pytest.approx(1.0)

    def test_returns_float(self, svc: ScreenService):
        """Return type must always be float."""
        with patch(_CTYPES_WINDLL, return_value=96):
            result = svc.get_dpi_scaling()
        assert isinstance(result, float)

    def test_returns_float_on_exception(self, svc: ScreenService):
        """Return type must be float even when the exception path is taken."""
        with patch(_CTYPES_WINDLL, side_effect=Exception("boom")):
            result = svc.get_dpi_scaling()
        assert isinstance(result, float)


# ===========================================================================
# get_cursor_location
# ===========================================================================


class TestGetCursorLocation:
    """Tests for ScreenService.get_cursor_location()."""

    def test_returns_tuple(self, svc: ScreenService):
        pos = MagicMock()
        pos.x, pos.y = 300, 400
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert isinstance(result, tuple)

    def test_returns_x_y_from_pg_position(self, svc: ScreenService):
        pos = MagicMock()
        pos.x, pos.y = 300, 400
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert result == (300, 400)

    def test_origin_coordinates(self, svc: ScreenService):
        """Cursor at (0, 0) must be returned as-is."""
        pos = MagicMock()
        pos.x, pos.y = 0, 0
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert result == (0, 0)

    def test_large_coordinates(self, svc: ScreenService):
        """Very large coordinates (multi-monitor far-right) are forwarded unchanged."""
        pos = MagicMock()
        pos.x, pos.y = 7680, 4320
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert result == (7680, 4320)

    def test_negative_coordinates(self, svc: ScreenService):
        """Negative coordinates (left of primary monitor) are forwarded unchanged."""
        pos = MagicMock()
        pos.x, pos.y = -100, -50
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert result == (-100, -50)

    def test_pg_position_called_once(self, svc: ScreenService):
        pos = MagicMock()
        pos.x, pos.y = 10, 20
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            svc.get_cursor_location()
        mock_pg.position.assert_called_once_with()

    def test_tuple_length_is_two(self, svc: ScreenService):
        pos = MagicMock()
        pos.x, pos.y = 50, 60
        with patch(_PG) as mock_pg:
            mock_pg.position.return_value = pos
            result = svc.get_cursor_location()
        assert len(result) == 2


# ===========================================================================
# get_element_under_cursor
# ===========================================================================


class TestGetElementUnderCursor:
    """Tests for ScreenService.get_element_under_cursor()."""

    def test_returns_uia_control_from_cursor(self, svc: ScreenService):
        """Return value must be exactly what ControlFromCursor() returns."""
        mock_control = MagicMock(name="UIAControl")
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromCursor.return_value = mock_control
            result = svc.get_element_under_cursor()
        assert result is mock_control

    def test_control_from_cursor_called_once(self, svc: ScreenService):
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromCursor.return_value = MagicMock()
            svc.get_element_under_cursor()
        mock_uia.ControlFromCursor.assert_called_once_with()

    def test_returns_none_when_no_control_under_cursor(self, svc: ScreenService):
        """ControlFromCursor returning None must be forwarded, not converted."""
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromCursor.return_value = None
            result = svc.get_element_under_cursor()
        assert result is None

    def test_different_control_objects_are_forwarded(self, svc: ScreenService):
        """Two successive calls each forward the distinct value from ControlFromCursor."""
        ctrl_a = MagicMock(name="ControlA")
        ctrl_b = MagicMock(name="ControlB")
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromCursor.side_effect = [ctrl_a, ctrl_b]
            assert svc.get_element_under_cursor() is ctrl_a
            assert svc.get_element_under_cursor() is ctrl_b


# ===========================================================================
# get_screenshot
# ===========================================================================


class TestGetScreenshot:
    """Tests for ScreenService.get_screenshot()."""

    def test_normal_path_uses_imagegrab(self, svc: ScreenService):
        """When ImageGrab.grab succeeds its return value is passed through."""
        fake_img = _make_rgb_image(1920, 1080)
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG):
            mock_grab_mod.grab.return_value = fake_img
            result = svc.get_screenshot()
        assert result is fake_img

    def test_imagegrab_called_with_all_screens_true(self, svc: ScreenService):
        """ImageGrab.grab must be invoked with all_screens=True."""
        fake_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG):
            mock_grab_mod.grab.return_value = fake_img
            svc.get_screenshot()
        mock_grab_mod.grab.assert_called_once_with(all_screens=True)

    def test_fallback_to_pg_screenshot_on_oserror(self, svc: ScreenService):
        """When ImageGrab.grab raises OSError the pyautogui fallback must be used."""
        fallback_img = _make_rgb_image(1280, 720)
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG) as mock_pg:
            mock_grab_mod.grab.side_effect = OSError("grab failed")
            mock_pg.screenshot.return_value = fallback_img
            result = svc.get_screenshot()
        assert result is fallback_img

    def test_fallback_to_pg_screenshot_on_exception(self, svc: ScreenService):
        """Any exception (not just OSError) from ImageGrab triggers the fallback."""
        fallback_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG) as mock_pg:
            mock_grab_mod.grab.side_effect = RuntimeError("unexpected")
            mock_pg.screenshot.return_value = fallback_img
            result = svc.get_screenshot()
        assert result is fallback_img

    def test_fallback_calls_pg_screenshot_no_args(self, svc: ScreenService):
        """The pyautogui fallback calls screenshot() with no positional arguments."""
        fallback_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG) as mock_pg:
            mock_grab_mod.grab.side_effect = OSError("fail")
            mock_pg.screenshot.return_value = fallback_img
            svc.get_screenshot()
        mock_pg.screenshot.assert_called_once_with()

    def test_pg_screenshot_not_called_on_success(self, svc: ScreenService):
        """The pyautogui fallback must not be invoked when ImageGrab succeeds."""
        fake_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG) as mock_pg:
            mock_grab_mod.grab.return_value = fake_img
            svc.get_screenshot()
        mock_pg.screenshot.assert_not_called()

    def test_imagegrab_not_called_after_single_failure(self, svc: ScreenService):
        """Only one ImageGrab.grab attempt is made; no retry loop."""
        fallback_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG) as mock_pg:
            mock_grab_mod.grab.side_effect = OSError("fail")
            mock_pg.screenshot.return_value = fallback_img
            svc.get_screenshot()
        assert mock_grab_mod.grab.call_count == 1

    def test_returns_pil_image(self, svc: ScreenService):
        """Return type is always Image.Image regardless of which path was taken."""
        fake_img = _make_rgb_image()
        with patch(_IMAGE_GRAB) as mock_grab_mod, patch(_PG):
            mock_grab_mod.grab.return_value = fake_img
            result = svc.get_screenshot()
        assert isinstance(result, Image.Image)


# ===========================================================================
# get_annotated_screenshot
# ===========================================================================


class TestGetAnnotatedScreenshot:
    """Tests for ScreenService.get_annotated_screenshot()."""

    # -----------------------------------------------------------------------
    # Padding / dimension arithmetic
    # -----------------------------------------------------------------------

    def test_output_width_includes_padding(self, svc: ScreenService):
        """Output image width must equal screenshot width + 2*padding (padding=5)."""
        base_img = _make_rgb_image(width=200, height=100)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 200, 100)
            result = svc.get_annotated_screenshot([])
        assert result.width == 210  # 200 + 2*5

    def test_output_height_includes_padding(self, svc: ScreenService):
        """Output image height must equal screenshot height + 2*padding (padding=5)."""
        base_img = _make_rgb_image(width=200, height=100)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 200, 100)
            result = svc.get_annotated_screenshot([])
        assert result.height == 110  # 100 + 2*5

    def test_output_mode_is_rgb(self, svc: ScreenService):
        """The padded output image must use RGB mode."""
        base_img = _make_rgb_image(100, 80)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 100, 80)
            result = svc.get_annotated_screenshot([])
        assert result.mode == "RGB"

    def test_empty_nodes_list_still_returns_image(self, svc: ScreenService):
        """An empty node list must produce a valid padded image without error."""
        base_img = _make_rgb_image(320, 240)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 320, 240)
            result = svc.get_annotated_screenshot([])
        assert result.width == 330
        assert result.height == 250

    def test_single_node_does_not_raise(self, svc: ScreenService):
        """A single TreeElementNode must be annotated without error."""
        base_img = _make_rgb_image(400, 300)
        node = _make_node(10, 20, 110, 70)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 400, 300)
            result = svc.get_annotated_screenshot([node])
        assert isinstance(result, Image.Image)

    def test_multiple_nodes_do_not_raise(self, svc: ScreenService):
        """Multiple nodes must all be annotated without error."""
        base_img = _make_rgb_image(800, 600)
        nodes = [
            _make_node(10, 10, 100, 50),
            _make_node(200, 100, 400, 200),
            _make_node(500, 400, 700, 580),
        ]
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 800, 600)
            result = svc.get_annotated_screenshot(nodes)
        assert isinstance(result, Image.Image)

    # -----------------------------------------------------------------------
    # UIA rect usage
    # -----------------------------------------------------------------------

    def test_get_virtual_screen_rect_called_once(self, svc: ScreenService):
        """GetVirtualScreenRect must be called exactly once per annotated screenshot."""
        base_img = _make_rgb_image(100, 100)
        node = _make_node(0, 0, 50, 50)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 100, 100)
            svc.get_annotated_screenshot([node])
        mock_uia.GetVirtualScreenRect.assert_called_once_with()

    def test_non_zero_screen_offset_does_not_raise(self, svc: ScreenService):
        """A negative left/top offset (secondary monitor left of primary) must not raise."""
        base_img = _make_rgb_image(400, 300)
        node = _make_node(50, 50, 200, 150)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            # Simulate a monitor layout where the virtual origin is at (-1920, 0)
            mock_uia.GetVirtualScreenRect.return_value = (-1920, 0, -1520, 300)
            result = svc.get_annotated_screenshot([node])
        assert isinstance(result, Image.Image)

    # -----------------------------------------------------------------------
    # Font loading fallback
    # -----------------------------------------------------------------------

    def test_ioerror_font_does_not_raise(self, svc: ScreenService):
        """When arial.ttf raises IOError the except branch calls load_default.

        We only intercept the arial.ttf path by checking the first argument inside
        a side_effect function; all other truetype calls (e.g. Pillow's own
        load_default internals) pass through to the real implementation so that
        PIL.ImageDraw.text() receives a functional font object.
        """
        import PIL.ImageFont as _real_font

        base_img = _make_rgb_image(200, 150)
        node = _make_node(10, 10, 100, 80)

        _original_truetype = _real_font.truetype

        def fake_truetype(path_or_stream, size=None, **kwargs):
            if path_or_stream == "arial.ttf":
                raise IOError("arial.ttf not found")
            return _original_truetype(path_or_stream, size, **kwargs)

        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 200, 150)
            with patch("windows_mcp.screen.service.ImageFont.truetype", side_effect=fake_truetype):
                result = svc.get_annotated_screenshot([node])
        assert isinstance(result, Image.Image)

    # -----------------------------------------------------------------------
    # Screenshot delegation
    # -----------------------------------------------------------------------

    def test_annotated_screenshot_calls_get_screenshot(self, svc: ScreenService):
        """get_annotated_screenshot must delegate to self.get_screenshot()."""
        base_img = _make_rgb_image(200, 150)
        with patch.object(svc, "get_screenshot", return_value=base_img) as mock_ss:
            with patch(_UIA) as mock_uia:
                mock_uia.GetVirtualScreenRect.return_value = (0, 0, 200, 150)
                svc.get_annotated_screenshot([])
        mock_ss.assert_called_once_with()

    def test_annotated_screenshot_uses_imagegrab_fallback(self, svc: ScreenService):
        """When ImageGrab raises, the annotated path must still succeed via fallback."""
        fallback_img = _make_rgb_image(640, 480)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_PG) as mock_pg, patch(_UIA) as mock_uia:
            mock_grab.grab.side_effect = OSError("grab unavailable")
            mock_pg.screenshot.return_value = fallback_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 640, 480)
            result = svc.get_annotated_screenshot([])
        assert result.width == 650
        assert result.height == 490

    # -----------------------------------------------------------------------
    # Dimension correctness with various base sizes
    # -----------------------------------------------------------------------

    @pytest.mark.parametrize(
        "w, h",
        [
            (1920, 1080),
            (1280, 720),
            (3840, 2160),
            (800, 600),
            (1, 1),
        ],
    )
    def test_padded_dimensions_for_various_resolutions(self, svc: ScreenService, w: int, h: int):
        """Padding of 5 px is always applied on each side for any resolution."""
        base_img = _make_rgb_image(w, h)
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, w, h)
            result = svc.get_annotated_screenshot([])
        assert result.width == w + 10
        assert result.height == h + 10

    # -----------------------------------------------------------------------
    # Node count does not affect output image dimensions
    # -----------------------------------------------------------------------

    def test_many_nodes_output_dimensions_unchanged(self, svc: ScreenService):
        """Annotating 20 nodes must not alter the padded output image dimensions."""
        base_img = _make_rgb_image(640, 480)
        nodes = [_make_node(i * 20, i * 10, i * 20 + 15, i * 10 + 10) for i in range(20)]
        with patch(_IMAGE_GRAB) as mock_grab, patch(_UIA) as mock_uia, patch(_PG):
            mock_grab.grab.return_value = base_img
            mock_uia.GetVirtualScreenRect.return_value = (0, 0, 640, 480)
            result = svc.get_annotated_screenshot(nodes)
        assert result.width == 650
        assert result.height == 490
