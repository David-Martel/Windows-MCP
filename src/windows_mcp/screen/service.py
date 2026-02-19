"""Screen capture, annotation, and display metrics.

Stateless service providing screenshot capture, bounding-box annotation,
screen size queries, DPI scaling, and cursor operations.
"""

import io
import logging
import random

import pyautogui as pg
from PIL import Image, ImageDraw, ImageFont, ImageGrab

import windows_mcp.uia as uia
from windows_mcp.desktop.views import Size
from windows_mcp.native import native_capture_screenshot_png
from windows_mcp.tree.views import TreeElementNode

logger = logging.getLogger(__name__)


class ScreenService:
    """Screen capture, annotation, and display metrics."""

    def get_screen_size(self) -> Size:
        """Return virtual screen dimensions."""
        width, height = uia.GetVirtualScreenSize()
        return Size(width=width, height=height)

    def get_dpi_scaling(self) -> float:
        """Return DPI scaling factor (1.0 = 96 DPI)."""
        import ctypes

        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
            return dpi / 96.0 if dpi > 0 else 1.0
        except Exception:
            return 1.0

    def get_cursor_location(self) -> tuple[int, int]:
        """Return current cursor (x, y) screen coordinates."""
        position = pg.position()
        return (position.x, position.y)

    def get_element_under_cursor(self) -> uia.Control:
        """Return the UIA control element under the cursor."""
        return uia.ControlFromCursor()

    def get_screenshot(self) -> Image.Image:
        """Capture the full virtual screen as a PIL Image.

        Tries Rust DXGI/GDI capture first (10-50ms), falls back to
        PIL ImageGrab (100-200ms), then pyautogui as last resort.
        """
        # Fast path: Rust DXGI/GDI capture
        png_bytes = native_capture_screenshot_png(0)
        if png_bytes is not None:
            try:
                return Image.open(io.BytesIO(png_bytes))
            except Exception:
                logger.warning("Failed to decode Rust screenshot PNG, falling back")

        # Fallback: PIL ImageGrab
        try:
            return ImageGrab.grab(all_screens=True)
        except Exception:
            logger.warning("Failed to capture virtual screen, using primary screen")
            return pg.screenshot()

    def get_annotated_screenshot(self, nodes: list[TreeElementNode]) -> Image.Image:
        """Capture screenshot with numbered bounding-box annotations on interactive elements."""
        screenshot = self.get_screenshot()
        # Add padding
        padding = 5
        width = screenshot.width + 2 * padding
        height = screenshot.height + 2 * padding
        padded_screenshot = Image.new("RGB", (width, height), color=(255, 255, 255))
        padded_screenshot.paste(screenshot, (padding, padding))

        draw = ImageDraw.Draw(padded_screenshot)
        font_size = 12
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except IOError:
            font = ImageFont.load_default()

        def get_random_color():
            return "#{:06x}".format(random.randint(0, 0xFFFFFF))

        left_offset, top_offset, _, _ = uia.GetVirtualScreenRect()

        def draw_annotation(label, node: TreeElementNode):
            box = node.bounding_box
            color = get_random_color()

            # Adjust for virtual screen offset so coordinates map to the screenshot image
            adjusted_box = (
                int(box.left - left_offset) + padding,
                int(box.top - top_offset) + padding,
                int(box.right - left_offset) + padding,
                int(box.bottom - top_offset) + padding,
            )
            # Draw bounding box
            draw.rectangle(adjusted_box, outline=color, width=2)

            # Label dimensions
            label_width = draw.textlength(str(label), font=font)
            label_height = font_size
            left, top, right, bottom = adjusted_box

            # Label position above bounding box
            label_x1 = right - label_width
            label_y1 = top - label_height - 4
            label_x2 = label_x1 + label_width
            label_y2 = label_y1 + label_height + 4

            # Draw label background and text
            draw.rectangle([(label_x1, label_y1), (label_x2, label_y2)], fill=color)
            draw.text(
                (label_x1 + 2, label_y1 + 2),
                str(label),
                fill=(255, 255, 255),
                font=font,
            )

        for label, node in enumerate(nodes):
            draw_annotation(label, node)
        return padded_screenshot
