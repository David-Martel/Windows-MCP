import random

from windows_mcp.uia import Control


def app_name_correction(app_name: str) -> str:
    """Normalize internal window class names to user-friendly display names."""
    match app_name:
        case "Progman":
            return "Desktop"
        case "Shell_TrayWnd" | "Shell_SecondaryTrayWnd":
            return "Taskbar"
        case "Microsoft.UI.Content.PopupWindowSiteBridge":
            return "Context Menu"
        case _:
            return app_name


def random_point_within_bounding_box(node: Control, scale_factor: float = 1.0) -> tuple[int, int]:
    """
    Generate a random point within a scaled-down bounding box.

    Args:
        node (Control): The node with a bounding rectangle
        scale_factor (float, optional): The factor to scale down the bounding box. Defaults to 1.0.

    Returns:
        tuple: A random point (x, y) within the scaled-down bounding box
    """
    box = node.BoundingRectangle
    scaled_width = int(box.width() * scale_factor)
    scaled_height = int(box.height() * scale_factor)
    scaled_left = box.left + (box.width() - scaled_width) // 2
    scaled_top = box.top + (box.height() - scaled_height) // 2
    x = random.randint(scaled_left, scaled_left + scaled_width)
    y = random.randint(scaled_top, scaled_top + scaled_height)
    return (x, y)
