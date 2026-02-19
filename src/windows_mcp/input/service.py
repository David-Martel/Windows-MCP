"""Mouse and keyboard input simulation.

Stateless service wrapping pyautogui and Windows UIA wheel events.
Uses Rust native extension (Win32 SendInput) as fast-path when available,
falling back to pyautogui for all operations.
"""

import logging
from typing import Literal

import pyautogui as pg  # noqa: E402

import windows_mcp.uia as uia  # noqa: E402
from windows_mcp.native import (
    native_send_click,
    native_send_mouse_move,
    native_send_text,
)

logger = logging.getLogger(__name__)


class InputService:
    """Simulate mouse clicks, keyboard input, scrolling, and drag operations."""

    def click(self, loc: tuple[int, int], button: str = "left", clicks: int = 1):
        """Click at screen coordinates.

        Args:
            loc: (x, y) screen coordinates.
            button: Mouse button -- "left", "right", or "middle".
            clicks: Number of clicks (1 for single, 2 for double).
        """
        x, y = loc
        if clicks == 1:
            result = native_send_click(x, y, button)
            if result is not None:
                if result == 0:
                    logger.warning(
                        "SendInput returned 0 events for click at (%d,%d) -- UIPI?", x, y
                    )
                return
        # Fall back to pyautogui for double-click and when native is unavailable
        pg.click(x, y, button=button, clicks=clicks, duration=0.1)

    def type(
        self,
        loc: tuple[int, int],
        text: str,
        caret_position: Literal["start", "idle", "end"] = "idle",
        clear: bool | str = False,
        press_enter: bool | str = False,
    ):
        """Click a field and type text into it.

        Args:
            loc: (x, y) screen coordinates of the target field.
            text: Text to type.
            caret_position: Where to position the caret before typing --
                "start" (Home key), "end" (End key), or "idle" (no movement).
            clear: If True or "true", select-all and delete existing text first.
            press_enter: If True or "true", press Enter after typing.
        """
        x, y = loc
        pg.leftClick(x, y)
        if caret_position == "start":
            pg.press("home")
        elif caret_position == "end":
            pg.press("end")
        # else "idle" -- no movement needed

        # Handle both boolean and string 'true'/'false'
        if clear is True or (isinstance(clear, str) and clear.lower() == "true"):
            pg.sleep(0.5)
            pg.hotkey("ctrl", "a")
            pg.press("backspace")

        # Rust fast-path: SendInput KEYEVENTF_UNICODE -- supports full Unicode,
        # sends all chars atomically in <1ms vs pyautogui's 20ms/char
        result = native_send_text(text)
        if result is None:
            pg.typewrite(text, interval=0.02)

        if press_enter is True or (isinstance(press_enter, str) and press_enter.lower() == "true"):
            pg.press("enter")

    def scroll(
        self,
        loc: tuple[int, int] | None = None,
        type: Literal["horizontal", "vertical"] = "vertical",
        direction: Literal["up", "down", "left", "right"] = "down",
        wheel_times: int = 1,
    ) -> str | None:
        """Scroll at a screen location.

        Args:
            loc: Optional (x, y) coordinates to move the cursor to before scrolling.
            type: Scroll axis -- "vertical" or "horizontal".
            direction: Scroll direction -- "up"/"down" for vertical, "left"/"right" for horizontal.
            wheel_times: Number of wheel notches to scroll.

        Returns:
            An error string if an invalid type/direction is supplied, else None.
        """
        if loc is not None:
            self.move(loc)
        match type:
            case "vertical":
                match direction:
                    case "up":
                        uia.WheelUp(wheel_times)
                    case "down":
                        uia.WheelDown(wheel_times)
                    case _:
                        return 'Invalid direction. Use "up" or "down".'
            case "horizontal":
                match direction:
                    case "left":
                        pg.keyDown("Shift")
                        try:
                            pg.sleep(0.05)
                            uia.WheelUp(wheel_times)
                            pg.sleep(0.05)
                        finally:
                            pg.keyUp("Shift")
                    case "right":
                        pg.keyDown("Shift")
                        try:
                            pg.sleep(0.05)
                            uia.WheelDown(wheel_times)
                            pg.sleep(0.05)
                        finally:
                            pg.keyUp("Shift")
                    case _:
                        return 'Invalid direction. Use "left" or "right".'
            case _:
                return 'Invalid type. Use "horizontal" or "vertical".'
        return None

    def drag(self, loc: tuple[int, int]):
        """Drag the cursor to screen coordinates (continues a drag already in progress).

        Args:
            loc: (x, y) destination coordinates.
        """
        x, y = loc
        pg.sleep(0.5)
        pg.dragTo(x, y, duration=0.6)

    def move(self, loc: tuple[int, int]):
        """Move the mouse cursor to screen coordinates without clicking.

        Args:
            loc: (x, y) destination coordinates.
        """
        x, y = loc
        result = native_send_mouse_move(x, y)
        if result is None:
            pg.moveTo(x, y, duration=0.1)

    def shortcut(self, shortcut: str):
        """Send a keyboard shortcut or single key press.

        Args:
            shortcut: Key name or ``+``-separated combination, e.g. ``"ctrl+c"`` or ``"enter"``.
        """
        keys = shortcut.split("+")
        if len(keys) > 1:
            pg.hotkey(*keys)
        else:
            pg.press("".join(keys))

    def multi_select(
        self, press_ctrl: bool | str = False, locs: list[tuple[int, int]] | None = None
    ):
        """Click multiple locations, optionally holding Ctrl for multi-selection.

        Args:
            press_ctrl: If True or "true", hold Ctrl while clicking each location.
            locs: List of (x, y) coordinate pairs to click.
        """
        if locs is None:
            locs = []
        hold_ctrl = press_ctrl is True or (
            isinstance(press_ctrl, str) and press_ctrl.lower() == "true"
        )
        if hold_ctrl:
            pg.keyDown("ctrl")
        try:
            for loc in locs:
                x, y = loc
                pg.click(x, y, duration=0.2)
                pg.sleep(0.5)
        finally:
            if hold_ctrl:
                pg.keyUp("ctrl")

    def multi_edit(self, locs: list[tuple[int, int, str]]):
        """Type text into multiple fields, clearing each before typing.

        Args:
            locs: List of (x, y, text) triples. Each field is clicked, cleared,
                then populated with the given text.
        """
        for loc in locs:
            x, y, text = loc
            self.type((x, y), text=text, clear=True)
