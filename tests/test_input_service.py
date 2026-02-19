"""Edge case and fuzz-style unit tests for InputService.

All pyautogui calls, native send functions, and uia wheel functions are mocked
so no real input events are generated during the test run.
"""

import logging
from unittest.mock import MagicMock, call, patch

import pytest

from windows_mcp.input.service import InputService

# ---------------------------------------------------------------------------
# Shared patch targets
# ---------------------------------------------------------------------------
_PG = "windows_mcp.input.service.pg"
_NATIVE_CLICK = "windows_mcp.input.service.native_send_click"
_NATIVE_TEXT = "windows_mcp.input.service.native_send_text"
_NATIVE_MOVE = "windows_mcp.input.service.native_send_mouse_move"
_NATIVE_SCROLL = "windows_mcp.input.service.native_send_scroll"
_NATIVE_DRAG = "windows_mcp.input.service.native_send_drag"
_UIA = "windows_mcp.input.service.uia"


# ---------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------


@pytest.fixture
def svc():
    return InputService()


# ---------------------------------------------------------------------------
# Helper: build a fully-patched context for tests that need all mocks at once
# ---------------------------------------------------------------------------


def _all_mocks(pg_kw=None, click_rv=1, text_rv=1, move_rv=1, scroll_rv=2, drag_rv=3):
    """Return a nested patch context that replaces all external symbols."""
    pg_defaults = dict(
        leftClick=MagicMock(),
        click=MagicMock(),
        press=MagicMock(),
        hotkey=MagicMock(),
        sleep=MagicMock(),
        typewrite=MagicMock(),
        keyDown=MagicMock(),
        keyUp=MagicMock(),
        moveTo=MagicMock(),
        dragTo=MagicMock(),
        position=MagicMock(return_value=(500, 500)),
    )
    if pg_kw:
        pg_defaults.update(pg_kw)
    return (
        patch(_PG, **{k: v for k, v in pg_defaults.items()}),
        patch(_NATIVE_CLICK, return_value=click_rv),
        patch(_NATIVE_TEXT, return_value=text_rv),
        patch(_NATIVE_MOVE, return_value=move_rv),
        patch(_NATIVE_SCROLL, return_value=scroll_rv),
        patch(_NATIVE_DRAG, return_value=drag_rv),
        patch(_UIA),
    )


# ===========================================================================
# click()
# ===========================================================================


class TestClick:
    """Tests for InputService.click()."""

    def test_single_click_uses_native_path(self, svc):
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG) as mock_pg:
            svc.click((100, 200))
            mock_click.assert_called_once_with(100, 200, "left")
            mock_pg.click.assert_not_called()

    def test_double_click_falls_back_to_pyautogui(self, svc):
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG) as mock_pg:
            svc.click((100, 200), clicks=2)
            mock_click.assert_not_called()
            mock_pg.click.assert_called_once_with(100, 200, button="left", clicks=2, duration=0.1)

    def test_triple_click_falls_back_to_pyautogui(self, svc):
        with patch(_NATIVE_CLICK, return_value=0) as mock_click, patch(_PG) as mock_pg:
            svc.click((50, 75), button="left", clicks=3)
            mock_click.assert_not_called()
            mock_pg.click.assert_called_once_with(50, 75, button="left", clicks=3, duration=0.1)

    def test_right_button_single_click(self, svc):
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG) as mock_pg:
            svc.click((10, 20), button="right")
            mock_click.assert_called_once_with(10, 20, "right")
            mock_pg.click.assert_not_called()

    def test_middle_button_single_click(self, svc):
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG) as mock_pg:
            svc.click((10, 20), button="middle")
            mock_click.assert_called_once_with(10, 20, "middle")
            mock_pg.click.assert_not_called()

    def test_native_returns_none_falls_back_to_pyautogui(self, svc):
        """When native is unavailable (returns None), pyautogui must be used."""
        with patch(_NATIVE_CLICK, return_value=None) as mock_click, patch(_PG) as mock_pg:
            svc.click((300, 400))
            mock_click.assert_called_once_with(300, 400, "left")
            mock_pg.click.assert_called_once_with(300, 400, button="left", clicks=1, duration=0.1)

    def test_native_returns_zero_logs_warning_and_does_not_fallback(self, svc, caplog):
        """SendInput returning 0 should log a UIPI warning but not call pyautogui."""
        with patch(_NATIVE_CLICK, return_value=0), patch(_PG) as mock_pg:
            with caplog.at_level(logging.WARNING, logger="windows_mcp.input.service"):
                svc.click((100, 100))
            mock_pg.click.assert_not_called()
            assert "UIPI" in caplog.text or "SendInput returned 0" in caplog.text

    def test_negative_coordinates(self, svc):
        """Negative coordinates are forwarded without modification."""
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG):
            svc.click((-1, -1))
            mock_click.assert_called_once_with(-1, -1, "left")

    def test_extreme_large_coordinates(self, svc):
        """Very large coordinates are forwarded without modification."""
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG):
            svc.click((99999, 99999))
            mock_click.assert_called_once_with(99999, 99999, "left")

    def test_zero_zero_coordinates(self, svc):
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG):
            svc.click((0, 0))
            mock_click.assert_called_once_with(0, 0, "left")

    def test_unknown_button_name_passed_through_to_native(self, svc):
        """Invalid button names are passed through; validation is the caller's concern."""
        with patch(_NATIVE_CLICK, return_value=2) as mock_click, patch(_PG):
            svc.click((50, 50), button="back")
            mock_click.assert_called_once_with(50, 50, "back")

    def test_unknown_button_double_click_passed_to_pyautogui(self, svc):
        with patch(_NATIVE_CLICK, return_value=2), patch(_PG) as mock_pg:
            svc.click((50, 50), button="extra", clicks=2)
            mock_pg.click.assert_called_once_with(50, 50, button="extra", clicks=2, duration=0.1)


# ===========================================================================
# type()
# ===========================================================================


class TestType:
    """Tests for InputService.type()."""

    def test_basic_type_calls_leftclick_and_native_send(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=5):
            svc.type((100, 200), text="hello")
            mock_pg.leftClick.assert_called_once_with(100, 200)

    def test_empty_text_still_clicks_and_native_called(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=0) as mock_text:
            svc.type((10, 20), text="")
            mock_pg.leftClick.assert_called_once_with(10, 20)
            mock_text.assert_called_once_with("")
            mock_pg.typewrite.assert_not_called()

    def test_native_send_text_returns_none_falls_back_to_typewrite(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=None):
            svc.type((10, 20), text="fallback text")
            mock_pg.typewrite.assert_called_once_with("fallback text", interval=0.02)

    def test_caret_position_start_presses_home(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="x", caret_position="start")
            mock_pg.press.assert_any_call("home")

    def test_caret_position_end_presses_end(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="x", caret_position="end")
            mock_pg.press.assert_any_call("end")

    def test_caret_position_idle_no_press(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="x", caret_position="idle")
            # press should not have been called for home or end
            home_calls = [c for c in mock_pg.press.call_args_list if c == call("home")]
            end_calls = [c for c in mock_pg.press.call_args_list if c == call("end")]
            assert not home_calls
            assert not end_calls

    def test_clear_bool_true_triggers_select_all_and_delete(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="abc", clear=True)
            mock_pg.hotkey.assert_called_with("ctrl", "a")
            mock_pg.press.assert_any_call("backspace")

    def test_clear_string_true_triggers_select_all_and_delete(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="abc", clear="true")
            mock_pg.hotkey.assert_called_with("ctrl", "a")
            mock_pg.press.assert_any_call("backspace")

    def test_clear_string_True_case_insensitive(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="abc", clear="True")
            mock_pg.hotkey.assert_called_with("ctrl", "a")

    def test_clear_string_false_does_not_trigger_select_all(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="abc", clear="false")
            home_calls = [c for c in mock_pg.hotkey.call_args_list if "ctrl" in c.args]
            assert not home_calls

    def test_clear_bool_false_does_not_trigger_select_all(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="abc", clear=False)
            mock_pg.hotkey.assert_not_called()

    def test_press_enter_bool_true_presses_enter(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="done", press_enter=True)
            mock_pg.press.assert_any_call("enter")

    def test_press_enter_string_true_presses_enter(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="done", press_enter="true")
            mock_pg.press.assert_any_call("enter")

    def test_press_enter_string_True_case_insensitive(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="done", press_enter="True")
            mock_pg.press.assert_any_call("enter")

    def test_press_enter_bool_false_no_enter(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="done", press_enter=False)
            enter_calls = [c for c in mock_pg.press.call_args_list if c == call("enter")]
            assert not enter_calls

    def test_press_enter_string_false_no_enter(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="done", press_enter="false")
            enter_calls = [c for c in mock_pg.press.call_args_list if c == call("enter")]
            assert not enter_calls

    def test_unicode_emoji_text_passed_to_native(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=4) as mock_text:
            svc.type((50, 50), text="Hello \U0001f600 World")
            mock_text.assert_called_once_with("Hello \U0001f600 World")
            mock_pg.typewrite.assert_not_called()

    def test_unicode_emoji_text_fallback_when_native_unavailable(self, svc):
        """Emoji text falls back to typewrite when native is None."""
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=None):
            svc.type((50, 50), text="\U0001f600")
            mock_pg.typewrite.assert_called_once_with("\U0001f600", interval=0.02)

    def test_very_long_text(self, svc):
        long_text = "a" * 10_000
        with patch(_PG), patch(_NATIVE_TEXT, return_value=10_000) as mock_text:
            svc.type((0, 0), text=long_text)
            mock_text.assert_called_once_with(long_text)

    def test_text_with_newlines_and_tabs(self, svc):
        with patch(_PG), patch(_NATIVE_TEXT, return_value=5) as mock_text:
            svc.type((0, 0), text="line1\nline2\ttabbed")
            mock_text.assert_called_once_with("line1\nline2\ttabbed")

    def test_clear_and_press_enter_together(self, svc):
        with patch(_PG) as mock_pg, patch(_NATIVE_TEXT, return_value=1):
            svc.type((10, 20), text="combo", clear=True, press_enter=True)
            mock_pg.hotkey.assert_called_with("ctrl", "a")
            mock_pg.press.assert_any_call("backspace")
            mock_pg.press.assert_any_call("enter")


# ===========================================================================
# scroll()
# ===========================================================================


class TestScroll:
    """Tests for InputService.scroll()."""

    # --- Rust fast-path tests ---

    def test_vertical_down_uses_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (500, 500)
            result = svc.scroll(type="vertical", direction="down", wheel_times=3)
            mock_scroll.assert_called_once_with(500, 500, -360, False)
            assert result is None

    def test_vertical_up_uses_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (100, 200)
            result = svc.scroll(type="vertical", direction="up", wheel_times=2)
            mock_scroll.assert_called_once_with(100, 200, 240, False)
            assert result is None

    def test_horizontal_left_uses_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (300, 400)
            result = svc.scroll(type="horizontal", direction="left", wheel_times=1)
            mock_scroll.assert_called_once_with(300, 400, -120, True)
            assert result is None

    def test_horizontal_right_uses_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (300, 400)
            result = svc.scroll(type="horizontal", direction="right", wheel_times=1)
            mock_scroll.assert_called_once_with(300, 400, 120, True)
            assert result is None

    def test_loc_provided_passes_coords_to_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG):
            result = svc.scroll(loc=(400, 300), type="vertical", direction="down")
            mock_scroll.assert_called_once_with(400, 300, -120, False)
            assert result is None

    def test_loc_zero_zero_passes_to_native_scroll(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG):
            result = svc.scroll(loc=(0, 0), type="vertical", direction="up")
            mock_scroll.assert_called_once_with(0, 0, 120, False)
            assert result is None

    def test_wheel_times_zero_sends_zero_delta(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (0, 0)
            result = svc.scroll(type="vertical", direction="down", wheel_times=0)
            mock_scroll.assert_called_once_with(0, 0, 0, False)
            assert result is None

    def test_wheel_times_large_value(self, svc):
        with patch(_NATIVE_SCROLL, return_value=2) as mock_scroll, patch(_PG) as mock_pg:
            mock_pg.position.return_value = (0, 0)
            result = svc.scroll(type="vertical", direction="up", wheel_times=1000)
            mock_scroll.assert_called_once_with(0, 0, 120000, False)
            assert result is None

    # --- Validation tests (no native needed) ---

    def test_vertical_invalid_direction_returns_error(self, svc):
        with patch(_NATIVE_SCROLL), patch(_PG) as mock_pg:
            mock_pg.position.return_value = (0, 0)
            result = svc.scroll(type="vertical", direction="left")
            assert result is not None
            assert "up" in result or "down" in result

    def test_horizontal_invalid_direction_returns_error(self, svc):
        with patch(_NATIVE_SCROLL), patch(_PG) as mock_pg:
            mock_pg.position.return_value = (0, 0)
            result = svc.scroll(type="horizontal", direction="up")
            assert result is not None
            assert "left" in result or "right" in result

    def test_invalid_type_returns_error(self, svc):
        with patch(_NATIVE_SCROLL), patch(_PG) as mock_pg:
            mock_pg.position.return_value = (0, 0)
            result = svc.scroll(type="invalid", direction="down")
            assert result is not None
            assert "horizontal" in result or "vertical" in result

    # --- Fallback tests (native returns None) ---

    def test_fallback_vertical_down_calls_wheel_down(self, svc):
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA) as mock_uia,
            patch(_PG) as mock_pg,
            patch(_NATIVE_MOVE, return_value=1),
        ):
            mock_pg.position.return_value = (500, 500)
            result = svc.scroll(type="vertical", direction="down", wheel_times=3)
            mock_uia.WheelDown.assert_called_once_with(3)
            assert result is None

    def test_fallback_vertical_up_calls_wheel_up(self, svc):
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA) as mock_uia,
            patch(_PG) as mock_pg,
            patch(_NATIVE_MOVE, return_value=1),
        ):
            mock_pg.position.return_value = (500, 500)
            result = svc.scroll(type="vertical", direction="up", wheel_times=2)
            mock_uia.WheelUp.assert_called_once_with(2)
            assert result is None

    def test_fallback_horizontal_left_holds_shift(self, svc):
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA) as mock_uia,
            patch(_PG) as mock_pg,
            patch(_NATIVE_MOVE, return_value=1),
        ):
            mock_pg.position.return_value = (500, 500)
            result = svc.scroll(type="horizontal", direction="left", wheel_times=1)
            mock_pg.keyDown.assert_called_with("Shift")
            mock_uia.WheelDown.assert_called_once_with(1)
            mock_pg.keyUp.assert_called_with("Shift")
            assert result is None

    def test_fallback_horizontal_right_holds_shift(self, svc):
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA) as mock_uia,
            patch(_PG) as mock_pg,
            patch(_NATIVE_MOVE, return_value=1),
        ):
            mock_pg.position.return_value = (500, 500)
            result = svc.scroll(type="horizontal", direction="right", wheel_times=1)
            mock_pg.keyDown.assert_called_with("Shift")
            mock_uia.WheelUp.assert_called_once_with(1)
            mock_pg.keyUp.assert_called_with("Shift")
            assert result is None

    def test_fallback_loc_provided_calls_move(self, svc):
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA),
            patch(_PG),
            patch(_NATIVE_MOVE, return_value=1) as mock_move,
        ):
            result = svc.scroll(loc=(400, 300), type="vertical", direction="down")
            mock_move.assert_called_once_with(400, 300)
            assert result is None

    def test_fallback_horizontal_shift_released_on_error(self, svc):
        """Shift must be released (via finally) even if WheelDown raises."""
        with (
            patch(_NATIVE_SCROLL, return_value=None),
            patch(_UIA) as mock_uia,
            patch(_PG) as mock_pg,
        ):
            mock_pg.position.return_value = (500, 500)
            mock_uia.WheelDown.side_effect = RuntimeError("COM error")
            with pytest.raises(RuntimeError):
                svc.scroll(type="horizontal", direction="left", wheel_times=1)
            mock_pg.keyUp.assert_called_with("Shift")


# ===========================================================================
# drag()
# ===========================================================================


class TestDrag:
    """Tests for InputService.drag()."""

    # --- Rust fast-path tests ---

    def test_drag_uses_native_drag(self, svc):
        with patch(_NATIVE_DRAG, return_value=3) as mock_drag, patch(_PG) as mock_pg:
            svc.drag((250, 350))
            mock_drag.assert_called_once_with(250, 350)
            mock_pg.dragTo.assert_not_called()

    def test_drag_native_negative_coordinates(self, svc):
        with patch(_NATIVE_DRAG, return_value=3) as mock_drag, patch(_PG):
            svc.drag((-10, -20))
            mock_drag.assert_called_once_with(-10, -20)

    def test_drag_native_zero_zero(self, svc):
        with patch(_NATIVE_DRAG, return_value=3) as mock_drag, patch(_PG):
            svc.drag((0, 0))
            mock_drag.assert_called_once_with(0, 0)

    # --- Fallback tests (native returns None) ---

    def test_drag_fallback_calls_pyautogui(self, svc):
        with patch(_NATIVE_DRAG, return_value=None), patch(_PG) as mock_pg:
            svc.drag((250, 350))
            mock_pg.dragTo.assert_called_once_with(250, 350, duration=0.6)

    def test_drag_fallback_calls_sleep_before_drag(self, svc):
        with patch(_NATIVE_DRAG, return_value=None), patch(_PG) as mock_pg:
            svc.drag((100, 200))
            mock_pg.sleep.assert_called_once_with(0.5)
            sleep_idx = mock_pg.method_calls.index(call.sleep(0.5))
            drag_idx = mock_pg.method_calls.index(call.dragTo(100, 200, duration=0.6))
            assert sleep_idx < drag_idx

    def test_drag_fallback_large_coordinates(self, svc):
        with patch(_NATIVE_DRAG, return_value=None), patch(_PG) as mock_pg:
            svc.drag((9999, 9999))
            mock_pg.dragTo.assert_called_once_with(9999, 9999, duration=0.6)


# ===========================================================================
# move()
# ===========================================================================


class TestMove:
    """Tests for InputService.move()."""

    def test_move_uses_native_when_available(self, svc):
        with patch(_NATIVE_MOVE, return_value=1) as mock_move, patch(_PG) as mock_pg:
            svc.move((100, 200))
            mock_move.assert_called_once_with(100, 200)
            mock_pg.moveTo.assert_not_called()

    def test_move_falls_back_when_native_returns_none(self, svc):
        with patch(_NATIVE_MOVE, return_value=None) as mock_move, patch(_PG) as mock_pg:
            svc.move((100, 200))
            mock_move.assert_called_once_with(100, 200)
            mock_pg.moveTo.assert_called_once_with(100, 200, duration=0.1)

    def test_move_zero_zero_native_path(self, svc):
        with patch(_NATIVE_MOVE, return_value=1) as mock_move, patch(_PG) as mock_pg:
            svc.move((0, 0))
            mock_move.assert_called_once_with(0, 0)
            mock_pg.moveTo.assert_not_called()

    def test_move_zero_zero_fallback_path(self, svc):
        with patch(_NATIVE_MOVE, return_value=None), patch(_PG) as mock_pg:
            svc.move((0, 0))
            mock_pg.moveTo.assert_called_once_with(0, 0, duration=0.1)

    def test_move_negative_coordinates_native(self, svc):
        with patch(_NATIVE_MOVE, return_value=1) as mock_move, patch(_PG):
            svc.move((-5, -10))
            mock_move.assert_called_once_with(-5, -10)

    def test_move_negative_coordinates_fallback(self, svc):
        with patch(_NATIVE_MOVE, return_value=None), patch(_PG) as mock_pg:
            svc.move((-5, -10))
            mock_pg.moveTo.assert_called_once_with(-5, -10, duration=0.1)

    def test_move_extreme_coordinates(self, svc):
        with patch(_NATIVE_MOVE, return_value=1) as mock_move, patch(_PG):
            svc.move((65535, 65535))
            mock_move.assert_called_once_with(65535, 65535)


# ===========================================================================
# shortcut()
# ===========================================================================


class TestShortcut:
    """Tests for InputService.shortcut()."""

    def test_single_key_calls_press(self, svc):
        with patch(_PG) as mock_pg:
            svc.shortcut("enter")
            mock_pg.press.assert_called_once_with("enter")
            mock_pg.hotkey.assert_not_called()

    def test_multi_key_calls_hotkey(self, svc):
        with patch(_PG) as mock_pg:
            svc.shortcut("ctrl+c")
            mock_pg.hotkey.assert_called_once_with("ctrl", "c")
            mock_pg.press.assert_not_called()

    def test_three_key_combination(self, svc):
        with patch(_PG) as mock_pg:
            svc.shortcut("ctrl+shift+s")
            mock_pg.hotkey.assert_called_once_with("ctrl", "shift", "s")

    def test_four_key_combination(self, svc):
        with patch(_PG) as mock_pg:
            svc.shortcut("ctrl+alt+shift+f4")
            mock_pg.hotkey.assert_called_once_with("ctrl", "alt", "shift", "f4")

    def test_empty_string_calls_press_with_empty(self, svc):
        """An empty shortcut string passes an empty string to pg.press."""
        with patch(_PG) as mock_pg:
            svc.shortcut("")
            mock_pg.press.assert_called_once_with("")
            mock_pg.hotkey.assert_not_called()

    def test_single_key_uppercase(self, svc):
        """Case is preserved -- pg is responsible for normalisation."""
        with patch(_PG) as mock_pg:
            svc.shortcut("F5")
            mock_pg.press.assert_called_once_with("F5")

    def test_plus_only_string_produces_empty_hotkey_parts(self, svc):
        """A lone '+' splits into two empty strings, routed to hotkey."""
        with patch(_PG) as mock_pg:
            svc.shortcut("+")
            mock_pg.hotkey.assert_called_once_with("", "")

    def test_trailing_plus_passes_through(self, svc):
        """ctrl+ splits into ['ctrl', ''] -- still routed to hotkey."""
        with patch(_PG) as mock_pg:
            svc.shortcut("ctrl+")
            mock_pg.hotkey.assert_called_once_with("ctrl", "")


# ===========================================================================
# multi_select()
# ===========================================================================


class TestMultiSelect:
    """Tests for InputService.multi_select()."""

    def test_empty_locs_list_no_clicks(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[])
            mock_pg.click.assert_not_called()
            mock_pg.keyDown.assert_not_called()
            mock_pg.keyUp.assert_not_called()

    def test_none_locs_defaults_to_empty_no_clicks(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=None)
            mock_pg.click.assert_not_called()

    def test_single_loc_no_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(10, 20)])
            mock_pg.click.assert_called_once_with(10, 20, duration=0.2)
            mock_pg.keyDown.assert_not_called()

    def test_multiple_locs_without_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(10, 20), (30, 40), (50, 60)])
            assert mock_pg.click.call_count == 3
            mock_pg.keyDown.assert_not_called()

    def test_press_ctrl_bool_true_holds_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(press_ctrl=True, locs=[(10, 20)])
            mock_pg.keyDown.assert_called_once_with("ctrl")
            mock_pg.keyUp.assert_called_once_with("ctrl")

    def test_press_ctrl_string_true_holds_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(press_ctrl="true", locs=[(10, 20)])
            mock_pg.keyDown.assert_called_once_with("ctrl")
            mock_pg.keyUp.assert_called_once_with("ctrl")

    def test_press_ctrl_string_True_case_insensitive(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(press_ctrl="True", locs=[(10, 20)])
            mock_pg.keyDown.assert_called_once_with("ctrl")

    def test_press_ctrl_string_false_no_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(press_ctrl="false", locs=[(10, 20)])
            mock_pg.keyDown.assert_not_called()
            mock_pg.keyUp.assert_not_called()

    def test_press_ctrl_bool_false_no_ctrl(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(press_ctrl=False, locs=[(10, 20)])
            mock_pg.keyDown.assert_not_called()

    def test_ctrl_released_in_finally_on_exception(self, svc):
        """ctrl must be released even if an intermediate click raises."""
        with patch(_PG) as mock_pg:
            mock_pg.click.side_effect = [None, RuntimeError("click failed")]
            with pytest.raises(RuntimeError):
                svc.multi_select(press_ctrl=True, locs=[(1, 2), (3, 4)])
            mock_pg.keyUp.assert_called_with("ctrl")

    def test_ctrl_not_released_if_not_held_on_exception(self, svc):
        """When ctrl is not held, keyUp should not be called on exception."""
        with patch(_PG) as mock_pg:
            mock_pg.click.side_effect = RuntimeError("click failed")
            with pytest.raises(RuntimeError):
                svc.multi_select(press_ctrl=False, locs=[(1, 2)])
            mock_pg.keyUp.assert_not_called()

    def test_sleep_called_between_clicks(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(10, 20), (30, 40)])
            assert mock_pg.sleep.call_count == 2
            mock_pg.sleep.assert_any_call(0.5)

    def test_click_order_matches_locs_order(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(1, 2), (3, 4), (5, 6)])
            calls = mock_pg.click.call_args_list
            assert calls[0] == call(1, 2, duration=0.2)
            assert calls[1] == call(3, 4, duration=0.2)
            assert calls[2] == call(5, 6, duration=0.2)

    def test_negative_coordinates_in_locs(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(-1, -2), (-100, -200)])
            assert mock_pg.click.call_count == 2

    def test_zero_zero_loc(self, svc):
        with patch(_PG) as mock_pg:
            svc.multi_select(locs=[(0, 0)])
            mock_pg.click.assert_called_once_with(0, 0, duration=0.2)


# ===========================================================================
# multi_edit()
# ===========================================================================


class TestMultiEdit:
    """Tests for InputService.multi_edit()."""

    def test_empty_list_no_calls(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([])
            mock_type.assert_not_called()

    def test_single_entry_calls_type_with_clear(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(100, 200, "hello")])
            mock_type.assert_called_once_with((100, 200), text="hello", clear=True)

    def test_multiple_entries_each_calls_type(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(10, 20, "alpha"), (30, 40, "beta"), (50, 60, "gamma")])
            assert mock_type.call_count == 3
            mock_type.assert_any_call((10, 20), text="alpha", clear=True)
            mock_type.assert_any_call((30, 40), text="beta", clear=True)
            mock_type.assert_any_call((50, 60), text="gamma", clear=True)

    def test_order_of_calls_matches_input_order(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(1, 1, "first"), (2, 2, "second")])
            calls = mock_type.call_args_list
            assert calls[0] == call((1, 1), text="first", clear=True)
            assert calls[1] == call((2, 2), text="second", clear=True)

    def test_empty_text_string_passed_to_type(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(5, 10, "")])
            mock_type.assert_called_once_with((5, 10), text="", clear=True)

    def test_unicode_text_passed_to_type(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(0, 0, "\U0001f600 emoji")])
            mock_type.assert_called_once_with((0, 0), text="\U0001f600 emoji", clear=True)

    def test_large_batch_calls_type_for_each(self, svc):
        entries = [(i, i, f"text_{i}") for i in range(50)]
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit(entries)
            assert mock_type.call_count == 50

    def test_always_passes_clear_true(self, svc):
        """multi_edit must always pass clear=True regardless of text content."""
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(0, 0, "anything")])
            _, kwargs = mock_type.call_args
            assert kwargs.get("clear") is True

    def test_negative_coordinates_forwarded(self, svc):
        with patch.object(svc, "type") as mock_type:
            svc.multi_edit([(-1, -2, "neg")])
            mock_type.assert_called_once_with((-1, -2), text="neg", clear=True)


# ===========================================================================
# _try_value_pattern()
# ===========================================================================


def _make_value_pattern(
    is_read_only: bool = False,
    current_value: str = "",
    set_value_side_effect=None,
):
    """Build a mock ValuePattern with configurable properties."""
    pattern = MagicMock()
    pattern.IsReadOnly = is_read_only
    pattern.Value = current_value
    if set_value_side_effect is not None:
        pattern.SetValue.side_effect = set_value_side_effect
    return pattern


def _make_uia_mock(element=MagicMock(), pattern=MagicMock()):
    """Build a mock uia module for _try_value_pattern tests."""
    mock_uia = MagicMock()
    mock_uia.ControlFromPoint.return_value = element
    element.GetPattern.return_value = pattern
    return mock_uia


class TestTryValuePattern:
    """Unit tests for InputService._try_value_pattern()."""

    # ------------------------------------------------------------------
    # Happy path: element found, writable pattern present
    # ------------------------------------------------------------------

    def test_clear_true_calls_set_value_with_text(self, svc):
        pattern = _make_value_pattern(current_value="old")
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            result = InputService._try_value_pattern(100, 200, "new_text", clear=True)
        assert result is True
        pattern.SetValue.assert_called_once_with("new_text")

    def test_clear_false_appends_to_current_value(self, svc):
        pattern = _make_value_pattern(current_value="hello")
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            result = InputService._try_value_pattern(100, 200, " world", clear=False)
        assert result is True
        pattern.SetValue.assert_called_once_with("hello world")

    def test_clear_false_with_none_current_value_treats_as_empty(self, svc):
        """When Value is None, treat it as empty string and just append."""
        pattern = _make_value_pattern(current_value=None)
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            result = InputService._try_value_pattern(0, 0, "appended", clear=False)
        assert result is True
        pattern.SetValue.assert_called_once_with("appended")

    def test_clear_false_empty_current_value_appends_text(self, svc):
        pattern = _make_value_pattern(current_value="")
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            result = InputService._try_value_pattern(0, 0, "typed", clear=False)
        assert result is True
        pattern.SetValue.assert_called_once_with("typed")

    def test_returns_true_on_success(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            result = InputService._try_value_pattern(50, 50, "text", clear=True)
        assert result is True

    def test_coordinates_passed_to_control_from_point(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = MagicMock()
            mock_uia.ControlFromPoint.return_value.GetPattern.return_value = pattern
            InputService._try_value_pattern(300, 400, "x", clear=True)
        mock_uia.ControlFromPoint.assert_called_once_with(300, 400)

    def test_pattern_id_value_pattern_used_for_get_pattern(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            InputService._try_value_pattern(0, 0, "x", clear=True)
        elem.GetPattern.assert_called_once_with(mock_uia.PatternId.ValuePattern)

    # ------------------------------------------------------------------
    # Failure cases: no element, no pattern, read-only, exception
    # ------------------------------------------------------------------

    def test_returns_false_when_no_element(self, svc):
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.return_value = None
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_when_no_pattern(self, svc):
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = None
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_when_pattern_is_falsy(self, svc):
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = False
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_when_read_only(self, svc):
        pattern = _make_value_pattern(is_read_only=True)
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False
        pattern.SetValue.assert_not_called()

    def test_returns_false_on_control_from_point_exception(self, svc):
        with patch(_UIA) as mock_uia:
            mock_uia.ControlFromPoint.side_effect = OSError("COM error")
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_on_get_pattern_exception(self, svc):
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.side_effect = RuntimeError("pattern failed")
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_on_set_value_exception(self, svc):
        pattern = _make_value_pattern(set_value_side_effect=RuntimeError("write failed"))
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    def test_returns_false_on_is_read_only_exception(self, svc):
        """If accessing IsReadOnly raises, return False gracefully."""
        pattern = MagicMock()
        type(pattern).IsReadOnly = property(lambda self: (_ for _ in ()).throw(OSError("access")))
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "text", clear=True)
        assert result is False

    # ------------------------------------------------------------------
    # Edge cases
    # ------------------------------------------------------------------

    def test_empty_text_clear_true_sets_empty_string(self, svc):
        pattern = _make_value_pattern(current_value="existing")
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "", clear=True)
        assert result is True
        pattern.SetValue.assert_called_once_with("")

    def test_unicode_text_passed_to_set_value(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, "\U0001f600 emoji", clear=True)
        assert result is True
        pattern.SetValue.assert_called_once_with("\U0001f600 emoji")

    def test_large_text_passed_through(self, svc):
        big_text = "x" * 10_000
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            result = InputService._try_value_pattern(0, 0, big_text, clear=True)
        assert result is True
        pattern.SetValue.assert_called_once_with(big_text)

    def test_zero_zero_coordinates_passed_to_control_from_point(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            InputService._try_value_pattern(0, 0, "x", clear=True)
        mock_uia.ControlFromPoint.assert_called_once_with(0, 0)

    def test_negative_coordinates_passed_to_control_from_point(self, svc):
        pattern = _make_value_pattern()
        with patch(_UIA) as mock_uia:
            elem = MagicMock()
            elem.GetPattern.return_value = pattern
            mock_uia.ControlFromPoint.return_value = elem
            InputService._try_value_pattern(-10, -20, "x", clear=True)
        mock_uia.ControlFromPoint.assert_called_once_with(-10, -20)


# ===========================================================================
# type() + ValuePattern integration
# ===========================================================================


class TestTypeValuePatternIntegration:
    """Integration tests for how type() dispatches to _try_value_pattern."""

    def test_idle_caret_success_skips_left_click(self, svc):
        """When ValuePattern succeeds and caret='idle', pg.leftClick is NOT called."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True),
        ):
            svc.type((100, 200), text="hello", caret_position="idle")
        mock_pg.leftClick.assert_not_called()

    def test_idle_caret_success_skips_native_send_text(self, svc):
        """When ValuePattern succeeds, native_send_text is NOT called."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT) as mock_text,
            patch.object(InputService, "_try_value_pattern", return_value=True),
        ):
            svc.type((100, 200), text="hello", caret_position="idle")
        mock_text.assert_not_called()

    def test_idle_caret_failure_falls_through_to_click_and_type(self, svc):
        """When ValuePattern fails and caret='idle', falls through to standard path."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=False),
        ):
            svc.type((100, 200), text="hello", caret_position="idle")
        mock_pg.leftClick.assert_called_once_with(100, 200)

    def test_start_caret_never_calls_value_pattern(self, svc):
        """caret_position='start' should skip ValuePattern entirely."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern") as mock_vp,
        ):
            svc.type((100, 200), text="hello", caret_position="start")
        mock_vp.assert_not_called()

    def test_end_caret_never_calls_value_pattern(self, svc):
        """caret_position='end' should skip ValuePattern entirely."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern") as mock_vp,
        ):
            svc.type((100, 200), text="hello", caret_position="end")
        mock_vp.assert_not_called()

    def test_idle_caret_success_with_press_enter_calls_enter(self, svc):
        """After ValuePattern success, press_enter=True should still press Enter."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True),
        ):
            svc.type((100, 200), text="hello", caret_position="idle", press_enter=True)
        mock_pg.press.assert_called_once_with("enter")

    def test_idle_caret_success_without_press_enter_no_enter_key(self, svc):
        """After ValuePattern success, press_enter=False means no Enter keypress."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True),
        ):
            svc.type((100, 200), text="hello", caret_position="idle", press_enter=False)
        mock_pg.press.assert_not_called()

    def test_value_pattern_called_with_correct_args_clear_true(self, svc):
        """Verify _try_value_pattern receives x, y, text, and is_clear correctly."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True) as mock_vp,
        ):
            svc.type((300, 400), text="abc", caret_position="idle", clear=True)
        mock_vp.assert_called_once_with(300, 400, "abc", True)

    def test_value_pattern_called_with_correct_args_clear_false(self, svc):
        """clear=False must propagate as False to _try_value_pattern."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True) as mock_vp,
        ):
            svc.type((300, 400), text="abc", caret_position="idle", clear=False)
        mock_vp.assert_called_once_with(300, 400, "abc", False)

    def test_value_pattern_called_with_string_clear_true_coerced(self, svc):
        """clear='true' (string) must coerce to bool True before reaching _try_value_pattern."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True) as mock_vp,
        ):
            svc.type((0, 0), text="x", caret_position="idle", clear="true")
        mock_vp.assert_called_once_with(0, 0, "x", True)

    def test_idle_caret_calls_value_pattern_once(self, svc):
        """_try_value_pattern must be called exactly once per type() invocation."""
        with (
            patch(_PG),
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=True) as mock_vp,
        ):
            svc.type((10, 20), text="test", caret_position="idle")
        assert mock_vp.call_count == 1

    def test_idle_failure_clear_true_falls_through_with_select_all(self, svc):
        """When ValuePattern fails and clear=True, standard path does ctrl+a, backspace."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
            patch.object(InputService, "_try_value_pattern", return_value=False),
        ):
            svc.type((0, 0), text="x", caret_position="idle", clear=True)
        mock_pg.hotkey.assert_called_with("ctrl", "a")
        mock_pg.press.assert_any_call("backspace")

    def test_start_caret_standard_path_uses_left_click_and_home(self, svc):
        """caret='start' always goes through standard path: leftClick + Home key."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
        ):
            svc.type((50, 60), text="hi", caret_position="start")
        mock_pg.leftClick.assert_called_once_with(50, 60)
        mock_pg.press.assert_any_call("home")

    def test_end_caret_standard_path_uses_left_click_and_end(self, svc):
        """caret='end' always goes through standard path: leftClick + End key."""
        with (
            patch(_PG) as mock_pg,
            patch(_NATIVE_TEXT, return_value=1),
        ):
            svc.type((50, 60), text="hi", caret_position="end")
        mock_pg.leftClick.assert_called_once_with(50, 60)
        mock_pg.press.assert_any_call("end")
