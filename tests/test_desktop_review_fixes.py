"""Unit tests for bug fixes in desktop/service.py.

Each test class maps to one fix. All UIA / COM / win32 interactions are mocked
so the suite runs headless with no live desktop.

Bug fixes covered:
  1. app() with name=None returns error string instead of raising AttributeError
  2. switch_app None-dereference guard after windows.get()
  3. is_overlay_window handles element.Name=None without AttributeError
  4. get_element_from_xpath raises ValueError for index=0 (1-based indexing)
  5. get_annotated_screenshot padding: width/height += 2*padding (not 1.5*padding)
  6. get_state raises ValueError for out-of-range scale values
  7. list_processes clamps negative limit to 1 via max(1, limit)
  8. auto_minimize skips ShowWindow when GetForegroundWindow returns 0
"""

import sys
import threading
from unittest.mock import MagicMock, patch

import pytest
from PIL import Image

from windows_mcp.desktop.views import DesktopState, Status, Window
from windows_mcp.tree.views import BoundingBox

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_UIA = "windows_mcp.desktop.service.uia"


def _make_bare_desktop():
    """Return a Desktop instance that bypasses __init__ (no COM/UIA calls).

    Only the attributes accessed by the methods under test are set.
    """
    from windows_mcp.desktop.service import Desktop

    d = Desktop.__new__(Desktop)
    d._state_lock = threading.Lock()
    d.desktop_state = None
    d._app_cache = None
    d._app_cache_time = 0.0
    d._APP_CACHE_TTL = 3600.0
    d._app_cache_lock = threading.Lock()
    d._process_name_cache = {}
    d._process_cache_lock = threading.Lock()
    # Sub-services - replaced per-test with MagicMock where needed
    d.tree = MagicMock()
    d._input = MagicMock()
    d._registry = MagicMock()
    d._shell = MagicMock()
    d._scraper = MagicMock()
    d._screen = MagicMock()
    d._window = MagicMock()
    return d


def _make_window(name="Test App", status=Status.NORMAL, handle=1001, process_id=2001):
    """Create a Window dataclass with sensible defaults."""
    bbox = BoundingBox(left=0, top=0, right=800, bottom=600, width=800, height=600)
    return Window(
        name=name,
        is_browser=False,
        depth=0,
        status=status,
        bounding_box=bbox,
        handle=handle,
        process_id=process_id,
    )


def _make_desktop_state(active_window=None, windows=None):
    return DesktopState(
        active_desktop={"id": "1", "name": "Desktop 1"},
        all_desktops=[{"id": "1", "name": "Desktop 1"}],
        active_window=active_window,
        windows=windows or [],
    )


# ===========================================================================
# Fix 1 -- app() with name=None
# ===========================================================================


class TestAppNameNoneGuard:
    """app() must return an error string when name=None for launch/switch modes."""

    def test_launch_mode_name_none_returns_error_string(self):
        d = _make_bare_desktop()
        result = d.app(mode="launch", name=None)
        assert isinstance(result, str)
        assert "required" in result.lower()

    def test_switch_mode_name_none_returns_error_string(self):
        d = _make_bare_desktop()
        result = d.app(mode="switch", name=None)
        assert isinstance(result, str)
        assert "required" in result.lower()

    def test_launch_name_none_does_not_raise_attribute_error(self):
        """Before the fix, name.title() would raise AttributeError when name=None."""
        d = _make_bare_desktop()
        # Must not raise -- just return an error string
        try:
            result = d.app(mode="launch", name=None)
        except AttributeError:
            pytest.fail("app(mode='launch', name=None) raised AttributeError")
        assert result is not None

    def test_switch_name_none_does_not_raise_attribute_error(self):
        d = _make_bare_desktop()
        try:
            result = d.app(mode="switch", name=None)
        except AttributeError:
            pytest.fail("app(mode='switch', name=None) raised AttributeError")
        assert result is not None

    def test_resize_mode_does_not_require_name(self):
        """Resize mode should NOT require name -- it works on the active window."""
        d = _make_bare_desktop()
        # Give resize_app something to work with so it doesn't blow up
        d.resize_app = MagicMock(return_value=("Resized OK.", 0))
        result = d.app(mode="resize", name=None)
        # Should have called resize_app, not returned an error about name
        d.resize_app.assert_called_once()
        assert result is not None

    def test_resize_mode_with_name_still_works(self):
        d = _make_bare_desktop()
        d.resize_app = MagicMock(return_value=("Resized OK.", 0))
        result = d.app(mode="resize", name="Notepad", size=(800, 600))
        d.resize_app.assert_called_once()
        assert result == "Resized OK."

    def test_error_message_content_launch(self):
        d = _make_bare_desktop()
        result = d.app(mode="launch", name=None)
        # Should mention both modes or at least mention launch
        assert "launch" in result.lower() or "name" in result.lower()

    def test_error_message_content_switch(self):
        d = _make_bare_desktop()
        result = d.app(mode="switch", name=None)
        assert "switch" in result.lower() or "name" in result.lower()


# ===========================================================================
# Fix 2 -- switch_app None-dereference guard
# ===========================================================================


class TestSwitchAppNoneDereference:
    """switch_app must handle windows.get() returning None gracefully."""

    def test_window_not_in_dict_returns_error_tuple(self):
        """
        Even if extractOne returns a key, windows.get(key) could return None.
        The fix adds an explicit None guard that returns an error tuple.
        """
        d = _make_bare_desktop()
        win = _make_window(name="Notepad", handle=999)
        state = _make_desktop_state(active_window=win, windows=[])
        d.desktop_state = state

        with patch(_UIA):
            # Patch thefuzz.process.extractOne to return a name that is NOT in the dict
            # This simulates the theoretical edge case where extractOne and dict diverge.
            with patch(
                "windows_mcp.desktop.service.process.extractOne",
                return_value=("Ghost Window", 90.0),
            ):
                result = d.switch_app("Ghost Window")

        # Should be a tuple with a non-zero status code
        assert isinstance(result, tuple)
        assert result[1] != 0

    def test_switch_app_no_match_returns_error(self):
        """extractOne returning None should produce a clear error, not a crash."""
        d = _make_bare_desktop()
        win = _make_window(name="Notepad")
        state = _make_desktop_state(active_window=win, windows=[])
        d.desktop_state = state

        with patch(_UIA):
            with patch(
                "windows_mcp.desktop.service.process.extractOne",
                return_value=None,
            ):
                result = d.switch_app("NonexistentApp")

        assert isinstance(result, tuple)
        assert result[1] != 0

    def test_switch_app_none_state_triggers_get_state(self):
        """When desktop_state is None, switch_app should call get_state() to refresh."""
        d = _make_bare_desktop()
        d.desktop_state = None

        refreshed_win = _make_window(name="Notepad")
        refreshed_state = _make_desktop_state(active_window=refreshed_win, windows=[])

        def fake_get_state():
            d.desktop_state = refreshed_state

        d.get_state = MagicMock(side_effect=fake_get_state)

        with patch(_UIA) as mock_uia:
            mock_uia.IsIconic.return_value = False
            with patch(
                "windows_mcp.desktop.service.process.extractOne",
                return_value=("Notepad", 95.0),
            ):
                d.bring_window_to_top = MagicMock()
                result = d.switch_app("Notepad")

        d.get_state.assert_called_once()
        assert isinstance(result, tuple)


# ===========================================================================
# Fix 3 -- is_overlay_window handles element.Name=None
# ===========================================================================


class TestIsOverlayWindowNoneName:
    """is_overlay_window must not raise AttributeError when element.Name is None."""

    def _make_svc(self):
        from windows_mcp.window.service import WindowService

        return WindowService()

    def _make_ctrl(self, name, children_count: int = 1):
        ctrl = MagicMock()
        ctrl.Name = name
        ctrl.GetChildren.return_value = [MagicMock() for _ in range(children_count)]
        return ctrl

    def test_none_name_does_not_raise(self):
        svc = self._make_svc()
        ctrl = self._make_ctrl(name=None, children_count=1)
        try:
            result = svc.is_overlay_window(ctrl)
        except AttributeError:
            pytest.fail("is_overlay_window raised AttributeError for Name=None")
        assert result is False

    def test_none_name_no_children_is_overlay(self):
        """No children => overlay regardless of name."""
        svc = self._make_svc()
        ctrl = self._make_ctrl(name=None, children_count=0)
        assert svc.is_overlay_window(ctrl) is True

    def test_overlay_in_name_is_overlay(self):
        svc = self._make_svc()
        ctrl = self._make_ctrl(name="NVIDIA Overlay", children_count=1)
        assert svc.is_overlay_window(ctrl) is True

    def test_normal_name_with_children_not_overlay(self):
        svc = self._make_svc()
        ctrl = self._make_ctrl(name="Notepad", children_count=3)
        assert svc.is_overlay_window(ctrl) is False

    def test_empty_name_string_not_overlay(self):
        svc = self._make_svc()
        ctrl = self._make_ctrl(name="", children_count=1)
        assert svc.is_overlay_window(ctrl) is False

    def test_overlay_keyword_case_sensitive(self):
        """'Overlay' must appear exactly (capital O) to be detected."""
        svc = self._make_svc()
        ctrl = self._make_ctrl(name="overlay window", children_count=1)
        assert svc.is_overlay_window(ctrl) is False

    def test_name_overlay_exact_word(self):
        svc = self._make_svc()
        ctrl = self._make_ctrl(name="Overlay", children_count=1)
        assert svc.is_overlay_window(ctrl) is True


# ===========================================================================
# Fix 5 -- get_element_from_xpath index=0 raises ValueError
# ===========================================================================


class TestGetElementFromXpathIndex:
    """XPath indices are 1-based; index=0 must raise ValueError."""

    def _build_xpath_desktop(self, children_count=2):
        """Return a desktop with GetRootControl mocked."""
        d = _make_bare_desktop()
        return d

    def test_index_zero_raises_value_error(self):
        d = _make_bare_desktop()
        child = MagicMock()
        child.ControlTypeName = "Button"

        root = MagicMock()
        root.GetChildren.return_value = [child, child]
        root.ControlTypeName = "Pane"

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            # xpath with index [0] -- 1-based indexing makes this invalid
            with pytest.raises(ValueError, match="index 0 out of range"):
                d.get_element_from_xpath("Pane/Button[0]")

    def test_index_one_resolves_first_child(self):
        d = _make_bare_desktop()
        child_a = MagicMock()
        child_a.ControlTypeName = "Button"
        child_b = MagicMock()
        child_b.ControlTypeName = "Button"

        root = MagicMock()
        root.GetChildren.return_value = [child_a, child_b]

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            result = d.get_element_from_xpath("Pane/Button[1]")
        assert result is child_a

    def test_index_two_resolves_second_child(self):
        d = _make_bare_desktop()
        child_a = MagicMock()
        child_a.ControlTypeName = "Button"
        child_b = MagicMock()
        child_b.ControlTypeName = "Button"

        root = MagicMock()
        root.GetChildren.return_value = [child_a, child_b]

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            result = d.get_element_from_xpath("Pane/Button[2]")
        assert result is child_b

    def test_index_out_of_range_raises_value_error(self):
        d = _make_bare_desktop()
        child = MagicMock()
        child.ControlTypeName = "Button"

        root = MagicMock()
        root.GetChildren.return_value = [child]

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            with pytest.raises(ValueError, match="index 5 out of range"):
                d.get_element_from_xpath("Pane/Button[5]")

    def test_no_index_resolves_first_child(self):
        """No bracketed index selects the first matching child."""
        d = _make_bare_desktop()
        child_a = MagicMock()
        child_a.ControlTypeName = "Button"
        child_b = MagicMock()
        child_b.ControlTypeName = "Button"

        root = MagicMock()
        root.GetChildren.return_value = [child_a, child_b]

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            result = d.get_element_from_xpath("Pane/Button")
        assert result is child_a

    def test_missing_control_type_raises_value_error(self):
        d = _make_bare_desktop()
        root = MagicMock()
        root.GetChildren.return_value = []  # no children at all

        with patch(_UIA) as mock_uia:
            mock_uia.GetRootControl.return_value = root
            with pytest.raises(ValueError, match="no children of type"):
                d.get_element_from_xpath("Pane/Button[1]")


# ===========================================================================
# Fix 6 -- get_annotated_screenshot padding dimensions
# ===========================================================================


class TestGetAnnotatedScreenshotPadding:
    """Padded image dimensions must be screenshot.width + 2*padding and
    screenshot.height + 2*padding (not 1.5*padding).

    Tests ScreenService directly since the logic was extracted from Desktop.
    """

    _SCREEN_UIA = "windows_mcp.screen.service.uia"

    def _make_screen_service(self):
        from windows_mcp.screen.service import ScreenService

        return ScreenService()

    def test_padded_width_is_screenshot_width_plus_2x_padding(self):
        svc = self._make_screen_service()
        padding = 5
        orig_w, orig_h = 200, 100

        fake_screenshot = Image.new("RGB", (orig_w, orig_h), color=(0, 0, 0))

        with patch.object(svc, "get_screenshot", return_value=fake_screenshot):
            with patch(self._SCREEN_UIA) as mock_uia:
                mock_uia.GetVirtualScreenRect.return_value = (0, 0, orig_w, orig_h)
                result = svc.get_annotated_screenshot(nodes=[])

        assert result.width == orig_w + 2 * padding
        assert result.height == orig_h + 2 * padding

    def test_padded_dimensions_are_not_1_5x_padding(self):
        """Guard against regression to 1.5*padding."""
        svc = self._make_screen_service()
        padding = 5
        orig_w, orig_h = 300, 200

        fake_screenshot = Image.new("RGB", (orig_w, orig_h))

        with patch.object(svc, "get_screenshot", return_value=fake_screenshot):
            with patch(self._SCREEN_UIA) as mock_uia:
                mock_uia.GetVirtualScreenRect.return_value = (0, 0, orig_w, orig_h)
                result = svc.get_annotated_screenshot(nodes=[])

        expected_w = orig_w + 2 * padding  # 310
        expected_h = orig_h + 2 * padding  # 210
        wrong_w = orig_w + int(1.5 * padding)  # 307 -- old wrong formula
        wrong_h = orig_h + int(1.5 * padding)  # 207

        assert result.width == expected_w, (
            f"Expected width {expected_w}, got {result.width} "
            f"(wrong 1.5x formula would give {wrong_w})"
        )
        assert result.height == expected_h, (
            f"Expected height {expected_h}, got {result.height} "
            f"(wrong 1.5x formula would give {wrong_h})"
        )

    def test_padded_image_is_rgb(self):
        svc = self._make_screen_service()
        orig_w, orig_h = 100, 80
        fake_screenshot = Image.new("RGB", (orig_w, orig_h))

        with patch.object(svc, "get_screenshot", return_value=fake_screenshot):
            with patch(self._SCREEN_UIA) as mock_uia:
                mock_uia.GetVirtualScreenRect.return_value = (0, 0, orig_w, orig_h)
                result = svc.get_annotated_screenshot(nodes=[])

        assert result.mode == "RGB"

    def test_padding_value_is_5(self):
        """Verify the hardcoded padding=5 gives the right offset for the paste."""
        svc = self._make_screen_service()
        orig_w, orig_h = 50, 40
        fake_screenshot = Image.new("RGB", (orig_w, orig_h), color=(255, 0, 0))

        with patch.object(svc, "get_screenshot", return_value=fake_screenshot):
            with patch(self._SCREEN_UIA) as mock_uia:
                mock_uia.GetVirtualScreenRect.return_value = (0, 0, orig_w, orig_h)
                result = svc.get_annotated_screenshot(nodes=[])

        # Pixel at (0,0) should be white (padding background)
        assert result.getpixel((0, 0)) == (255, 255, 255)
        # Pixel at (5,5) should be red (the pasted screenshot)
        assert result.getpixel((5, 5)) == (255, 0, 0)


# ===========================================================================
# Fix 7 -- get_state scale validation
# ===========================================================================


class TestGetStateScaleValidation:
    """get_state must raise ValueError for scale outside [0.1, 4.0]."""

    def test_scale_zero_raises_value_error(self):
        d = _make_bare_desktop()
        with pytest.raises(ValueError, match="scale"):
            d.get_state(scale=0)

    def test_scale_negative_raises_value_error(self):
        d = _make_bare_desktop()
        with pytest.raises(ValueError, match="scale"):
            d.get_state(scale=-1)

    def test_scale_above_max_raises_value_error(self):
        d = _make_bare_desktop()
        with pytest.raises(ValueError, match="scale"):
            d.get_state(scale=5.0)

    def test_scale_exactly_max_boundary_raises(self):
        """4.01 is just above the limit and must raise."""
        d = _make_bare_desktop()
        with pytest.raises(ValueError):
            d.get_state(scale=4.01)

    def test_scale_exactly_min_boundary_raises(self):
        """0.09 is just below the minimum and must raise."""
        d = _make_bare_desktop()
        with pytest.raises(ValueError):
            d.get_state(scale=0.09)

    def test_scale_1_0_is_valid(self):
        """scale=1.0 is the default and must not raise before other logic runs."""
        d = _make_bare_desktop()
        # get_state calls many other methods; we only need to confirm the ValueError
        # does NOT fire at the scale check. The subsequent calls will fail on their
        # own mocked/absent dependencies -- we catch anything that isn't ValueError.
        try:
            d.get_state(scale=1.0)
        except ValueError as exc:
            if "scale" in str(exc):
                pytest.fail(f"scale=1.0 raised ValueError: {exc}")
        except Exception:
            pass  # expected -- other unrelated mocks are not set up

    def test_scale_0_5_is_valid(self):
        d = _make_bare_desktop()
        try:
            d.get_state(scale=0.5)
        except ValueError as exc:
            if "scale" in str(exc):
                pytest.fail(f"scale=0.5 raised ValueError: {exc}")
        except Exception:
            pass

    def test_scale_2_0_is_valid(self):
        d = _make_bare_desktop()
        try:
            d.get_state(scale=2.0)
        except ValueError as exc:
            if "scale" in str(exc):
                pytest.fail(f"scale=2.0 raised ValueError: {exc}")
        except Exception:
            pass

    def test_scale_0_1_exact_min_is_valid(self):
        d = _make_bare_desktop()
        try:
            d.get_state(scale=0.1)
        except ValueError as exc:
            if "scale" in str(exc):
                pytest.fail(f"scale=0.1 raised ValueError: {exc}")
        except Exception:
            pass

    def test_scale_4_0_exact_max_is_valid(self):
        d = _make_bare_desktop()
        try:
            d.get_state(scale=4.0)
        except ValueError as exc:
            if "scale" in str(exc):
                pytest.fail(f"scale=4.0 raised ValueError: {exc}")
        except Exception:
            pass

    def test_scale_error_message_contains_value(self):
        d = _make_bare_desktop()
        with pytest.raises(ValueError, match="5.0"):
            d.get_state(scale=5.0)


# ===========================================================================
# Fix 8 -- list_processes clamps negative limit
# ===========================================================================


class TestListProcessesNegativeLimit:
    """limit=-1 must be clamped to 1 (not return all-but-last via slicing).

    list_processes() imports psutil and tabulate inside the function body, so
    they are patched via sys.modules rather than as module-level attributes.
    """

    def _make_proc_info(self, pid: int, name: str, cpu: float = 0.0, mem_rss: int = 1024 * 1024):
        mem = MagicMock()
        mem.rss = mem_rss
        return {"pid": pid, "name": name, "cpu_percent": cpu, "memory_info": mem}

    def _make_mock_psutil(self, proc_infos: list) -> MagicMock:
        """Build a psutil mock whose process_iter yields the given info dicts."""
        mock_psutil = MagicMock()
        mock_p_list = []
        for info in proc_infos:
            p = MagicMock()
            p.info = info
            mock_p_list.append(p)
        mock_psutil.process_iter.return_value = mock_p_list
        # NoSuchProcess and AccessDenied are used as exception types in the except clause;
        # using a base Exception subclass keeps the guard active without triggering it.
        mock_psutil.NoSuchProcess = type("NoSuchProcess", (Exception,), {})
        mock_psutil.AccessDenied = type("AccessDenied", (Exception,), {})
        return mock_psutil

    def _run_list_processes(self, d, proc_infos: list, limit: int, mock_tab: MagicMock):
        """Inject psutil and tabulate via sys.modules and call list_processes."""
        mock_psutil = self._make_mock_psutil(proc_infos)
        # tabulate is also a local import inside the function; patch via sys.modules
        orig_psutil = sys.modules.get("psutil")
        orig_tabulate = sys.modules.get("tabulate")
        try:
            sys.modules["psutil"] = mock_psutil
            sys.modules["tabulate"] = MagicMock(tabulate=mock_tab)
            d.list_processes(limit=limit)
        finally:
            if orig_psutil is None:
                sys.modules.pop("psutil", None)
            else:
                sys.modules["psutil"] = orig_psutil
            if orig_tabulate is None:
                sys.modules.pop("tabulate", None)
            else:
                sys.modules["tabulate"] = orig_tabulate
        return mock_psutil

    def test_negative_limit_clamped_to_one(self):
        d = _make_bare_desktop()
        proc_infos = [
            self._make_proc_info(1, "proc_a", mem_rss=100 * 1024 * 1024),
            self._make_proc_info(2, "proc_b", mem_rss=200 * 1024 * 1024),
            self._make_proc_info(3, "proc_c", mem_rss=300 * 1024 * 1024),
        ]
        mock_tab = MagicMock(return_value="table")
        self._run_list_processes(d, proc_infos, limit=-1, mock_tab=mock_tab)

        call_args = mock_tab.call_args
        rows = call_args[0][0]  # first positional arg is the rows list
        assert len(rows) == 1, f"limit=-1 should clamp to 1 row, got {len(rows)}"

    def test_negative_large_limit_clamped_to_one(self):
        d = _make_bare_desktop()
        proc_infos = [self._make_proc_info(i, f"proc_{i}") for i in range(5)]
        mock_tab = MagicMock(return_value="table")
        self._run_list_processes(d, proc_infos, limit=-100, mock_tab=mock_tab)

        rows = mock_tab.call_args[0][0]
        assert len(rows) == 1

    def test_limit_zero_also_clamped_to_one(self):
        """limit=0 is also non-positive and should clamp to 1."""
        d = _make_bare_desktop()
        proc_infos = [self._make_proc_info(i, f"proc_{i}") for i in range(3)]
        mock_tab = MagicMock(return_value="table")
        self._run_list_processes(d, proc_infos, limit=0, mock_tab=mock_tab)

        rows = mock_tab.call_args[0][0]
        assert len(rows) == 1

    def test_positive_limit_not_clamped(self):
        """Positive limit values must be respected as-is."""
        d = _make_bare_desktop()
        proc_infos = [self._make_proc_info(i, f"proc_{i}") for i in range(10)]
        mock_tab = MagicMock(return_value="table")
        self._run_list_processes(d, proc_infos, limit=5, mock_tab=mock_tab)

        rows = mock_tab.call_args[0][0]
        assert len(rows) == 5

    def test_limit_1_returns_one_row(self):
        d = _make_bare_desktop()
        proc_infos = [self._make_proc_info(i, f"proc_{i}") for i in range(4)]
        mock_tab = MagicMock(return_value="table")
        self._run_list_processes(d, proc_infos, limit=1, mock_tab=mock_tab)

        rows = mock_tab.call_args[0][0]
        assert len(rows) == 1


# ===========================================================================
# Fix 8 -- auto_minimize skips ShowWindow when GetForegroundWindow returns 0
# ===========================================================================

_WINDOW_UIA = "windows_mcp.window.service.uia"


class TestAutoMinimizeHandleZero:
    """When GetForegroundWindow returns 0 (no foreground window),
    auto_minimize must yield without calling ShowWindow."""

    def _make_svc(self):
        from windows_mcp.window.service import WindowService

        return WindowService()

    def test_handle_zero_skips_show_window(self):
        svc = self._make_svc()
        with patch(_WINDOW_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 0
            executed = []

            with svc.auto_minimize():
                executed.append("body")

        assert "body" in executed, "Context body must execute even with handle=0"
        mock_uia.ShowWindow.assert_not_called()

    def test_handle_nonzero_calls_show_window(self):
        """With a valid handle, ShowWindow should be called for minimize and restore."""
        import win32con

        svc = self._make_svc()
        with patch(_WINDOW_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 12345

            with svc.auto_minimize():
                pass

        calls = mock_uia.ShowWindow.call_args_list
        assert len(calls) == 2
        handles = [c[0][0] for c in calls]
        assert all(h == 12345 for h in handles)
        sw_values = [c[0][1] for c in calls]
        assert win32con.SW_MINIMIZE in sw_values
        assert win32con.SW_RESTORE in sw_values

    def test_handle_zero_body_exception_propagates(self):
        """Exceptions in the body must still propagate when handle=0."""
        svc = self._make_svc()
        with patch(_WINDOW_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 0

            with pytest.raises(RuntimeError, match="test error"):
                with svc.auto_minimize():
                    raise RuntimeError("test error")

        mock_uia.ShowWindow.assert_not_called()

    def test_handle_nonzero_restore_called_after_exception(self):
        """Even if the body raises, restore (SW_RESTORE) must be called via finally."""
        import win32con

        svc = self._make_svc()
        with patch(_WINDOW_UIA) as mock_uia:
            mock_uia.GetForegroundWindow.return_value = 9999

            with pytest.raises(ValueError):
                with svc.auto_minimize():
                    raise ValueError("inner error")

        restore_calls = [
            c for c in mock_uia.ShowWindow.call_args_list if c[0][1] == win32con.SW_RESTORE
        ]
        assert len(restore_calls) == 1
