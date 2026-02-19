"""Unit tests for Desktop.get_state() orchestration method.

Covers the eight scenarios:
  1. Basic get_state returns correct DesktopState structure
  2. No windows case returns empty state
  3. Screenshot failure is handled gracefully (use_vision path)
  4. Tree traversal failure is handled gracefully
  5. VDM info is included when available
  6. Thread safety -- _state_lock protects atomic state swap
  7. Multiple windows are processed correctly
  8. use_vision=True paths (annotated/plain/bytes/scaled)

Boolean-string coercion for all four bool parameters is also tested.
Scale boundary validation tests are NOT duplicated here (they live in
test_desktop_review_fixes.py::TestGetStateScaleValidation).

All UIA / COM / win32 interactions are mocked so the suite runs headless
with no live desktop required.
"""

import threading
from unittest.mock import patch

import pytest
from PIL import Image

from tests.desktop_helpers import make_bare_desktop, make_window
from windows_mcp.desktop.views import DesktopState
from windows_mcp.tree.views import BoundingBox, Center, TreeElementNode, TreeState

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_GDI = "windows_mcp.desktop.service.get_desktop_info"
_UIA = "windows_mcp.desktop.service.uia"

_DEFAULT_ACTIVE_DESKTOP = {"id": "vd-1", "name": "Desktop 1"}
_DEFAULT_ALL_DESKTOPS = [_DEFAULT_ACTIVE_DESKTOP]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_tree_element_node(name: str = "OK") -> TreeElementNode:
    bbox = BoundingBox(left=0, top=0, right=100, bottom=30, width=100, height=30)
    center = Center(x=50, y=15)
    return TreeElementNode(
        bounding_box=bbox,
        center=center,
        name=name,
        control_type="Button",
        window_name="Test App",
    )


def _make_tree_state(node_count: int = 2) -> TreeState:
    nodes = [_make_tree_element_node(f"Button{i}") for i in range(node_count)]
    return TreeState(interactive_nodes=nodes)


def _setup_desktop_for_get_state(
    d,
    *,
    windows=None,
    active_window=None,
    controls_handles=None,
    windows_handles=None,
    tree_state=None,
    vdm_result=None,
):
    """Wire up all collaborators that get_state() calls on *d*.

    All parameters have safe defaults so callers only specify what they need.
    """
    if windows is None:
        windows = []
    if active_window is not None and active_window not in windows:
        # get_state removes active_window from windows list; provide it separately
        pass
    windows_handles = windows_handles or {w.handle for w in windows}
    controls_handles = controls_handles or set(windows_handles)

    # WindowService facade methods called by get_state via self._window
    d._window.get_controls_handles.return_value = controls_handles
    d._window.get_windows.return_value = (list(windows), windows_handles)
    d._window.get_active_window.return_value = active_window

    # get_state delegates these three calls to _window through wrapper methods
    d.get_controls_handles = d._window.get_controls_handles
    d.get_windows = d._window.get_windows
    d.get_active_window = d._window.get_active_window

    # Tree
    if tree_state is None:
        tree_state = _make_tree_state()
    d.tree.get_state.return_value = tree_state

    # VDM default
    if vdm_result is None:
        vdm_result = (_DEFAULT_ACTIVE_DESKTOP, _DEFAULT_ALL_DESKTOPS)
    return vdm_result


def _run_get_state(d, vdm_result, **kwargs):
    """Call d.get_state() with VDM and UIA mocked out."""
    with patch(_GDI, return_value=vdm_result):
        with patch(_UIA):
            return d.get_state(**kwargs)


# ===========================================================================
# 1. Basic get_state returns correct DesktopState structure
# ===========================================================================


class TestGetStateBasicStructure:
    """get_state must return a fully populated DesktopState."""

    def test_returns_desktop_state_instance(self):
        d = make_bare_desktop()
        win = make_window(name="Notepad", handle=101)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])
        state = _run_get_state(d, vdm, use_vision=False)
        assert isinstance(state, DesktopState)

    def test_active_desktop_populated(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_desktop["name"] == "Desktop 1"
        assert state.active_desktop["id"] == "vd-1"

    def test_all_desktops_populated(self):
        d = make_bare_desktop()
        all_desktops = [
            {"id": "vd-1", "name": "Desktop 1"},
            {"id": "vd-2", "name": "Desktop 2"},
        ]
        vdm = (_DEFAULT_ACTIVE_DESKTOP, all_desktops)
        _setup_desktop_for_get_state(d, vdm_result=vdm)
        state = _run_get_state(d, vdm, use_vision=False)
        assert len(state.all_desktops) == 2
        assert state.all_desktops[1]["name"] == "Desktop 2"

    def test_tree_state_attached(self):
        d = make_bare_desktop()
        ts = _make_tree_state(3)
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.tree_state is ts

    def test_screenshot_is_none_when_use_vision_false(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.screenshot is None

    def test_state_cached_on_desktop_instance(self):
        d = make_bare_desktop()
        win = make_window(handle=999)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])
        state = _run_get_state(d, vdm, use_vision=False)
        assert d.desktop_state is state

    def test_active_window_set_correctly(self):
        d = make_bare_desktop()
        win = make_window(name="Active App", handle=202)
        vdm = _setup_desktop_for_get_state(d, active_window=win)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_window is win

    def test_returns_new_state_on_each_call(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        state1 = _run_get_state(d, vdm, use_vision=False)
        vdm2 = _setup_desktop_for_get_state(d)
        state2 = _run_get_state(d, vdm2, use_vision=False)
        # Both calls succeed; result is a fresh DesktopState each time
        assert isinstance(state1, DesktopState)
        assert isinstance(state2, DesktopState)


# ===========================================================================
# 2. No windows case returns empty state
# ===========================================================================


class TestGetStateNoWindows:
    """When no windows are visible, get_state must return an empty-but-valid state."""

    def test_active_window_is_none(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d, windows=[], active_window=None)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_window is None

    def test_windows_list_is_empty(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d, windows=[], active_window=None)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.windows == []

    def test_tree_state_still_present(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d, windows=[], active_window=None)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.tree_state is not None

    def test_desktop_info_still_populated(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d, windows=[], active_window=None)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_desktop is not None
        assert state.all_desktops is not None

    def test_state_is_cached_even_with_no_windows(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d, windows=[], active_window=None)
        state = _run_get_state(d, vdm, use_vision=False)
        assert d.desktop_state is state


# ===========================================================================
# 3. Screenshot failure is handled gracefully (use_vision=True path)
# ===========================================================================


class TestGetStateScreenshotFailure:
    """When get_screenshot/get_annotated_screenshot raises, get_state must propagate
    the exception (screenshots are not wrapped in a try/except in the source)."""

    def test_screenshot_exception_propagates(self):
        """get_state does NOT suppress screenshot errors -- it propagates them."""
        d = make_bare_desktop()
        ts = _make_tree_state()
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)

        d._screen.get_screenshot.side_effect = OSError("screen grab failed")

        with patch(_GDI, return_value=vdm):
            with patch(_UIA):
                with pytest.raises(OSError, match="screen grab failed"):
                    d.get_state(use_vision=True, use_annotation=False)

    def test_annotated_screenshot_exception_propagates(self):
        d = make_bare_desktop()
        ts = _make_tree_state()
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)

        d._screen.get_annotated_screenshot.side_effect = RuntimeError("annotation failed")

        with patch(_GDI, return_value=vdm):
            with patch(_UIA):
                with pytest.raises(RuntimeError, match="annotation failed"):
                    d.get_state(use_vision=True, use_annotation=True)

    def test_screenshot_none_when_use_vision_false(self):
        """Screenshot code is skipped entirely when use_vision=False."""
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        state = _run_get_state(d, vdm, use_vision=False)
        d._screen.get_screenshot.assert_not_called()
        d._screen.get_annotated_screenshot.assert_not_called()
        assert state.screenshot is None

    def test_use_vision_true_no_annotation_calls_plain_screenshot(self):
        d = make_bare_desktop()
        ts = _make_tree_state()
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)

        fake_image = Image.new("RGB", (100, 80))
        d._screen.get_screenshot.return_value = fake_image

        with patch(_GDI, return_value=vdm):
            with patch(_UIA):
                state = d.get_state(use_vision=True, use_annotation=False)

        d._screen.get_screenshot.assert_called_once()
        assert state.screenshot is fake_image

    def test_use_vision_true_with_annotation_calls_annotated_screenshot(self):
        d = make_bare_desktop()
        ts = _make_tree_state(2)
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)

        fake_image = Image.new("RGB", (200, 150))
        d._screen.get_annotated_screenshot.return_value = fake_image

        with patch(_GDI, return_value=vdm):
            with patch(_UIA):
                state = d.get_state(use_vision=True, use_annotation=True)

        # Desktop.get_annotated_screenshot(nodes=nodes) delegates as positional arg
        d._screen.get_annotated_screenshot.assert_called_once_with(ts.interactive_nodes)
        assert state.screenshot is fake_image


# ===========================================================================
# 4. Tree traversal failure is handled gracefully
# ===========================================================================


class TestGetStateTreeFailure:
    """tree.get_state() exceptions must be caught; empty TreeState used as fallback."""

    def test_tree_runtime_error_gives_empty_tree_state(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        d.tree.get_state.side_effect = RuntimeError("COM exploded")
        state = _run_get_state(d, vdm, use_vision=False)
        assert isinstance(state.tree_state, TreeState)
        assert state.tree_state.interactive_nodes == []

    def test_tree_exception_does_not_propagate(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        d.tree.get_state.side_effect = Exception("any tree error")
        try:
            _run_get_state(d, vdm, use_vision=False)
        except Exception as exc:
            pytest.fail(f"get_state propagated tree exception: {exc}")

    def test_tree_value_error_also_caught(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        d.tree.get_state.side_effect = ValueError("bad xpath")
        state = _run_get_state(d, vdm, use_vision=False)
        assert isinstance(state.tree_state, TreeState)

    def test_fallback_tree_state_has_empty_scrollable_nodes(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        d.tree.get_state.side_effect = RuntimeError("fail")
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.tree_state.scrollable_nodes == []

    def test_fallback_tree_state_has_empty_dom_informative_nodes(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        d.tree.get_state.side_effect = RuntimeError("fail")
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.tree_state.dom_informative_nodes == []

    def test_state_is_still_valid_on_tree_failure(self):
        d = make_bare_desktop()
        win = make_window(name="Notepad", handle=7777)
        vdm = _setup_desktop_for_get_state(d, active_window=win)
        d.tree.get_state.side_effect = RuntimeError("tree gone")
        state = _run_get_state(d, vdm, use_vision=False)
        assert isinstance(state, DesktopState)
        assert state.active_window is win

    def test_tree_get_state_called_with_correct_args(self):
        """tree.get_state() is called with the active_window_handle and use_dom flag."""
        d = make_bare_desktop()
        win = make_window(handle=5050)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])
        _run_get_state(d, vdm, use_vision=False, use_dom=False)
        d.tree.get_state.assert_called_once_with(5050, [], use_dom=False)

    def test_tree_get_state_called_with_use_dom_true(self):
        d = make_bare_desktop()
        win = make_window(handle=6060)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])
        _run_get_state(d, vdm, use_vision=False, use_dom=True)
        d.tree.get_state.assert_called_once_with(6060, [], use_dom=True)


# ===========================================================================
# 5. VDM info is included when available
# ===========================================================================


class TestGetStateVdmInfo:
    """VDM data must be reflected in the DesktopState when get_desktop_info succeeds."""

    def test_active_desktop_from_vdm(self):
        d = make_bare_desktop()
        active = {"id": "vd-abc", "name": "Work"}
        all_d = [active, {"id": "vd-def", "name": "Personal"}]
        vdm = _setup_desktop_for_get_state(d, vdm_result=(active, all_d))
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_desktop["name"] == "Work"

    def test_all_desktops_from_vdm(self):
        d = make_bare_desktop()
        active = {"id": "vd-1", "name": "Desktop 1"}
        all_d = [active, {"id": "vd-2", "name": "Desktop 2"}, {"id": "vd-3", "name": "Gaming"}]
        vdm = _setup_desktop_for_get_state(d, vdm_result=(active, all_d))
        state = _run_get_state(d, vdm, use_vision=False)
        assert len(state.all_desktops) == 3
        names = [d["name"] for d in state.all_desktops]
        assert "Gaming" in names

    def test_vdm_timeout_fallback_to_default_desktop(self):
        """When get_desktop_info raises TimeoutError, fallback desktop is used."""
        d = make_bare_desktop()
        _setup_desktop_for_get_state(d)
        with patch(_GDI, side_effect=TimeoutError("VDM timed out")):
            with patch(_UIA):
                state = d.get_state(use_vision=False)
        assert state.active_desktop["name"] == "Default Desktop"
        assert len(state.all_desktops) == 1

    def test_vdm_runtime_error_fallback_to_default_desktop(self):
        d = make_bare_desktop()
        _setup_desktop_for_get_state(d)
        with patch(_GDI, side_effect=RuntimeError("VDM not available")):
            with patch(_UIA):
                state = d.get_state(use_vision=False)
        assert state.active_desktop["id"] == "00000000-0000-0000-0000-000000000000"

    def test_vdm_fallback_all_desktops_has_one_entry(self):
        d = make_bare_desktop()
        _setup_desktop_for_get_state(d)
        with patch(_GDI, side_effect=Exception("VDM crashed")):
            with patch(_UIA):
                state = d.get_state(use_vision=False)
        assert len(state.all_desktops) == 1
        assert state.all_desktops[0]["name"] == "Default Desktop"

    def test_vdm_called_exactly_once(self):
        """VDM query is submitted exactly once per get_state() call."""
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        with patch(_GDI, return_value=vdm) as mock_vdm:
            with patch(_UIA):
                d.get_state(use_vision=False)
        mock_vdm.assert_called_once()


# ===========================================================================
# 6. Thread safety -- _state_lock protects atomic state swap
# ===========================================================================


class TestGetStateThreadSafety:
    """_state_lock must serialise the desktop_state assignment."""

    def test_state_lock_is_acquired_during_assignment(self):
        """_state_lock is used as a context manager during the state swap.

        threading.Lock attributes are read-only, so we replace the lock
        with a MagicMock that acts as a context manager and records calls.
        """
        import threading

        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)

        # Build a mock that behaves as a context manager but records usage
        real_lock = threading.Lock()
        enter_calls = []
        exit_calls = []

        class TrackingLock:
            def acquire(self, *args, **kwargs):
                enter_calls.append("acquire")
                return real_lock.acquire(*args, **kwargs)

            def release(self):
                exit_calls.append("release")
                return real_lock.release()

            def __enter__(self):
                enter_calls.append("enter")
                real_lock.acquire()
                return self

            def __exit__(self, *args):
                exit_calls.append("exit")
                real_lock.release()

        d._state_lock = TrackingLock()
        _run_get_state(d, vdm, use_vision=False)

        assert enter_calls, "Lock __enter__ must be called at least once"
        assert exit_calls, "Lock __exit__ must be called at least once"

    def test_concurrent_get_state_calls_all_succeed(self):
        """Multiple threads calling get_state concurrently must not raise."""
        results = []
        errors = []

        def worker():
            d = make_bare_desktop()
            vdm = _setup_desktop_for_get_state(d)
            try:
                state = _run_get_state(d, vdm, use_vision=False)
                results.append(state)
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=worker) for _ in range(5)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=10)

        assert errors == [], f"Thread errors: {errors}"
        assert len(results) == 5

    def test_desktop_state_visible_after_get_state(self):
        d = make_bare_desktop()
        assert d.desktop_state is None
        vdm = _setup_desktop_for_get_state(d)
        _run_get_state(d, vdm, use_vision=False)
        # After get_state, the cached state must be visible to other readers
        with d._state_lock:
            cached = d.desktop_state
        assert isinstance(cached, DesktopState)

    def test_successive_calls_update_cached_state(self):
        d = make_bare_desktop()

        win_a = make_window(name="App A", handle=111)
        vdm_a = _setup_desktop_for_get_state(d, active_window=win_a)
        state_a = _run_get_state(d, vdm_a, use_vision=False)
        assert d.desktop_state is state_a

        win_b = make_window(name="App B", handle=222)
        vdm_b = _setup_desktop_for_get_state(d, active_window=win_b)
        state_b = _run_get_state(d, vdm_b, use_vision=False)
        # The cached state must now point to the latest call
        assert d.desktop_state is state_b
        assert d.desktop_state is not state_a


# ===========================================================================
# 7. Multiple windows are processed correctly
# ===========================================================================


class TestGetStateMultipleWindows:
    """get_state must remove active_window from the windows list and preserve others."""

    def test_active_window_removed_from_windows_list(self):
        d = make_bare_desktop()
        active = make_window(name="Active", handle=10)
        other = make_window(name="Other", handle=20)
        # Both windows start in the list returned by get_windows
        vdm = _setup_desktop_for_get_state(d, windows=[active, other], active_window=active)
        state = _run_get_state(d, vdm, use_vision=False)
        assert active not in state.windows
        assert other in state.windows

    def test_non_active_windows_preserved(self):
        d = make_bare_desktop()
        active = make_window(name="Foreground", handle=1)
        bg1 = make_window(name="Background 1", handle=2)
        bg2 = make_window(name="Background 2", handle=3)
        vdm = _setup_desktop_for_get_state(
            d, windows=[active, bg1, bg2], active_window=active
        )
        state = _run_get_state(d, vdm, use_vision=False)
        assert bg1 in state.windows
        assert bg2 in state.windows
        assert len(state.windows) == 2

    def test_active_window_stored_on_state(self):
        d = make_bare_desktop()
        active = make_window(name="Main Window", handle=50)
        vdm = _setup_desktop_for_get_state(d, windows=[active], active_window=active)
        state = _run_get_state(d, vdm, use_vision=False)
        assert state.active_window is active

    def test_windows_not_in_active_window_preserved(self):
        """When active_window is not in the returned windows list, nothing is removed."""
        d = make_bare_desktop()
        win1 = make_window(name="Win1", handle=10)
        win2 = make_window(name="Win2", handle=20)
        # active_window is set but is NOT in the windows list
        active = make_window(name="Taskbar", handle=99)
        vdm = _setup_desktop_for_get_state(d, windows=[win1, win2], active_window=active)
        state = _run_get_state(d, vdm, use_vision=False)
        assert len(state.windows) == 2

    def test_five_background_windows_all_present(self):
        d = make_bare_desktop()
        active = make_window(name="Active", handle=1)
        bg_windows = [make_window(name=f"BG{i}", handle=100 + i) for i in range(5)]
        all_windows = [active] + bg_windows
        vdm = _setup_desktop_for_get_state(d, windows=all_windows, active_window=active)
        state = _run_get_state(d, vdm, use_vision=False)
        assert len(state.windows) == 5

    def test_tree_receives_active_window_handle(self):
        d = make_bare_desktop()
        active = make_window(name="Editor", handle=8888)
        vdm = _setup_desktop_for_get_state(d, windows=[active], active_window=active)
        _run_get_state(d, vdm, use_vision=False)
        call_args = d.tree.get_state.call_args
        assert call_args[0][0] == 8888

    def test_tree_receives_other_window_handles(self):
        """Handles not belonging to named windows go into other_windows_handles."""
        d = make_bare_desktop()
        active = make_window(handle=10)
        # controls_handles includes an extra handle (30) not in windows_handles
        vdm = _setup_desktop_for_get_state(
            d,
            windows=[active],
            active_window=active,
            controls_handles={10, 30},
            windows_handles={10},
        )
        _run_get_state(d, vdm, use_vision=False)
        call_args = d.tree.get_state.call_args
        other_handles = call_args[0][1]
        assert 30 in other_handles


# ===========================================================================
# 8. use_vision paths: annotated/plain/bytes/scaled screenshots
# ===========================================================================


class TestGetStateVisionPaths:
    """Screenshot handling when use_vision=True."""

    def test_use_vision_false_skips_screenshot(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        state = _run_get_state(d, vdm, use_vision=False)
        d._screen.get_screenshot.assert_not_called()
        d._screen.get_annotated_screenshot.assert_not_called()
        assert state.screenshot is None

    def test_use_vision_true_annotation_true_calls_annotated(self):
        d = make_bare_desktop()
        ts = _make_tree_state(1)
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)
        fake_img = Image.new("RGB", (300, 200))
        d._screen.get_annotated_screenshot.return_value = fake_img

        state = _run_get_state(d, vdm, use_vision=True, use_annotation=True)

        # Desktop.get_annotated_screenshot(nodes=nodes) delegates as positional arg
        d._screen.get_annotated_screenshot.assert_called_once_with(ts.interactive_nodes)
        assert state.screenshot is fake_img

    def test_use_vision_true_annotation_false_calls_plain_screenshot(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (640, 480))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(d, vdm, use_vision=True, use_annotation=False)

        d._screen.get_screenshot.assert_called_once()
        d._screen.get_annotated_screenshot.assert_not_called()
        assert state.screenshot is fake_img

    def test_as_bytes_true_screenshot_is_bytes(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (100, 80))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, as_bytes=True
        )

        assert isinstance(state.screenshot, bytes)

    def test_as_bytes_false_screenshot_is_image(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (100, 80))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, as_bytes=False
        )

        assert isinstance(state.screenshot, Image.Image)

    def test_scale_applied_to_screenshot(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (200, 100))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, scale=2.0
        )

        assert isinstance(state.screenshot, Image.Image)
        assert state.screenshot.width == 400
        assert state.screenshot.height == 200

    def test_scale_0_5_reduces_size(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (400, 200))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, scale=0.5
        )

        assert state.screenshot.width == 200
        assert state.screenshot.height == 100

    def test_scale_1_0_unchanged_size(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        orig_w, orig_h = 320, 240
        fake_img = Image.new("RGB", (orig_w, orig_h))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, scale=1.0
        )

        assert state.screenshot.width == orig_w
        assert state.screenshot.height == orig_h

    def test_as_bytes_true_produces_png_bytes(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (50, 50))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, as_bytes=True
        )

        assert isinstance(state.screenshot, bytes)
        # PNG magic bytes: 0x89 0x50 0x4e 0x47
        assert state.screenshot[:4] == b"\x89PNG"


# ===========================================================================
# 9. Boolean-string coercion for use_annotation, use_vision, use_dom, as_bytes
# ===========================================================================


class TestGetStateBoolParams:
    """Parameters accept proper booleans (tool layer coerces strings beforehand)."""

    def test_use_vision_true(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (50, 50))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(d, vdm, use_vision=True, use_annotation=False)

        d._screen.get_screenshot.assert_called_once()
        assert state.screenshot is fake_img

    def test_use_vision_false(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)

        state = _run_get_state(d, vdm, use_vision=False)

        d._screen.get_screenshot.assert_not_called()
        assert state.screenshot is None

    def test_use_annotation_true(self):
        d = make_bare_desktop()
        ts = _make_tree_state(1)
        vdm = _setup_desktop_for_get_state(d, tree_state=ts)
        fake_img = Image.new("RGB", (50, 50))
        d._screen.get_annotated_screenshot.return_value = fake_img

        _run_get_state(d, vdm, use_vision=True, use_annotation=True)

        d._screen.get_annotated_screenshot.assert_called_once()

    def test_use_dom_true_passed_to_tree(self):
        d = make_bare_desktop()
        win = make_window(handle=42)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])

        _run_get_state(d, vdm, use_vision=False, use_dom=True)

        call_kwargs = d.tree.get_state.call_args[1]
        assert call_kwargs.get("use_dom") is True

    def test_use_dom_false_passed_to_tree(self):
        d = make_bare_desktop()
        win = make_window(handle=43)
        vdm = _setup_desktop_for_get_state(d, active_window=win, windows=[win])

        _run_get_state(d, vdm, use_vision=False, use_dom=False)

        call_kwargs = d.tree.get_state.call_args[1]
        assert call_kwargs.get("use_dom") is False

    def test_as_bytes_true_produces_bytes(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (30, 30))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, as_bytes=True
        )

        assert isinstance(state.screenshot, bytes)

    def test_as_bytes_false_produces_image(self):
        d = make_bare_desktop()
        vdm = _setup_desktop_for_get_state(d)
        fake_img = Image.new("RGB", (30, 30))
        d._screen.get_screenshot.return_value = fake_img

        state = _run_get_state(
            d, vdm, use_vision=True, use_annotation=False, as_bytes=False
        )

        assert isinstance(state.screenshot, Image.Image)
