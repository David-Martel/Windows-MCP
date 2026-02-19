"""Tests for windows_mcp.native adapter module -- import-exception branch.

The native extension (windows_mcp_core.pyd) IS installed in this environment,
which means lines 26-28 (the success path) execute on every normal import.
Coverage gaps are lines 29-33 -- the ``except ImportError`` fallback branch:

    29: except ImportError:
    30:     windows_mcp_core = None
    31:     HAS_NATIVE = False
    32:     NATIVE_VERSION = None
    33:     logger.debug("Native extension not available ...")

Strategy
--------
We force the ``except ImportError`` branch to execute by setting
``sys.modules["windows_mcp_core"] = None`` before calling
``importlib.reload(windows_mcp.native)``.  Python raises ``ImportError``
for any module whose sys.modules entry is ``None``, which triggers the
except block.

Every test that reloads the module saves the original ``sys.modules`` entry
and the original module attributes, then restores them in a ``finally`` block
so the real extension is intact for all other tests.

The wrapper-function tests (Tests 2-6) use ``patch.object`` to control
``HAS_NATIVE`` at the attribute level without reloading, which is faster and
does not disturb the real extension.
"""

from __future__ import annotations

import importlib
import sys
from unittest.mock import MagicMock, patch

import windows_mcp.native as _native_module

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_CORE_MODULE_NAME = "windows_mcp_core"

# Sentinel to distinguish "key absent" from "key is None"
_SENTINEL = object()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_fake_core(version: str = "0.1.0") -> MagicMock:
    """Return a MagicMock that behaves like the real windows_mcp_core extension."""
    fake = MagicMock(name="windows_mcp_core")
    fake.__version__ = version
    fake.system_info.return_value = {
        "os_name": "Windows",
        "os_version": "10.0.19045",
        "hostname": "TESTPC",
        "cpu_count": 4,
        "cpu_usage_percent": [10.0, 20.0, 30.0, 40.0],
        "total_memory_bytes": 16 * 1024**3,
        "used_memory_bytes": 8 * 1024**3,
        "disks": [],
    }
    fake.capture_tree.return_value = []
    fake.send_text.return_value = 5
    fake.send_click.return_value = 2
    fake.send_key.return_value = 1
    fake.send_mouse_move.return_value = 1
    fake.send_hotkey.return_value = 2
    fake.send_scroll.return_value = 1
    return fake


class _ExtensionAbsent:
    """Context manager: force the ``except ImportError`` branch in native.py.

    Sets ``sys.modules["windows_mcp_core"] = None`` so that
    ``import windows_mcp_core`` raises ``ImportError`` during reload, then
    restores the original sys.modules entry and the real module attributes.

    Usage::

        with _ExtensionAbsent() as native_mod:
            assert native_mod.HAS_NATIVE is False
    """

    def __enter__(self):
        # Snapshot current state
        self._prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        self._prior_has_native = _native_module.HAS_NATIVE
        self._prior_version = _native_module.NATIVE_VERSION
        self._prior_core_attr = _native_module.windows_mcp_core

        # Force ImportError on next "import windows_mcp_core"
        sys.modules[_CORE_MODULE_NAME] = None  # type: ignore[assignment]
        importlib.reload(_native_module)
        return _native_module

    def __exit__(self, *_exc):
        # Restore sys.modules
        if self._prior_sys is _SENTINEL:
            sys.modules.pop(_CORE_MODULE_NAME, None)
        else:
            sys.modules[_CORE_MODULE_NAME] = self._prior_sys

        # Restore module attributes directly (faster than another reload)
        _native_module.HAS_NATIVE = self._prior_has_native
        _native_module.NATIVE_VERSION = self._prior_version
        _native_module.windows_mcp_core = self._prior_core_attr


class _FakeCoreInstalled:
    """Context manager: inject a fake extension as windows_mcp_core.

    Replaces the real extension with a MagicMock so wrapper functions can be
    exercised with controlled return values and side effects.

    Usage::

        with _FakeCoreInstalled(version="0.2.0") as native_mod:
            assert native_mod.HAS_NATIVE is True
            assert native_mod.NATIVE_VERSION == "0.2.0"
    """

    def __init__(self, version: str = "0.1.0") -> None:
        self._version = version

    def __enter__(self):
        self._prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        self._prior_has_native = _native_module.HAS_NATIVE
        self._prior_version = _native_module.NATIVE_VERSION
        self._prior_core_attr = _native_module.windows_mcp_core

        sys.modules[_CORE_MODULE_NAME] = _make_fake_core(self._version)
        importlib.reload(_native_module)
        return _native_module

    def __exit__(self, *_exc):
        if self._prior_sys is _SENTINEL:
            sys.modules.pop(_CORE_MODULE_NAME, None)
        else:
            sys.modules[_CORE_MODULE_NAME] = self._prior_sys

        _native_module.HAS_NATIVE = self._prior_has_native
        _native_module.NATIVE_VERSION = self._prior_version
        _native_module.windows_mcp_core = self._prior_core_attr


# ---------------------------------------------------------------------------
# Test 1: except ImportError branch sets HAS_NATIVE = False, NATIVE_VERSION = None
# ---------------------------------------------------------------------------


class TestImportExceptBranch:
    """Verify lines 29-33 of native.py execute when extension is missing.

    These are the currently uncovered lines that the task targets.
    """

    def test_has_native_false_when_import_fails(self):
        """HAS_NATIVE must be False when windows_mcp_core raises ImportError."""
        with _ExtensionAbsent() as native:
            assert native.HAS_NATIVE is False

    def test_native_version_none_when_import_fails(self):
        """NATIVE_VERSION must be None when the import fails."""
        with _ExtensionAbsent() as native:
            assert native.NATIVE_VERSION is None

    def test_windows_mcp_core_attribute_is_none_when_import_fails(self):
        """The module-level windows_mcp_core binding must be None."""
        with _ExtensionAbsent() as native:
            assert native.windows_mcp_core is None

    def test_state_restored_after_extension_absent_context(self):
        """After _ExtensionAbsent exits, HAS_NATIVE reverts to its prior value."""
        prior_has_native = _native_module.HAS_NATIVE
        with _ExtensionAbsent():
            assert _native_module.HAS_NATIVE is False
        assert _native_module.HAS_NATIVE == prior_has_native

    def test_native_version_restored_after_extension_absent_context(self):
        """NATIVE_VERSION reverts to its prior value after context exit."""
        prior_version = _native_module.NATIVE_VERSION
        with _ExtensionAbsent():
            assert _native_module.NATIVE_VERSION is None
        assert _native_module.NATIVE_VERSION == prior_version


# ---------------------------------------------------------------------------
# Test 1b: success path (lines 26-28) -- real extension is installed
# ---------------------------------------------------------------------------


class TestImportSuccessBranch:
    """Verify lines 26-28 execute when a real or fake extension is importable."""

    def test_has_native_true_with_real_extension(self):
        """The real extension is installed; HAS_NATIVE must be True."""
        assert _native_module.HAS_NATIVE is True

    def test_native_version_string_with_real_extension(self):
        """NATIVE_VERSION must be a non-empty string."""
        assert isinstance(_native_module.NATIVE_VERSION, str)
        assert _native_module.NATIVE_VERSION != ""

    def test_has_native_true_when_fake_extension_injected(self):
        """Injecting a fake extension also activates the success path."""
        with _FakeCoreInstalled(version="9.0.0") as native:
            assert native.HAS_NATIVE is True

    def test_native_version_from_fake_extension(self):
        """NATIVE_VERSION reflects the fake extension's __version__."""
        with _FakeCoreInstalled(version="9.0.0") as native:
            assert native.NATIVE_VERSION == "9.0.0"

    def test_windows_mcp_core_attribute_is_the_fake(self):
        """The module-level windows_mcp_core binding is the injected mock."""
        fake = _make_fake_core("2.0.0")
        prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        prior_has = _native_module.HAS_NATIVE
        prior_ver = _native_module.NATIVE_VERSION
        prior_attr = _native_module.windows_mcp_core
        try:
            sys.modules[_CORE_MODULE_NAME] = fake
            importlib.reload(_native_module)
            assert _native_module.windows_mcp_core is fake
        finally:
            if prior_sys is _SENTINEL:
                sys.modules.pop(_CORE_MODULE_NAME, None)
            else:
                sys.modules[_CORE_MODULE_NAME] = prior_sys
            _native_module.HAS_NATIVE = prior_has
            _native_module.NATIVE_VERSION = prior_ver
            _native_module.windows_mcp_core = prior_attr


# ---------------------------------------------------------------------------
# Test 2: native_system_info() returns dict when extension is present
# ---------------------------------------------------------------------------


class TestNativeSystemInfoWithExtension:
    """native_system_info() happy path -- extension IS available."""

    def test_returns_dict(self):
        """Result must be a dict when the fake extension succeeds."""
        with _FakeCoreInstalled() as native:
            result = native.native_system_info()
        assert isinstance(result, dict)

    def test_returns_dict_with_known_keys(self):
        """Returned dict must contain keys from the fake extension."""
        with _FakeCoreInstalled() as native:
            result = native.native_system_info()
        assert "os_name" in result
        assert "cpu_count" in result

    def test_delegates_to_extension_system_info(self):
        """native_system_info() must call windows_mcp_core.system_info()."""
        fake = _make_fake_core()
        prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        prior_has = _native_module.HAS_NATIVE
        prior_ver = _native_module.NATIVE_VERSION
        prior_attr = _native_module.windows_mcp_core
        try:
            sys.modules[_CORE_MODULE_NAME] = fake
            importlib.reload(_native_module)
            _native_module.native_system_info()
            fake.system_info.assert_called_once()
        finally:
            if prior_sys is _SENTINEL:
                sys.modules.pop(_CORE_MODULE_NAME, None)
            else:
                sys.modules[_CORE_MODULE_NAME] = prior_sys
            _native_module.HAS_NATIVE = prior_has
            _native_module.NATIVE_VERSION = prior_ver
            _native_module.windows_mcp_core = prior_attr


# ---------------------------------------------------------------------------
# Test 3: native_system_info() returns None and logs when extension raises
# ---------------------------------------------------------------------------


class TestNativeSystemInfoExceptionFallback:
    """native_system_info() must return None when the extension raises."""

    def test_returns_none_on_runtime_error(self):
        """RuntimeError from extension must be swallowed; None returned."""
        fake = _make_fake_core()
        fake.system_info.side_effect = RuntimeError("sysinfo_failure")
        prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        prior_has = _native_module.HAS_NATIVE
        prior_ver = _native_module.NATIVE_VERSION
        prior_attr = _native_module.windows_mcp_core
        try:
            sys.modules[_CORE_MODULE_NAME] = fake
            importlib.reload(_native_module)
            result = _native_module.native_system_info()
        finally:
            if prior_sys is _SENTINEL:
                sys.modules.pop(_CORE_MODULE_NAME, None)
            else:
                sys.modules[_CORE_MODULE_NAME] = prior_sys
            _native_module.HAS_NATIVE = prior_has
            _native_module.NATIVE_VERSION = prior_ver
            _native_module.windows_mcp_core = prior_attr
        assert result is None

    def test_returns_none_on_os_error(self):
        """OSError from extension must also be swallowed."""
        with patch.object(_native_module, "HAS_NATIVE", True), patch.object(
            _native_module, "windows_mcp_core", _make_fake_core()
        ) as mock_core:
            mock_core.system_info.side_effect = OSError("access denied")
            result = _native_module.native_system_info()
        assert result is None

    def test_logs_warning_on_exception(self, caplog):
        """A warning must be logged when system_info raises."""
        import logging

        fake = _make_fake_core()
        fake.system_info.side_effect = RuntimeError("boom")
        prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        prior_has = _native_module.HAS_NATIVE
        prior_ver = _native_module.NATIVE_VERSION
        prior_attr = _native_module.windows_mcp_core
        try:
            sys.modules[_CORE_MODULE_NAME] = fake
            importlib.reload(_native_module)
            with caplog.at_level(logging.WARNING, logger="windows_mcp.native"):
                _native_module.native_system_info()
        finally:
            if prior_sys is _SENTINEL:
                sys.modules.pop(_CORE_MODULE_NAME, None)
            else:
                sys.modules[_CORE_MODULE_NAME] = prior_sys
            _native_module.HAS_NATIVE = prior_has
            _native_module.NATIVE_VERSION = prior_ver
            _native_module.windows_mcp_core = prior_attr
        assert any("native_system_info failed" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# Test 4: all wrapper functions return None when HAS_NATIVE is False
# ---------------------------------------------------------------------------


class TestAllWrappersReturnNoneWhenNoExtension:
    """Every wrapper must short-circuit with None when HAS_NATIVE is False.

    Uses ``patch.object`` (no reload) -- fastest and safest approach since
    HAS_NATIVE is just a boolean attribute on the module.
    """

    def test_native_system_info_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_system_info() is None

    def test_native_capture_tree_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_capture_tree([1, 2, 3]) is None

    def test_native_send_text_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_text("hello") is None

    def test_native_send_click_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_click(100, 200) is None

    def test_native_send_key_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_key(0x0D) is None

    def test_native_send_mouse_move_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_mouse_move(300, 400) is None

    def test_native_send_hotkey_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_hotkey([0x11, 0x43]) is None

    def test_native_send_scroll_returns_none(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_scroll(0, 0, 120) is None

    def test_all_wrappers_return_none_in_single_pass(self):
        """Bulk assertion: every public wrapper returns None when disabled."""
        wrapper_calls = [
            lambda n: n.native_system_info(),
            lambda n: n.native_capture_tree([]),
            lambda n: n.native_send_text("x"),
            lambda n: n.native_send_click(0, 0),
            lambda n: n.native_send_key(0x41),
            lambda n: n.native_send_mouse_move(0, 0),
            lambda n: n.native_send_hotkey([0x11]),
            lambda n: n.native_send_scroll(0, 0, 120),
        ]
        with patch.object(_native_module, "HAS_NATIVE", False):
            for call in wrapper_calls:
                result = call(_native_module)
                assert result is None, f"Expected None from {call}, got {result!r}"

    def test_all_wrappers_none_via_extension_absent(self):
        """Simulate true absence via _ExtensionAbsent -- reload exercises lines 29-33."""
        with _ExtensionAbsent() as native:
            assert native.native_system_info() is None
            assert native.native_capture_tree([]) is None
            assert native.native_send_text("x") is None
            assert native.native_send_click(0, 0) is None
            assert native.native_send_key(0x41) is None
            assert native.native_send_mouse_move(0, 0) is None
            assert native.native_send_hotkey([0x11]) is None
            assert native.native_send_scroll(0, 0, 120) is None


# ---------------------------------------------------------------------------
# Test 5: native_send_hotkey() returns None on exception gracefully
# ---------------------------------------------------------------------------


class TestNativeSendHotkeyExceptionFallback:
    """native_send_hotkey() must return None on any extension exception."""

    def test_returns_none_on_runtime_error(self):
        with patch.object(_native_module, "HAS_NATIVE", True), patch.object(
            _native_module, "windows_mcp_core", _make_fake_core()
        ) as mock_core:
            mock_core.send_hotkey.side_effect = RuntimeError("hotkey_failure")
            result = _native_module.native_send_hotkey([0x11, 0x43])
        assert result is None

    def test_returns_none_on_value_error(self):
        with patch.object(_native_module, "HAS_NATIVE", True), patch.object(
            _native_module, "windows_mcp_core", _make_fake_core()
        ) as mock_core:
            mock_core.send_hotkey.side_effect = ValueError("bad vk_codes")
            result = _native_module.native_send_hotkey([])
        assert result is None

    def test_returns_none_when_has_native_false(self):
        """Short-circuit when HAS_NATIVE is False (no exception needed)."""
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_send_hotkey([0x11]) is None

    def test_logs_warning_on_hotkey_exception(self, caplog):
        """A warning must be emitted when send_hotkey raises."""
        import logging

        fake = _make_fake_core()
        fake.send_hotkey.side_effect = RuntimeError("hotkey exploded")
        prior_sys = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        prior_has = _native_module.HAS_NATIVE
        prior_ver = _native_module.NATIVE_VERSION
        prior_attr = _native_module.windows_mcp_core
        try:
            sys.modules[_CORE_MODULE_NAME] = fake
            importlib.reload(_native_module)
            with caplog.at_level(logging.WARNING, logger="windows_mcp.native"):
                _native_module.native_send_hotkey([0x43])
        finally:
            if prior_sys is _SENTINEL:
                sys.modules.pop(_CORE_MODULE_NAME, None)
            else:
                sys.modules[_CORE_MODULE_NAME] = prior_sys
            _native_module.HAS_NATIVE = prior_has
            _native_module.NATIVE_VERSION = prior_ver
            _native_module.windows_mcp_core = prior_attr
        assert any("native_send_hotkey failed" in r.message for r in caplog.records)

    def test_successful_hotkey_returns_count(self):
        """Verify the return value is forwarded when extension succeeds."""
        fake = _make_fake_core()
        fake.send_hotkey.return_value = 4
        with (
            patch.object(_native_module, "HAS_NATIVE", True),
            patch.object(_native_module, "windows_mcp_core", fake),
        ):
            result = _native_module.native_send_hotkey([0x11, 0x10, 0x53])
        assert result == 4
        fake.send_hotkey.assert_called_once_with([0x11, 0x10, 0x53])


# ---------------------------------------------------------------------------
# Test 6: native_capture_tree() returns None when extension not available
# ---------------------------------------------------------------------------


class TestNativeCaptureTreeNotAvailable:
    """native_capture_tree() returns None when HAS_NATIVE is False."""

    def test_returns_none_when_has_native_false(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_capture_tree([1, 2, 3]) is None

    def test_returns_none_with_empty_handles_when_disabled(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_capture_tree([]) is None

    def test_returns_none_with_custom_max_depth_when_disabled(self):
        with patch.object(_native_module, "HAS_NATIVE", False):
            assert _native_module.native_capture_tree([99], max_depth=10) is None

    def test_returns_none_via_extension_absent_reload(self):
        """Verify None result when lines 29-33 have executed (real absent path)."""
        with _ExtensionAbsent() as native:
            assert native.native_capture_tree([42]) is None

    def test_returns_list_when_extension_present(self):
        """Success path: extension returns a list, wrapper passes it through."""
        with _FakeCoreInstalled() as native:
            result = native.native_capture_tree([], max_depth=5)
        assert isinstance(result, list)

    def test_returns_none_on_capture_tree_exception(self):
        """Exception from extension is swallowed; None returned."""
        with patch.object(_native_module, "HAS_NATIVE", True), patch.object(
            _native_module, "windows_mcp_core", _make_fake_core()
        ) as mock_core:
            mock_core.capture_tree.side_effect = OSError("COM error")
            result = _native_module.native_capture_tree([1])
        assert result is None

    def test_extension_called_with_correct_args(self):
        """Wrapper must forward handles and max_depth to the extension."""
        fake = _make_fake_core()
        with (
            patch.object(_native_module, "HAS_NATIVE", True),
            patch.object(_native_module, "windows_mcp_core", fake),
        ):
            _native_module.native_capture_tree([42, 43], max_depth=7)
        fake.capture_tree.assert_called_once_with([42, 43], max_depth=7)


# ---------------------------------------------------------------------------
# Test 7: sys.modules hygiene -- _ExtensionAbsent restores state correctly
# ---------------------------------------------------------------------------


class TestSysModulesHygiene:
    """Verify that context manager teardown correctly restores state."""

    def test_prior_sys_modules_restored_after_extension_absent(self):
        """The original sys.modules[windows_mcp_core] is restored on exit."""
        prior = sys.modules.get(_CORE_MODULE_NAME, _SENTINEL)
        with _ExtensionAbsent():
            # Inside: the entry is None (forces ImportError)
            assert sys.modules.get(_CORE_MODULE_NAME) is None
        # Outside: restored to prior value
        if prior is _SENTINEL:
            assert _CORE_MODULE_NAME not in sys.modules
        else:
            assert sys.modules.get(_CORE_MODULE_NAME) is prior

    def test_has_native_restored_after_extension_absent(self):
        """HAS_NATIVE reverts to its pre-test value after context exit."""
        prior_has_native = _native_module.HAS_NATIVE
        with _ExtensionAbsent():
            assert _native_module.HAS_NATIVE is False
        assert _native_module.HAS_NATIVE == prior_has_native

    def test_native_version_restored_after_extension_absent(self):
        """NATIVE_VERSION reverts to its pre-test value after context exit."""
        prior_version = _native_module.NATIVE_VERSION
        with _ExtensionAbsent():
            assert _native_module.NATIVE_VERSION is None
        assert _native_module.NATIVE_VERSION == prior_version

    def test_fake_core_installed_restores_real_extension(self):
        """_FakeCoreInstalled restores the real extension module on exit."""
        with _FakeCoreInstalled(version="9.9.9"):
            assert _native_module.NATIVE_VERSION == "9.9.9"
        # After exit: version should be back to real extension's value
        assert _native_module.NATIVE_VERSION == _native_module.windows_mcp_core.__version__

    def test_nested_contexts_restore_correctly(self):
        """Nested _FakeCoreInstalled contexts restore to outer version."""
        with _FakeCoreInstalled(version="1.0.0") as native:
            assert native.NATIVE_VERSION == "1.0.0"
            with _FakeCoreInstalled(version="2.0.0") as inner_native:
                assert inner_native.NATIVE_VERSION == "2.0.0"
            # Inner context restored to outer fake version
            assert native.NATIVE_VERSION == "1.0.0"
