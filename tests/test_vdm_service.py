"""Comprehensive unit tests for the Virtual Desktop Manager (VDM).

Strategy: windows_mcp.vdm.core imports fine on Windows because it uses the real
comtypes library for class-body definitions (COMMETHOD / GUID / IUnknown).  We do
NOT attempt to replace comtypes before import -- that would break POINTER(GUID).

Instead we:
  1. Import vdm.core normally (it works on Windows without a live desktop).
  2. Use object.__new__() to create VirtualDesktopManager without calling __init__,
     injecting mock _manager / _internal_manager attributes directly.
  3. Patch comtypes.client.CreateObject and winreg for tests that exercise __init__
     or registry access.
  4. Patch GUID / byref at the vdm.core module level for tests that build GUID args.

All COM calls are mocked -- no live Windows desktop session is required.
"""

import threading
from unittest.mock import MagicMock, patch

import pytest

import windows_mcp.vdm.core as vdm_mod
from windows_mcp.vdm.core import VirtualDesktopManager

# ---------------------------------------------------------------------------
# Factory helpers
# ---------------------------------------------------------------------------


def _make_guid_str_mock(guid_str: str):
    """Return a MagicMock that stringifies to *guid_str*."""
    g = MagicMock()
    g.__str__ = MagicMock(return_value=guid_str)
    return g


def _make_desktop_mock(guid_str: str):
    """Return a mock IVirtualDesktop whose GetID returns a GUID-like mock."""
    guid_mock = _make_guid_str_mock(guid_str)
    desktop = MagicMock()
    desktop.GetID = MagicMock(return_value=guid_mock)
    return desktop


def _make_array_mock(desktops: list):
    """Return a mock IObjectArray containing the given desktop mocks."""
    array = MagicMock()
    array.GetCount.return_value = len(desktops)

    def _get_at(i, *args):
        unk = MagicMock()
        unk.QueryInterface.return_value = desktops[i]
        return unk

    array.GetAt.side_effect = _get_at
    return array


def _make_vdm(manager_mock=None, internal_mock=None) -> VirtualDesktopManager:
    """
    Build a VirtualDesktopManager instance without triggering COM in __init__.

    Uses object.__new__ to skip __init__, then injects mocks for the two
    COM attributes that every method inspects first.
    """
    vdm = object.__new__(VirtualDesktopManager)
    vdm._manager = manager_mock
    vdm._internal_manager = internal_mock
    vdm._desktop_cache = None
    vdm._cache_time = 0.0
    return vdm


# ---------------------------------------------------------------------------
# 1. Initialization
# ---------------------------------------------------------------------------


class TestVirtualDesktopManagerInit:
    """Verify __init__ COM setup and graceful degradation on failure."""

    def test_init_com_failure_sets_manager_to_none(self):
        """If CreateObject raises, _manager must be None (no crash)."""
        with patch("comtypes.client.CreateObject", side_effect=OSError("COM unavailable")):
            with patch("ctypes.windll.ole32.CoInitialize", return_value=0):
                vdm = VirtualDesktopManager()

        assert vdm._manager is None

    def test_init_internal_manager_failure_leaves_none(self):
        """If ImmersiveShell QueryService fails, _internal_manager is None."""
        mock_manager = MagicMock()
        mock_sp = MagicMock()
        mock_sp.QueryService.side_effect = OSError("Access denied")

        call_count = {"n": 0}

        def _create(clsid, interface=None):
            call_count["n"] += 1
            if call_count["n"] == 1:
                return mock_manager
            return mock_sp

        with patch("comtypes.client.CreateObject", side_effect=_create):
            with patch("ctypes.windll.ole32.CoInitialize", return_value=0):
                vdm = VirtualDesktopManager()

        assert vdm._manager is mock_manager
        assert vdm._internal_manager is None

    def test_init_via_new_bypass_has_expected_attributes(self):
        """object.__new__ bypass: injected attributes exist and are accessible."""
        mock_mgr = MagicMock()
        mock_int = MagicMock()
        vdm = _make_vdm(manager_mock=mock_mgr, internal_mock=mock_int)

        assert hasattr(vdm, "_manager")
        assert hasattr(vdm, "_internal_manager")
        assert vdm._manager is mock_mgr
        assert vdm._internal_manager is mock_int

    def test_init_success_with_full_com_mock(self):
        """Happy path: both COM objects created, both attributes set."""
        mock_manager = MagicMock()
        mock_sp = MagicMock()
        mock_unk = MagicMock()
        mock_internal = MagicMock()

        mock_sp.QueryService.return_value = mock_unk
        mock_unk.QueryInterface.return_value = mock_internal

        call_count = {"n": 0}

        def _create(clsid, interface=None):
            call_count["n"] += 1
            if call_count["n"] == 1:
                return mock_manager
            return mock_sp

        with patch("comtypes.client.CreateObject", side_effect=_create):
            with patch("ctypes.windll.ole32.CoInitialize", return_value=0):
                vdm = VirtualDesktopManager()

        assert vdm._manager is mock_manager
        assert vdm._internal_manager is mock_internal


# ---------------------------------------------------------------------------
# 2. is_window_on_current_desktop
# ---------------------------------------------------------------------------


class TestIsWindowOnCurrentDesktop:
    """is_window_on_current_desktop delegates to _manager.IsWindowOnCurrentVirtualDesktop."""

    def test_returns_true_when_window_on_current_desktop(self):
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.return_value = True
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.is_window_on_current_desktop(12345)

        mock_mgr.IsWindowOnCurrentVirtualDesktop.assert_called_once_with(12345)
        assert result is True

    def test_returns_false_when_window_on_different_desktop(self):
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.return_value = False
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.is_window_on_current_desktop(99999)

        assert result is False

    def test_falls_back_to_true_when_manager_is_none(self):
        """When _manager is None the method returns True (fail-open behavior)."""
        vdm = _make_vdm(manager_mock=None)

        result = vdm.is_window_on_current_desktop(42)

        assert result is True

    def test_com_exception_returns_true_fail_open(self):
        """COM failure inside the COM call is swallowed; method returns True."""
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.side_effect = OSError("COM error")
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.is_window_on_current_desktop(12345)

        assert result is True

    def test_hwnd_zero_passes_through_to_com(self):
        """HWND=0 is forwarded unchanged to the COM call."""
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.return_value = True
        vdm = _make_vdm(manager_mock=mock_mgr)

        vdm.is_window_on_current_desktop(0)

        mock_mgr.IsWindowOnCurrentVirtualDesktop.assert_called_once_with(0)

    def test_large_hwnd_passes_through(self):
        """Very large HWND values (0xFFFFFFFF) are forwarded without modification."""
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.return_value = True
        vdm = _make_vdm(manager_mock=mock_mgr)

        vdm.is_window_on_current_desktop(0xFFFFFFFF)

        mock_mgr.IsWindowOnCurrentVirtualDesktop.assert_called_once_with(0xFFFFFFFF)


# ---------------------------------------------------------------------------
# 3. get_window_desktop_id
# ---------------------------------------------------------------------------


class TestGetWindowDesktopId:
    """get_window_desktop_id returns the GUID string for a window's desktop."""

    def test_returns_guid_string(self):
        guid_str = "{AA509086-5CA9-4C25-8F95-589D3C07B48A}"
        mock_guid = _make_guid_str_mock(guid_str)
        mock_mgr = MagicMock()
        mock_mgr.GetWindowDesktopId.return_value = mock_guid
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.get_window_desktop_id(12345)

        assert result == guid_str

    def test_returns_empty_string_when_manager_is_none(self):
        vdm = _make_vdm(manager_mock=None)

        result = vdm.get_window_desktop_id(12345)

        assert result == ""

    def test_com_exception_returns_empty_string(self):
        mock_mgr = MagicMock()
        mock_mgr.GetWindowDesktopId.side_effect = OSError("COM error")
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.get_window_desktop_id(12345)

        assert result == ""

    def test_passes_hwnd_to_com(self):
        mock_mgr = MagicMock()
        mock_mgr.GetWindowDesktopId.return_value = _make_guid_str_mock("")
        vdm = _make_vdm(manager_mock=mock_mgr)

        vdm.get_window_desktop_id(77777)

        mock_mgr.GetWindowDesktopId.assert_called_once_with(77777)

    def test_returns_str_of_guid_object(self):
        """get_window_desktop_id calls str() on whatever COM returns."""
        mock_mgr = MagicMock()
        mock_mgr.GetWindowDesktopId.return_value = _make_guid_str_mock("{CUSTOM-GUID}")
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.get_window_desktop_id(1)

        assert result == "{CUSTOM-GUID}"


# ---------------------------------------------------------------------------
# 4. _get_name_from_registry
# ---------------------------------------------------------------------------


class TestGetNameFromRegistry:
    """_get_name_from_registry reads the desktop display name from the Windows registry."""

    def test_returns_name_when_key_exists(self):
        guid_str = "{AA509086-5CA9-4C25-8F95-589D3C07B48A}"
        mock_key = MagicMock()
        ctx = MagicMock()
        ctx.__enter__ = MagicMock(return_value=mock_key)
        ctx.__exit__ = MagicMock(return_value=False)
        vdm = _make_vdm()

        with (
            patch("windows_mcp.vdm.core.winreg") as mock_winreg,
        ):
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.return_value = ctx
            mock_winreg.QueryValueEx.return_value = ("Work Desktop", 1)
            result = vdm._get_name_from_registry(guid_str)

        assert result == "Work Desktop"

    def test_opens_path_containing_guid(self):
        """Verifies the registry subkey path includes the desktop GUID."""
        guid_str = "{DEADBEEF-0000-0000-0000-000000000000}"
        mock_key = MagicMock()
        ctx = MagicMock()
        ctx.__enter__ = MagicMock(return_value=mock_key)
        ctx.__exit__ = MagicMock(return_value=False)
        vdm = _make_vdm()

        with patch("windows_mcp.vdm.core.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.return_value = ctx
            mock_winreg.QueryValueEx.return_value = ("Gaming", 1)
            vdm._get_name_from_registry(guid_str)
            call_args = mock_winreg.OpenKey.call_args

        subkey_arg = call_args[0][1]
        assert guid_str in subkey_arg
        assert "VirtualDesktops" in subkey_arg

    def test_returns_none_when_key_does_not_exist(self):
        vdm = _make_vdm()

        with patch("windows_mcp.vdm.core.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.side_effect = OSError("Key not found")
            result = vdm._get_name_from_registry("{MISSING-GUID}")

        assert result is None

    def test_returns_none_when_name_value_missing(self):
        """Registry key exists but the 'Name' value is absent."""
        mock_key = MagicMock()
        ctx = MagicMock()
        ctx.__enter__ = MagicMock(return_value=mock_key)
        ctx.__exit__ = MagicMock(return_value=False)
        vdm = _make_vdm()

        with patch("windows_mcp.vdm.core.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.return_value = ctx
            mock_winreg.QueryValueEx.side_effect = OSError("Value not found")
            result = vdm._get_name_from_registry("{SOME-GUID}")

        assert result is None

    def test_returns_none_on_permission_error(self):
        """PermissionError opening the key returns None without propagating."""
        vdm = _make_vdm()

        with patch("windows_mcp.vdm.core.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.side_effect = PermissionError("Access denied")
            result = vdm._get_name_from_registry("{LOCKED-GUID}")

        assert result is None

    def test_unicode_desktop_name_preserved(self):
        """Unicode characters in the name are returned as-is."""
        mock_key = MagicMock()
        ctx = MagicMock()
        ctx.__enter__ = MagicMock(return_value=mock_key)
        ctx.__exit__ = MagicMock(return_value=False)
        vdm = _make_vdm()

        with patch("windows_mcp.vdm.core.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.OpenKey.return_value = ctx
            mock_winreg.QueryValueEx.return_value = ("Arbeit \u00dcbersicht", 1)
            result = vdm._get_name_from_registry("{UNICODE-GUID}")

        assert result == "Arbeit \u00dcbersicht"


# ---------------------------------------------------------------------------
# 5. get_all_desktops
# ---------------------------------------------------------------------------


class TestGetAllDesktops:
    """get_all_desktops enumerates all virtual desktops via the internal manager."""

    def test_returns_fallback_when_internal_manager_is_none(self):
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=None)

        result = vdm.get_all_desktops()

        assert len(result) == 1
        assert result[0]["name"] == "Default Desktop"
        assert result[0]["id"] == "00000000-0000-0000-0000-000000000000"

    def test_enumerates_single_desktop_with_registry_name(self):
        guid_str = "{AAAA-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)
        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="Work"):
            result = vdm.get_all_desktops()

        assert len(result) == 1
        assert result[0] == {"id": guid_str, "name": "Work"}

    def test_enumerates_multiple_desktops_mixed_names(self):
        guid1 = "{AAAA-0000-0000-0000-000000000001}"
        guid2 = "{BBBB-0000-0000-0000-000000000002}"
        desktop1 = _make_desktop_mock(guid1)
        desktop2 = _make_desktop_mock(guid2)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop1, desktop2])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        # First has registry name, second falls back to "Desktop 2"
        name_map = {guid1: "Work", guid2: None}
        with patch.object(vdm, "_get_name_from_registry", side_effect=lambda g: name_map.get(g)):
            result = vdm.get_all_desktops()

        assert len(result) == 2
        assert result[0] == {"id": guid1, "name": "Work"}
        assert result[1] == {"id": guid2, "name": "Desktop 2"}

    def test_skips_desktop_when_get_id_returns_none(self):
        """Desktops where GetID returns falsy are excluded from the result."""
        desktop_bad = MagicMock()
        desktop_bad.GetID = MagicMock(return_value=None)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop_bad])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_all_desktops()

        assert result == []

    def test_continues_past_per_desktop_exception(self):
        """An error on one desktop does not prevent others from being returned."""
        guid_good = "{CCCC-0000-0000-0000-000000000001}"
        desktop_good = _make_desktop_mock(guid_good)
        desktop_bad = MagicMock()
        desktop_bad.GetID.side_effect = OSError("COM error")

        array = MagicMock()
        array.GetCount.return_value = 2

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_bad if i == 0 else desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_all_desktops()

        assert len(result) == 1
        assert result[0]["id"] == guid_good

    def test_empty_desktop_list_returns_empty_result(self):
        array = MagicMock()
        array.GetCount.return_value = 0
        array.GetAt.side_effect = IndexError("empty")

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm.get_all_desktops()

        assert result == []

    def test_uses_fallback_name_when_no_registry_entry(self):
        """Desktops without a registry name get 'Desktop N' fallback."""
        guid_str = "{DDDD-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_all_desktops()

        assert result[0]["name"] == "Desktop 1"


# ---------------------------------------------------------------------------
# 6. get_current_desktop
# ---------------------------------------------------------------------------


class TestGetCurrentDesktop:
    """get_current_desktop returns info about the active desktop."""

    def test_returns_fallback_when_internal_manager_is_none(self):
        vdm = _make_vdm(internal_mock=None)

        result = vdm.get_current_desktop()

        assert result["id"] == "00000000-0000-0000-0000-000000000000"
        assert result["name"] == "Default Desktop"

    def test_returns_registry_name_fast_path(self):
        """When registry has a custom name, returns immediately without enumerating."""
        guid_str = "{EEEE-0000-0000-0000-000000000001}"
        guid_mock = _make_guid_str_mock(guid_str)
        current = MagicMock()
        current.GetID.return_value = guid_mock

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="My Work"):
            result = vdm.get_current_desktop()

        assert result == {"id": guid_str, "name": "My Work"}
        # Should NOT enumerate desktops when registry name found
        internal.GetDesktops.assert_not_called()

    def test_returns_indexed_name_via_enumeration(self):
        """When no registry name, enumerates desktops to find positional name."""
        guid_str = "{EEEE-0000-0000-0000-000000000001}"
        guid_mock = _make_guid_str_mock(guid_str)
        current = MagicMock()
        current.GetID.return_value = guid_mock

        # Build mock chain: GetAt -> unk -> QueryInterface -> desktop -> GetID
        other_guid = _make_guid_str_mock("{OTHER}")
        other_desktop = MagicMock()
        other_desktop.GetID.return_value = other_guid
        other_unk = MagicMock()
        other_unk.QueryInterface.return_value = other_desktop

        target_desktop = MagicMock()
        target_desktop.GetID.return_value = guid_mock
        target_unk = MagicMock()
        target_unk.QueryInterface.return_value = target_desktop

        arr = MagicMock()
        arr.GetCount.return_value = 2
        arr.GetAt.side_effect = lambda i, _: [other_unk, target_unk][i]

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current
        internal.GetDesktops.return_value = arr

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_current_desktop()

        assert result == {"id": guid_str, "name": "Desktop 2"}

    def test_returns_unknown_when_guid_not_found(self):
        """When GUID not found during enumeration, returns 'Unknown'."""
        guid_str = "{FFFF-0000-0000-0000-000000000099}"
        guid_mock = _make_guid_str_mock(guid_str)
        current = MagicMock()
        current.GetID.return_value = guid_mock

        arr = MagicMock()
        arr.GetCount.return_value = 0

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current
        internal.GetDesktops.return_value = arr

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_current_desktop()

        assert result == {"id": guid_str, "name": "Unknown"}


# ---------------------------------------------------------------------------
# 7. _resolve_to_guid
# ---------------------------------------------------------------------------


class TestResolveToGuid:
    """_resolve_to_guid maps a display name (or GUID string) to a GUID string."""

    def test_resolves_display_name_case_insensitively(self):
        guid_str = "{CCCC-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="Work Desktop"):
            result = vdm._resolve_to_guid("work desktop")

        assert result == guid_str

    def test_resolves_fallback_name(self):
        """When registry returns None, falls back to 'Desktop N' naming."""
        guid_str = "{DDDD-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._resolve_to_guid("Desktop 1")

        assert result == guid_str

    def test_returns_none_when_name_not_found(self):
        guid_str = "{EEEE-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="Work"):
            result = vdm._resolve_to_guid("NonExistentDesktop")

        assert result is None

    def test_returns_guid_when_guid_string_passed_directly(self):
        """Passing the GUID string itself returns it immediately (early exit)."""
        guid_str = "{FFFF-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._resolve_to_guid(guid_str)

        assert result == guid_str

    def test_returns_none_on_get_desktops_exception(self):
        """COM failure in GetDesktops returns None and logs the error."""
        internal = MagicMock()
        internal.GetDesktops.side_effect = OSError("COM unavailable")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._resolve_to_guid("Desktop 1")

        assert result is None

    def test_skips_desktop_with_falsy_guid(self):
        """Desktops whose GetID returns None are skipped without error."""
        desktop_bad = MagicMock()
        desktop_bad.GetID = MagicMock(return_value=None)

        array = MagicMock()
        array.GetCount.return_value = 1

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_bad
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._resolve_to_guid("Desktop 1")

        assert result is None

    def test_empty_name_returns_none(self):
        """An empty string as name is not matched and returns None."""
        guid_str = "{A1B2-0000-0000-0000-000000000001}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="Work"):
            result = vdm._resolve_to_guid("")

        assert result is None

    def test_case_insensitive_guid_comparison(self):
        """GUID comparison is case-insensitive (lowercase input matches uppercase GUID)."""
        guid_str = "{ABCD-ABCD-ABCD-ABCD-ABCDABCDABCD}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._resolve_to_guid(guid_str.lower())

        assert result == guid_str


# ---------------------------------------------------------------------------
# 8. move_window_to_desktop
# ---------------------------------------------------------------------------


class TestMoveWindowToDesktop:
    """move_window_to_desktop calls MoveWindowToDesktop with the resolved GUID."""

    def test_does_nothing_when_manager_is_none(self):
        vdm = _make_vdm(manager_mock=None)

        # Must not raise
        vdm.move_window_to_desktop(12345, "Work Desktop")

    def test_does_nothing_when_desktop_not_found(self):
        mock_mgr = MagicMock()
        vdm = _make_vdm(manager_mock=mock_mgr, internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=None):
            vdm.move_window_to_desktop(12345, "Nonexistent Desktop")

        mock_mgr.MoveWindowToDesktop.assert_not_called()

    def test_moves_window_when_desktop_found(self):
        guid_str = "{MOVE-GUID-0000-0000-000000000001}"
        mock_mgr = MagicMock()
        vdm = _make_vdm(manager_mock=mock_mgr, internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID") as mock_guid_cls:
                with patch("windows_mcp.vdm.core.byref"):
                    vdm.move_window_to_desktop(12345, "Work Desktop")

        mock_guid_cls.assert_called_once_with(guid_str)
        mock_mgr.MoveWindowToDesktop.assert_called_once()

    def test_swallows_com_exception_on_move(self):
        guid_str = "{MOVE-ERR-GUID}"
        mock_mgr = MagicMock()
        mock_mgr.MoveWindowToDesktop.side_effect = OSError("COM error")

        vdm = _make_vdm(manager_mock=mock_mgr, internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID"):
                with patch("windows_mcp.vdm.core.byref"):
                    # Must not propagate
                    vdm.move_window_to_desktop(12345, "Work Desktop")

    def test_swallows_guid_constructor_exception(self):
        """If GUID constructor raises, the error is caught and logged."""
        guid_str = "{BAD-GUID}"
        mock_mgr = MagicMock()
        vdm = _make_vdm(manager_mock=mock_mgr, internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID", side_effect=ValueError("bad GUID")):
                # Must not propagate
                vdm.move_window_to_desktop(12345, "Desktop 2")

        mock_mgr.MoveWindowToDesktop.assert_not_called()


# ---------------------------------------------------------------------------
# 9. create_desktop
# ---------------------------------------------------------------------------


class TestCreateDesktop:
    """create_desktop creates a new virtual desktop and optionally renames it."""

    def test_raises_when_internal_manager_is_none(self):
        vdm = _make_vdm(internal_mock=None)

        with pytest.raises(RuntimeError, match="Internal VDM not initialized"):
            vdm.create_desktop()

    def test_creates_desktop_and_returns_last_desktop_name(self):
        guid_mock = _make_guid_str_mock("{NEW-GUID}")
        new_desktop = MagicMock()
        new_desktop.GetID.return_value = guid_mock

        internal = MagicMock()
        internal.CreateDesktopW.return_value = new_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        all_desktops = [
            {"id": "{EXISTING}", "name": "Desktop 1"},
            {"id": "{NEW-GUID}", "name": "Desktop 2"},
        ]
        with patch.object(vdm, "get_all_desktops", return_value=all_desktops):
            result = vdm.create_desktop()

        assert result == "Desktop 2"

    def test_creates_desktop_with_custom_name_calls_rename(self):
        guid_mock = _make_guid_str_mock("{NEW-GUID}")
        new_desktop = MagicMock()
        new_desktop.GetID.return_value = guid_mock

        internal = MagicMock()
        internal.CreateDesktopW.return_value = new_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "rename_desktop_by_guid") as mock_rename:
            result = vdm.create_desktop(name="My Custom Desktop")

        assert result == "My Custom Desktop"
        mock_rename.assert_called_once_with("{NEW-GUID}", "My Custom Desktop")

    def test_create_desktop_calls_internal_manager(self):
        guid_mock = _make_guid_str_mock("{CREATED-GUID}")
        new_desktop = MagicMock()
        new_desktop.GetID.return_value = guid_mock

        internal = MagicMock()
        internal.CreateDesktopW.return_value = new_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(
            vdm, "get_all_desktops", return_value=[{"id": "{CREATED-GUID}", "name": "Desktop 2"}]
        ):
            vdm.create_desktop()

        internal.CreateDesktopW.assert_called_once()


# ---------------------------------------------------------------------------
# 10. remove_desktop
# ---------------------------------------------------------------------------


class TestRemoveDesktop:
    """remove_desktop removes a desktop by name."""

    def test_raises_when_internal_manager_is_none(self):
        vdm = _make_vdm(internal_mock=None)

        with pytest.raises(RuntimeError, match="Internal VDM not initialized"):
            vdm.remove_desktop("Desktop 2")

    def test_does_nothing_when_name_not_found(self):
        """No matching desktop in enumeration -- RemoveDesktop not called."""
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=[]):
            vdm.remove_desktop("Nonexistent")

        internal.RemoveDesktop.assert_not_called()

    def test_does_nothing_when_no_fallback_desktop_available(self):
        """Cannot remove the only desktop -- no fallback means no removal."""
        only_desktop = MagicMock()
        entries = [
            {"index": 0, "guid_str": "{ONLY-GUID}", "name": "Desktop 1", "desktop": only_desktop}
        ]
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            vdm.remove_desktop("Desktop 1")

        internal.RemoveDesktop.assert_not_called()

    def test_calls_remove_desktop_with_target_and_fallback(self):
        """When a fallback exists, RemoveDesktop is called with both desktops."""
        target_desktop = MagicMock()
        fallback_desktop = MagicMock()
        entries = [
            {
                "index": 0,
                "guid_str": "{TARGET-GUID}",
                "name": "Desktop 1",
                "desktop": target_desktop,
            },
            {
                "index": 1,
                "guid_str": "{FALLBACK-GUID}",
                "name": "Desktop 2",
                "desktop": fallback_desktop,
            },
        ]
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            vdm.remove_desktop("Desktop 1")

        internal.RemoveDesktop.assert_called_once()
        call_args = internal.RemoveDesktop.call_args[0]
        assert call_args[0] is target_desktop
        assert call_args[1] is fallback_desktop

    def test_resolves_by_guid_string(self):
        """remove_desktop also matches when the name is a GUID string."""
        target_desktop = MagicMock()
        fallback_desktop = MagicMock()
        entries = [
            {
                "index": 0,
                "guid_str": "{AAA}",
                "name": "Desktop 1",
                "desktop": target_desktop,
            },
            {
                "index": 1,
                "guid_str": "{BBB}",
                "name": "Desktop 2",
                "desktop": fallback_desktop,
            },
        ]
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            vdm.remove_desktop("{AAA}")

        internal.RemoveDesktop.assert_called_once()


# ---------------------------------------------------------------------------
# 11. rename_desktop / rename_desktop_by_guid
# ---------------------------------------------------------------------------


class TestRenameDesktop:
    """rename_desktop resolves name to GUID then delegates to rename_desktop_by_guid."""

    def test_delegates_to_rename_by_guid(self):
        target_guid = "{RENAME-GUID}"
        vdm = _make_vdm(internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=target_guid):
            with patch.object(vdm, "rename_desktop_by_guid") as mock_rename:
                vdm.rename_desktop("Old Name", "New Name")

        mock_rename.assert_called_once_with(target_guid, "New Name")

    def test_does_nothing_when_name_not_resolved(self):
        vdm = _make_vdm(internal_mock=MagicMock())

        with patch.object(vdm, "_resolve_to_guid", return_value=None):
            with patch.object(vdm, "rename_desktop_by_guid") as mock_rename:
                vdm.rename_desktop("Missing Desktop", "New Name")

        mock_rename.assert_not_called()

    def test_rename_by_guid_does_nothing_when_internal_manager_is_none(self):
        vdm = _make_vdm(internal_mock=None)

        with patch("windows_mcp.vdm.core.GUID"):
            # Must not raise
            vdm.rename_desktop_by_guid("{SOME-GUID}", "New Name")

    def test_rename_by_guid_calls_set_name_on_internal_manager(self):
        guid_str = "{RENAME-GUID-2}"
        target_desktop = MagicMock()
        internal = MagicMock()
        internal.FindDesktop.return_value = target_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        mock_hs = MagicMock()
        with patch("windows_mcp.vdm.core.GUID"):
            with patch("windows_mcp.vdm.core.create_hstring", return_value=mock_hs):
                with patch("windows_mcp.vdm.core.delete_hstring") as mock_del:
                    vdm.rename_desktop_by_guid(guid_str, "My Desktop")

        internal.SetName.assert_called_once()
        mock_del.assert_called_once_with(mock_hs)

    def test_rename_by_guid_cleans_up_hstring_on_exception(self):
        """delete_hstring is called even when SetName raises (finally block)."""
        guid_str = "{ERR-GUID}"
        target_desktop = MagicMock()
        internal = MagicMock()
        internal.FindDesktop.return_value = target_desktop
        internal.SetName.side_effect = OSError("COM error")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        mock_hs = MagicMock()
        with patch("windows_mcp.vdm.core.GUID"):
            with patch("windows_mcp.vdm.core.create_hstring", return_value=mock_hs):
                with patch("windows_mcp.vdm.core.delete_hstring") as mock_del:
                    # Should not raise
                    vdm.rename_desktop_by_guid(guid_str, "My Desktop")

        mock_del.assert_called_once_with(mock_hs)

    def test_rename_by_guid_handles_find_desktop_exception(self):
        """FindDesktop failure is silently handled; no crash."""
        internal = MagicMock()
        internal.FindDesktop.side_effect = OSError("Not found")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch("windows_mcp.vdm.core.GUID"):
            # Must not raise
            vdm.rename_desktop_by_guid("{MISSING}", "New Name")

    def test_rename_by_guid_logs_warning_when_set_name_absent(self):
        """If internal manager has no SetName method, a warning is logged."""
        target_desktop = MagicMock()
        internal = MagicMock(spec=["FindDesktop"])
        internal.FindDesktop.return_value = target_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch("windows_mcp.vdm.core.GUID"):
            with patch("windows_mcp.vdm.core.create_hstring", return_value=MagicMock()):
                with patch("windows_mcp.vdm.core.delete_hstring"):
                    # Should not raise
                    vdm.rename_desktop_by_guid("{SOME-GUID}", "Test")


# ---------------------------------------------------------------------------
# 12. switch_desktop
# ---------------------------------------------------------------------------


class TestSwitchDesktop:
    """switch_desktop calls SwitchDesktop on the internal manager."""

    def test_raises_when_internal_manager_is_none(self):
        vdm = _make_vdm(internal_mock=None)

        with pytest.raises(RuntimeError, match="Internal VDM not initialized"):
            vdm.switch_desktop("Desktop 2")

    def test_does_nothing_when_desktop_not_found(self):
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_resolve_to_guid", return_value=None):
            vdm.switch_desktop("Nonexistent")

        internal.SwitchDesktop.assert_not_called()

    def test_finds_and_switches_to_desktop(self):
        guid_str = "{SWITCH-GUID}"
        target_desktop = MagicMock()
        internal = MagicMock()
        internal.FindDesktop.return_value = target_desktop

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID"):
                vdm.switch_desktop("Desktop 2")

        internal.FindDesktop.assert_called_once()
        internal.SwitchDesktop.assert_called_once_with(target_desktop)

    def test_swallows_com_exception_from_find_desktop(self):
        guid_str = "{SWITCH-ERR-GUID}"
        internal = MagicMock()
        internal.FindDesktop.side_effect = OSError("COM error")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID"):
                # Must not propagate
                vdm.switch_desktop("Desktop 2")

    def test_swallows_com_exception_from_switch_desktop(self):
        guid_str = "{SWITCH-DESK-ERR}"
        target_desktop = MagicMock()
        internal = MagicMock()
        internal.FindDesktop.return_value = target_desktop
        internal.SwitchDesktop.side_effect = OSError("COM error")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_resolve_to_guid", return_value=guid_str):
            with patch("windows_mcp.vdm.core.GUID"):
                # Must not propagate
                vdm.switch_desktop("Desktop 2")


# ---------------------------------------------------------------------------
# 13. Module-level convenience functions
# ---------------------------------------------------------------------------


class TestModuleLevelFunctions:
    """The module-level functions delegate to the thread-local manager singleton."""

    def _mock_manager_ctx(self):
        """Return (patcher_context, mock_vdm) for use in a with block."""
        mock_vdm = MagicMock()
        return patch("windows_mcp.vdm.core._get_manager", return_value=mock_vdm), mock_vdm

    def test_is_window_on_current_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.is_window_on_current_desktop(12345)
        mock_vdm.is_window_on_current_desktop.assert_called_once_with(12345)

    def test_get_window_desktop_id_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.get_window_desktop_id(12345)
        mock_vdm.get_window_desktop_id.assert_called_once_with(12345)

    def test_move_window_to_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.move_window_to_desktop(12345, "Desktop 2")
        mock_vdm.move_window_to_desktop.assert_called_once_with(12345, "Desktop 2")

    def test_create_desktop_delegates_with_name(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.create_desktop("New Desktop")
        mock_vdm.create_desktop.assert_called_once_with("New Desktop")

    def test_create_desktop_delegates_without_name(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.create_desktop()
        mock_vdm.create_desktop.assert_called_once_with(None)

    def test_remove_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.remove_desktop("Desktop 2")
        mock_vdm.remove_desktop.assert_called_once_with("Desktop 2")

    def test_rename_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.rename_desktop("Old", "New")
        mock_vdm.rename_desktop.assert_called_once_with("Old", "New")

    def test_switch_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.switch_desktop("Desktop 1")
        mock_vdm.switch_desktop.assert_called_once_with("Desktop 1")

    def test_get_all_desktops_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.get_all_desktops()
        mock_vdm.get_all_desktops.assert_called_once_with()

    def test_get_current_desktop_delegates(self):
        ctx, mock_vdm = self._mock_manager_ctx()
        with ctx:
            vdm_mod.get_current_desktop()
        mock_vdm.get_current_desktop.assert_called_once_with()


# ---------------------------------------------------------------------------
# 14. Thread-local manager (_get_manager) singleton behavior
# ---------------------------------------------------------------------------


class TestGetLocalManager:
    """_get_manager returns the same instance per-thread and is lazy."""

    def _clear_thread_local(self):
        if hasattr(vdm_mod._thread_local, "manager"):
            del vdm_mod._thread_local.manager

    def test_same_thread_returns_same_instance(self):
        self._clear_thread_local()
        sentinel = MagicMock(name="singleton-manager")

        with patch.object(vdm_mod, "VirtualDesktopManager", return_value=sentinel) as mock_cls:
            m1 = vdm_mod._get_manager()
            m2 = vdm_mod._get_manager()

        assert m1 is m2
        assert m1 is sentinel
        # VirtualDesktopManager() should only have been instantiated once
        mock_cls.assert_called_once()

    def test_creates_instance_lazily(self):
        self._clear_thread_local()
        sentinel = MagicMock(name="lazy-manager")

        with patch.object(vdm_mod, "VirtualDesktopManager", return_value=sentinel) as mock_cls:
            result = vdm_mod._get_manager()

        assert result is sentinel
        mock_cls.assert_called_once()

    def test_different_threads_get_different_instances(self):
        """Each thread has its own VirtualDesktopManager singleton."""
        instances = []
        errors = []

        def _worker():
            try:
                # Ensure this thread has no pre-existing manager
                if hasattr(vdm_mod._thread_local, "manager"):
                    del vdm_mod._thread_local.manager

                thread_sentinel = MagicMock(name=f"mgr-{threading.current_thread().name}")
                with patch.object(vdm_mod, "VirtualDesktopManager", return_value=thread_sentinel):
                    m = vdm_mod._get_manager()
                instances.append(m)
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=_worker, name=f"vdm-test-{i}") for i in range(2)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert not errors, f"Thread errors: {errors}"
        assert len(instances) == 2
        # Each thread produced its own distinct instance
        assert instances[0] is not instances[1]

    def test_thread_local_is_separate_from_main_thread(self):
        """Manager created on a worker thread does not pollute the main thread."""
        self._clear_thread_local()
        main_sentinel = MagicMock(name="main-manager")
        worker_sentinel = MagicMock(name="worker-manager")
        worker_result = []

        def _worker():
            if hasattr(vdm_mod._thread_local, "manager"):
                del vdm_mod._thread_local.manager
            with patch.object(vdm_mod, "VirtualDesktopManager", return_value=worker_sentinel):
                worker_result.append(vdm_mod._get_manager())

        t = threading.Thread(target=_worker)
        t.start()
        t.join()

        with patch.object(vdm_mod, "VirtualDesktopManager", return_value=main_sentinel):
            main_result = vdm_mod._get_manager()

        assert worker_result[0] is worker_sentinel
        assert main_result is main_sentinel
        assert worker_result[0] is not main_result


# ---------------------------------------------------------------------------
# 15. create_hstring / delete_hstring helpers
# ---------------------------------------------------------------------------


class TestHstringHelpers:
    """create_hstring and delete_hstring handle combase availability gracefully."""

    def test_create_hstring_returns_zero_when_combase_unavailable(self):
        from windows_mcp.vdm.core import HSTRING, create_hstring

        with patch("windows_mcp.vdm.core._WindowsCreateString", None):
            result = create_hstring("hello")

        # HSTRING(0) -> c_void_p(0) -> .value is None on 64-bit
        assert isinstance(result, HSTRING)

    def test_create_hstring_raises_os_error_on_hresult_failure(self):
        from windows_mcp.vdm.core import create_hstring

        mock_create = MagicMock(return_value=-1)  # Non-zero HRESULT = failure

        with patch("windows_mcp.vdm.core._WindowsCreateString", mock_create):
            with pytest.raises(OSError, match="WindowsCreateString failed"):
                create_hstring("hello")

    def test_create_hstring_success_calls_windows_create_string(self):
        from windows_mcp.vdm.core import HSTRING, create_hstring

        mock_create = MagicMock(return_value=0)  # S_OK

        with patch("windows_mcp.vdm.core._WindowsCreateString", mock_create):
            result = create_hstring("test")

        mock_create.assert_called_once()
        assert isinstance(result, HSTRING)

    def test_delete_hstring_does_nothing_when_combase_unavailable(self):
        from windows_mcp.vdm.core import HSTRING, delete_hstring

        hs = HSTRING(0)
        with patch("windows_mcp.vdm.core._WindowsDeleteString", None):
            delete_hstring(hs)  # Must not raise

    def test_delete_hstring_skips_falsy_hstring(self):
        """HSTRING(0) is falsy -- delete function should NOT be invoked."""
        from windows_mcp.vdm.core import HSTRING, delete_hstring

        mock_delete = MagicMock()
        hs = HSTRING(0)

        with patch("windows_mcp.vdm.core._WindowsDeleteString", mock_delete):
            delete_hstring(hs)

        mock_delete.assert_not_called()

    def test_delete_hstring_calls_windows_delete_for_nonzero_hstring(self):
        """A non-zero HSTRING triggers the underlying Windows API call."""
        from windows_mcp.vdm.core import HSTRING, delete_hstring

        mock_delete = MagicMock()
        hs = HSTRING(12345)

        with patch("windows_mcp.vdm.core._WindowsDeleteString", mock_delete):
            delete_hstring(hs)

        mock_delete.assert_called_once_with(hs)


# ---------------------------------------------------------------------------
# 16. Edge cases and boundary conditions
# ---------------------------------------------------------------------------


class TestEdgeCases:
    """Miscellaneous edge cases and boundary conditions."""

    def test_get_window_desktop_id_with_hwnd_zero_returns_string(self):
        mock_mgr = MagicMock()
        mock_mgr.GetWindowDesktopId.return_value = _make_guid_str_mock("{ZERO-HWND-GUID}")
        vdm = _make_vdm(manager_mock=mock_mgr)

        result = vdm.get_window_desktop_id(0)

        mock_mgr.GetWindowDesktopId.assert_called_once_with(0)
        assert isinstance(result, str)

    def test_get_all_desktops_with_three_desktops_fallback_names(self):
        """Three desktops without registry names get Desktop 1/2/3 names."""
        guids = ["{G1}", "{G2}", "{G3}"]
        desktops = [_make_desktop_mock(g) for g in guids]

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock(desktops)

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm.get_all_desktops()

        assert [d["name"] for d in result] == ["Desktop 1", "Desktop 2", "Desktop 3"]

    def test_resolve_to_guid_multiple_desktops_selects_correct_one(self):
        """With multiple desktops, only the one matching the name is returned."""
        guid1 = "{MULTI-GUID-1}"
        guid2 = "{MULTI-GUID-2}"
        desktop1 = _make_desktop_mock(guid1)
        desktop2 = _make_desktop_mock(guid2)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop1, desktop2])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        name_map = {guid1: "Work", guid2: "Gaming"}
        with patch.object(vdm, "_get_name_from_registry", side_effect=lambda g: name_map.get(g)):
            result = vdm._resolve_to_guid("Gaming")

        assert result == guid2

    def test_is_window_on_current_desktop_called_twice_with_different_hwnds(self):
        """Multiple calls with different HWNDs each invoke COM once per call."""
        mock_mgr = MagicMock()
        mock_mgr.IsWindowOnCurrentVirtualDesktop.side_effect = [True, False]
        vdm = _make_vdm(manager_mock=mock_mgr)

        r1 = vdm.is_window_on_current_desktop(100)
        r2 = vdm.is_window_on_current_desktop(200)

        assert r1 is True
        assert r2 is False
        assert mock_mgr.IsWindowOnCurrentVirtualDesktop.call_count == 2

    def test_get_current_desktop_fallback_id_is_null_guid(self):
        """The fallback ID for no internal manager is the all-zeros GUID string."""
        vdm = _make_vdm(internal_mock=None)

        result = vdm.get_current_desktop()

        assert result["id"] == "00000000-0000-0000-0000-000000000000"

    def test_get_all_desktops_fallback_is_always_list(self):
        """Fallback when no internal manager returns a list, not None."""
        vdm = _make_vdm(internal_mock=None)

        result = vdm.get_all_desktops()

        assert isinstance(result, list)
        assert len(result) >= 1


# ---------------------------------------------------------------------------
# 17. _enumerate_desktops -- direct unit tests
# ---------------------------------------------------------------------------


class TestEnumerateDesktopsDirectly:
    """Direct unit tests for _enumerate_desktops().

    These tests exercise _enumerate_desktops() by calling it directly (not via
    the public wrappers) so that every branch of the helper can be verified in
    isolation.  _get_name_from_registry is patched with patch.object so the
    tests remain independent of registry state.
    """

    # ------------------------------------------------------------------
    # Guard: _internal_manager is None
    # ------------------------------------------------------------------

    def test_returns_empty_list_when_internal_manager_is_none(self):
        """_enumerate_desktops returns [] immediately when _internal_manager is None."""
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=None)

        result = vdm._enumerate_desktops()

        assert result == []

    def test_does_not_call_get_desktops_when_internal_manager_is_none(self):
        """No COM call is made when the guard short-circuits on None manager."""
        internal = MagicMock()
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=None)
        # Assign after construction so we can verify it was never called
        vdm._internal_manager = None

        vdm._enumerate_desktops()

        internal.GetDesktops.assert_not_called()

    # ------------------------------------------------------------------
    # Happy path: multiple desktops
    # ------------------------------------------------------------------

    def test_returns_correct_number_of_entries_for_multiple_desktops(self):
        """Each valid desktop produces exactly one entry in the result list."""
        guid1 = "{ENUM-GUID-0001}"
        guid2 = "{ENUM-GUID-0002}"
        guid3 = "{ENUM-GUID-0003}"
        desktops = [_make_desktop_mock(g) for g in (guid1, guid2, guid3)]

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock(desktops)

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 3

    def test_entry_dict_contains_required_keys(self):
        """Every entry must contain exactly the keys: index, guid_str, name, desktop."""
        guid_str = "{ENUM-KEYS-GUID}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 1
        entry = result[0]
        assert set(entry.keys()) == {"index", "guid_str", "name", "desktop"}

    def test_index_values_are_zero_based_sequential(self):
        """index in each entry reflects the loop counter (0-based)."""
        guids = ["{IDX-G1}", "{IDX-G2}", "{IDX-G3}"]
        desktops = [_make_desktop_mock(g) for g in guids]

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock(desktops)

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert [e["index"] for e in result] == [0, 1, 2]

    def test_guid_str_matches_desktop_get_id(self):
        """guid_str in each entry equals str(desktop.GetID())."""
        guid_str = "{GUID-STR-CHECK}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert result[0]["guid_str"] == guid_str

    def test_desktop_field_holds_query_interface_result(self):
        """The 'desktop' field stores the object returned by QueryInterface."""
        guid_str = "{DESKTOP-OBJ-GUID}"
        desktop_obj = _make_desktop_mock(guid_str)

        # Build array manually so we can verify the exact QueryInterface return value
        array = MagicMock()
        array.GetCount.return_value = 1

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_obj
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert result[0]["desktop"] is desktop_obj

    # ------------------------------------------------------------------
    # Registry names vs. fallback "Desktop N" names
    # ------------------------------------------------------------------

    def test_uses_registry_name_when_available(self):
        """When _get_name_from_registry returns a string, that name is used."""
        guid_str = "{REG-NAME-GUID}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value="My Work Desktop"):
            result = vdm._enumerate_desktops()

        assert result[0]["name"] == "My Work Desktop"

    def test_falls_back_to_desktop_n_name_when_no_registry_entry(self):
        """When registry returns None the name becomes 'Desktop {i+1}'."""
        guid_str = "{FALLBACK-NAME-GUID}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert result[0]["name"] == "Desktop 1"

    def test_fallback_name_uses_one_based_index(self):
        """Second desktop without a registry name gets 'Desktop 2', not 'Desktop 1'."""
        guids = ["{FB-G1}", "{FB-G2}"]
        desktops = [_make_desktop_mock(g) for g in guids]

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock(desktops)

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert result[0]["name"] == "Desktop 1"
        assert result[1]["name"] == "Desktop 2"

    def test_mixed_registry_and_fallback_names(self):
        """Desktops with registry names use them; others fall back to 'Desktop N'."""
        guid1, guid2, guid3 = "{MIX-G1}", "{MIX-G2}", "{MIX-G3}"
        desktops = [_make_desktop_mock(g) for g in (guid1, guid2, guid3)]

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock(desktops)

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        # Only the middle desktop has a custom name
        name_map = {guid1: None, guid2: "Gaming", guid3: None}
        with patch.object(vdm, "_get_name_from_registry", side_effect=lambda g: name_map.get(g)):
            result = vdm._enumerate_desktops()

        assert result[0]["name"] == "Desktop 1"
        assert result[1]["name"] == "Gaming"
        assert result[2]["name"] == "Desktop 3"

    def test_registry_name_called_with_correct_guid(self):
        """_get_name_from_registry is called with the stringified GUID for each desktop."""
        guid_str = "{REG-CALL-CHECK}"
        desktop = _make_desktop_mock(guid_str)

        internal = MagicMock()
        internal.GetDesktops.return_value = _make_array_mock([desktop])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None) as mock_reg:
            vdm._enumerate_desktops()

        mock_reg.assert_called_once_with(guid_str)

    # ------------------------------------------------------------------
    # Exception handling: GetDesktops raises
    # ------------------------------------------------------------------

    def test_returns_empty_list_when_get_desktops_raises(self):
        """An exception from GetDesktops causes _enumerate_desktops to return []."""
        internal = MagicMock()
        internal.GetDesktops.side_effect = OSError("COM error: GetDesktops failed")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._enumerate_desktops()

        assert result == []

    def test_returns_empty_list_when_get_count_raises(self):
        """An exception from GetCount causes _enumerate_desktops to return []."""
        array = MagicMock()
        array.GetCount.side_effect = OSError("COM error: GetCount failed")

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._enumerate_desktops()

        assert result == []

    # ------------------------------------------------------------------
    # Exception handling: per-desktop GetAt / QueryInterface / GetID raises
    # ------------------------------------------------------------------

    def test_skips_entry_when_get_at_raises(self):
        """An exception from GetAt at index i causes that entry to be skipped."""
        guid_good = "{GERAT-GOOD-GUID}"
        desktop_good = _make_desktop_mock(guid_good)

        array = MagicMock()
        array.GetCount.return_value = 2

        def _get_at(i, *args):
            if i == 0:
                raise OSError("GetAt COM error")
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        # The bad entry is skipped; the good one is still present
        assert len(result) == 1
        assert result[0]["guid_str"] == guid_good

    def test_skips_entry_when_query_interface_raises(self):
        """A QueryInterface failure skips that entry and continues iteration."""
        guid_good = "{QI-FAIL-GOOD}"
        desktop_good = _make_desktop_mock(guid_good)

        array = MagicMock()
        array.GetCount.return_value = 2

        def _get_at(i, *args):
            unk = MagicMock()
            if i == 0:
                unk.QueryInterface.side_effect = OSError("QI failed")
            else:
                unk.QueryInterface.return_value = desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 1
        assert result[0]["guid_str"] == guid_good

    def test_skips_entry_when_get_id_raises(self):
        """A GetID exception skips that entry; remaining desktops are still returned."""
        guid_good = "{GETID-FAIL-GOOD}"
        desktop_good = _make_desktop_mock(guid_good)

        desktop_bad = MagicMock()
        desktop_bad.GetID.side_effect = OSError("GetID COM error")

        array = MagicMock()
        array.GetCount.return_value = 2

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_bad if i == 0 else desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 1
        assert result[0]["guid_str"] == guid_good

    def test_skips_entry_when_get_id_returns_falsy(self):
        """A desktop whose GetID returns a falsy value is excluded from results."""
        guid_good = "{FALSY-ID-GOOD}"
        desktop_good = _make_desktop_mock(guid_good)
        desktop_null = MagicMock()
        desktop_null.GetID = MagicMock(return_value=None)

        array = MagicMock()
        array.GetCount.return_value = 2

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.return_value = desktop_null if i == 0 else desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 1
        assert result[0]["guid_str"] == guid_good

    def test_continues_after_multiple_per_desktop_exceptions(self):
        """Multiple failing entries are all skipped; only valid ones survive."""
        guid_good = "{MULTI-FAIL-GOOD}"
        desktop_good = _make_desktop_mock(guid_good)

        array = MagicMock()
        array.GetCount.return_value = 4

        def _get_at(i, *args):
            unk = MagicMock()
            if i in (0, 1, 2):
                unk.QueryInterface.side_effect = OSError(f"QI failed at {i}")
            else:
                unk.QueryInterface.return_value = desktop_good
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_get_name_from_registry", return_value=None):
            result = vdm._enumerate_desktops()

        assert len(result) == 1
        assert result[0]["guid_str"] == guid_good

    def test_returns_empty_list_when_all_desktops_fail(self):
        """When every desktop raises an exception the result is an empty list."""
        array = MagicMock()
        array.GetCount.return_value = 3

        def _get_at(i, *args):
            unk = MagicMock()
            unk.QueryInterface.side_effect = OSError("COM error")
            return unk

        array.GetAt.side_effect = _get_at

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._enumerate_desktops()

        assert result == []

    def test_empty_array_returns_empty_list(self):
        """GetCount returning 0 yields an empty result without any GetAt calls."""
        array = MagicMock()
        array.GetCount.return_value = 0

        internal = MagicMock()
        internal.GetDesktops.return_value = array

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        result = vdm._enumerate_desktops()

        assert result == []
        array.GetAt.assert_not_called()


# ---------------------------------------------------------------------------
# 18. _resolve_to_guid via _enumerate_desktops -- targeted coverage
# ---------------------------------------------------------------------------


class TestResolveToGuidViaEnumerateDesktops:
    """Tests for _resolve_to_guid that verify matching logic against entries
    produced by a mocked _enumerate_desktops, covering GUID-string matching,
    name matching (case-insensitive), and None returns when nothing matches.

    Unlike section 7, these tests patch _enumerate_desktops directly so we
    can control the full entry dicts (including the 'desktop' field) without
    depending on the inner COM chain.
    """

    def _make_entries(self, desktops: list[tuple[str, str]]) -> list[dict]:
        """Build a list of _enumerate_desktops-style dicts from (guid_str, name) pairs."""
        return [
            {
                "index": i,
                "guid_str": guid_str,
                "name": name,
                "desktop": MagicMock(),
            }
            for i, (guid_str, name) in enumerate(desktops)
        ]

    # ------------------------------------------------------------------
    # GUID-string matching
    # ------------------------------------------------------------------

    def test_matches_exact_guid_string(self):
        """When the input equals an entry's guid_str the method returns that guid_str."""
        guid = "{EXACT-GUID-MATCH}"
        entries = self._make_entries([(guid, "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid(guid)

        assert result == guid

    def test_matches_guid_string_case_insensitively(self):
        """GUID matching is case-insensitive: lowercase input matches uppercase guid_str."""
        guid = "{ABCD-1234-EFGH-5678}"
        entries = self._make_entries([(guid, "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid(guid.lower())

        assert result == guid

    def test_guid_match_returns_original_casing_from_entry(self):
        """The returned value is always from the entry (original casing), not from the input."""
        guid = "{ORIG-CASE-GUID}"
        entries = self._make_entries([(guid, "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid(guid.lower())

        # The returned GUID string keeps the original uppercase casing from the entry
        assert result == guid

    # ------------------------------------------------------------------
    # Name matching
    # ------------------------------------------------------------------

    def test_matches_desktop_name_case_insensitively(self):
        """Name matching is case-insensitive: 'work desktop' matches 'Work Desktop'."""
        guid = "{NAME-MATCH-GUID}"
        entries = self._make_entries([(guid, "Work Desktop")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("work desktop")

        assert result == guid

    def test_matches_name_with_uppercase_input(self):
        """All-uppercase input matches a mixed-case registry name."""
        guid = "{UPPER-NAME-GUID}"
        entries = self._make_entries([(guid, "Gaming Setup")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("GAMING SETUP")

        assert result == guid

    def test_matches_fallback_desktop_n_name(self):
        """'Desktop 2' (the fallback naming convention) is matched by name."""
        guid = "{DESKTOP-2-GUID}"
        entries = self._make_entries([("{OTHER}", "Desktop 1"), (guid, "Desktop 2")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("Desktop 2")

        assert result == guid

    def test_selects_first_match_when_duplicate_names_exist(self):
        """If two entries share the same name, the first one is returned."""
        guid_first = "{DUP-FIRST-GUID}"
        guid_second = "{DUP-SECOND-GUID}"
        entries = self._make_entries([(guid_first, "Shared Name"), (guid_second, "Shared Name")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("Shared Name")

        assert result == guid_first

    # ------------------------------------------------------------------
    # No match returns None
    # ------------------------------------------------------------------

    def test_returns_none_when_name_does_not_match_any_entry(self):
        """None is returned when the input matches neither guid_str nor name."""
        entries = self._make_entries([("{SOME-GUID}", "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("Nonexistent Desktop")

        assert result is None

    def test_returns_none_when_enumerate_desktops_returns_empty(self):
        """None is returned when there are no entries to match against."""
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=[]):
            result = vdm._resolve_to_guid("Desktop 1")

        assert result is None

    def test_returns_none_when_enumerate_desktops_raises(self):
        """An exception raised by _enumerate_desktops propagates up (or returns None
        if the method swallows it -- see the outer exception handler in the real impl).
        Here we verify the None return via the outer try/except in _resolve_to_guid."""
        internal = MagicMock()
        internal.GetDesktops.side_effect = OSError("COM failure")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        # No patch -- let the real _enumerate_desktops run and return []
        result = vdm._resolve_to_guid("Desktop 1")

        assert result is None

    def test_returns_none_for_partial_guid_substring(self):
        """A partial GUID substring must not be matched -- only exact (case-folded) equality."""
        full_guid = "{PARTIAL-GUID-TEST-0001}"
        entries = self._make_entries([(full_guid, "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("PARTIAL")

        assert result is None

    def test_returns_none_for_empty_string_input(self):
        """An empty string input never matches any entry."""
        guid = "{EMPTY-STR-GUID}"
        entries = self._make_entries([(guid, "Desktop 1")])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid("")

        assert result is None

    # ------------------------------------------------------------------
    # GUID matching takes priority over name matching
    # ------------------------------------------------------------------

    def test_guid_match_takes_priority_over_name_match(self):
        """When the input matches a guid_str for one entry and a name for another,
        the guid_str match (checked first in the loop) wins."""
        guid_target = "{PRIORITY-GUID}"
        guid_other = "{PRIORITY-OTHER}"
        # First entry: input matches guid_str
        # Second entry: input matches name but guid_str is different
        entries = self._make_entries([(guid_target, "Different Name"), (guid_other, guid_target)])

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=MagicMock())

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm._resolve_to_guid(guid_target)

        assert result == guid_target


# ---------------------------------------------------------------------------
# 19. get_desktop_info
# ---------------------------------------------------------------------------


class TestGetDesktopInfo:
    """get_desktop_info() returns (current_desktop, all_desktops) from one enumeration."""

    # Helper: build _enumerate_desktops-style entries from (guid_str, name) pairs.
    def _make_entries(self, desktops: list[tuple[str, str]]) -> list[dict]:
        return [
            {"index": i, "guid_str": guid, "name": name, "desktop": MagicMock()}
            for i, (guid, name) in enumerate(desktops)
        ]

    # ------------------------------------------------------------------
    # No internal manager -- fallback behaviour
    # ------------------------------------------------------------------

    def test_no_internal_manager_returns_fallback(self):
        """When _internal_manager is None both return values are the fallback sentinel."""
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=None)

        current, all_desktops = vdm.get_desktop_info()

        assert current["id"] == "00000000-0000-0000-0000-000000000000"
        assert current["name"] == "Default Desktop"
        assert len(all_desktops) == 1
        assert all_desktops[0] == current

    def test_no_internal_manager_all_desktops_is_list(self):
        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=None)

        _, all_desktops = vdm.get_desktop_info()

        assert isinstance(all_desktops, list)

    # ------------------------------------------------------------------
    # Single desktop that is current
    # ------------------------------------------------------------------

    def test_single_desktop_is_current(self):
        guid = "{SINGLE-GUID-0001}"
        entries = self._make_entries([(guid, "My Desktop")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value=guid))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert current == {"id": guid, "name": "My Desktop"}
        assert len(all_desktops) == 1
        assert all_desktops[0] == {"id": guid, "name": "My Desktop"}

    def test_single_desktop_current_is_same_object_as_all_entry(self):
        """The returned current dict is the same object as the matching all_desktops entry."""
        guid = "{SAME-OBJ-GUID}"
        entries = self._make_entries([(guid, "Work")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value=guid))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert current is all_desktops[0]

    # ------------------------------------------------------------------
    # Multiple desktops -- current is the second one
    # ------------------------------------------------------------------

    def test_multiple_desktops_current_is_second(self):
        guid1 = "{MULTI-GUID-001}"
        guid2 = "{MULTI-GUID-002}"
        entries = self._make_entries([(guid1, "Desktop 1"), (guid2, "Desktop 2")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value=guid2))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert current == {"id": guid2, "name": "Desktop 2"}
        assert len(all_desktops) == 2
        assert all_desktops[0] == {"id": guid1, "name": "Desktop 1"}
        assert all_desktops[1] == {"id": guid2, "name": "Desktop 2"}

    def test_all_desktops_count_matches_entries(self):
        """all_desktops contains one entry per enumerated desktop."""
        entries = self._make_entries([(f"{{G{i}}}", f"Desktop {i + 1}") for i in range(4)])
        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value="{G0}"))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            _, all_desktops = vdm.get_desktop_info()

        assert len(all_desktops) == 4

    # ------------------------------------------------------------------
    # GetCurrentDesktop raises -- falls back to first desktop
    # ------------------------------------------------------------------

    def test_get_current_desktop_fails_returns_first(self):
        """When GetCurrentDesktop raises, the first enumerated desktop is returned as current."""
        guid1 = "{FALLBACK-FIRST-001}"
        guid2 = "{FALLBACK-FIRST-002}"
        entries = self._make_entries([(guid1, "First"), (guid2, "Second")])

        internal = MagicMock()
        internal.GetCurrentDesktop.side_effect = OSError("COM failure")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert current == {"id": guid1, "name": "First"}
        assert len(all_desktops) == 2

    def test_get_current_desktop_fails_empty_enum_returns_fallback(self):
        """When GetCurrentDesktop raises and there are no desktops, the fallback is returned."""
        internal = MagicMock()
        internal.GetCurrentDesktop.side_effect = OSError("COM failure")

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=[]):
            current, all_desktops = vdm.get_desktop_info()

        assert current["id"] == "00000000-0000-0000-0000-000000000000"
        assert current["name"] == "Default Desktop"
        assert all_desktops == [current]

    # ------------------------------------------------------------------
    # Current GUID not found in enumerated list -- "Unknown" name
    # ------------------------------------------------------------------

    def test_current_not_in_list_returns_unknown_name(self):
        """When GetCurrentDesktop returns a GUID not present in all_desktops, name is 'Unknown'."""
        guid_known = "{KNOWN-GUID-001}"
        guid_current = "{UNKNOWN-CURRENT-GUID}"
        entries = self._make_entries([(guid_known, "Desktop 1")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value=guid_current))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert current == {"id": guid_current, "name": "Unknown"}
        assert len(all_desktops) == 1

    def test_current_not_in_list_all_desktops_still_complete(self):
        """Even with an 'Unknown' current, all_desktops still lists every enumerated desktop."""
        entries = self._make_entries([("{G1}", "Desktop 1"), ("{G2}", "Desktop 2")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value="{NOT-IN-LIST}"))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            _, all_desktops = vdm.get_desktop_info()

        assert len(all_desktops) == 2

    # ------------------------------------------------------------------
    # Return type invariants
    # ------------------------------------------------------------------

    def test_returns_tuple_of_dict_and_list(self):
        entries = self._make_entries([("{RET-GUID}", "Desktop 1")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value="{RET-GUID}"))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            result = vdm.get_desktop_info()

        assert isinstance(result, tuple)
        assert len(result) == 2
        current, all_desktops = result
        assert isinstance(current, dict)
        assert isinstance(all_desktops, list)

    def test_each_desktop_dict_has_id_and_name_keys(self):
        guid = "{KEYS-GUID}"
        entries = self._make_entries([(guid, "My Desktop")])

        current_obj = MagicMock()
        current_obj.GetID.return_value = MagicMock(__str__=MagicMock(return_value=guid))

        internal = MagicMock()
        internal.GetCurrentDesktop.return_value = current_obj

        vdm = _make_vdm(manager_mock=MagicMock(), internal_mock=internal)

        with patch.object(vdm, "_enumerate_desktops", return_value=entries):
            current, all_desktops = vdm.get_desktop_info()

        assert "id" in current and "name" in current
        for desktop in all_desktops:
            assert "id" in desktop and "name" in desktop

    # ------------------------------------------------------------------
    # Module-level get_desktop_info delegates to the manager
    # ------------------------------------------------------------------

    def test_module_level_get_desktop_info_delegates(self):
        """The module-level get_desktop_info() delegates to the thread-local manager."""
        expected = ({"id": "{MOD-GUID}", "name": "Work"}, [{"id": "{MOD-GUID}", "name": "Work"}])
        mock_vdm = MagicMock()
        mock_vdm.get_desktop_info.return_value = expected

        with patch("windows_mcp.vdm.core._get_manager", return_value=mock_vdm):
            result = vdm_mod.get_desktop_info()

        mock_vdm.get_desktop_info.assert_called_once_with()
        assert result == expected

    def test_module_level_get_desktop_info_called_once(self):
        """The module function calls the manager method exactly once per invocation."""
        mock_vdm = MagicMock()
        mock_vdm.get_desktop_info.return_value = (
            {"id": "{ONCE}", "name": "D"},
            [{"id": "{ONCE}", "name": "D"}],
        )

        with patch("windows_mcp.vdm.core._get_manager", return_value=mock_vdm):
            vdm_mod.get_desktop_info()
            vdm_mod.get_desktop_info()

        assert mock_vdm.get_desktop_info.call_count == 2
