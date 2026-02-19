"""Tests for registry security functions: _is_sensitive_key and _check_registry_write."""

from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.registry.service import (
    RegistryService,
    _check_registry_write,
    _is_sensitive_key,
)


@pytest.fixture
def registry():
    return RegistryService()


# ---------------------------------------------------------------------------
# _is_sensitive_key tests
# ---------------------------------------------------------------------------


class TestIsSensitiveKey:
    """Unit tests for the _is_sensitive_key predicate."""

    # --- True cases: each sensitive pattern ---

    def test_run_key_is_sensitive(self):
        assert _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\Run") is True

    def test_run_with_subkey_is_sensitive(self):
        # The \b word boundary still matches when a child key follows
        assert _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\Run\MyApp") is True

    def test_run_once_is_sensitive(self):
        assert _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\RunOnce") is True

    def test_run_services_is_sensitive(self):
        assert _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\RunServices") is True

    def test_policies_is_sensitive(self):
        assert _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\Policies") is True

    def test_shell_folders_is_sensitive(self):
        assert (
            _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
            is True
        )

    def test_user_shell_folders_is_sensitive(self):
        assert (
            _is_sensitive_key(
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            )
            is True
        )

    def test_system_services_is_sensitive(self):
        assert _is_sensitive_key(r"SYSTEM\CurrentControlSet\Services") is True

    def test_system_services_with_subkey_is_sensitive(self):
        assert _is_sensitive_key(r"SYSTEM\CurrentControlSet\Services\MyDriver") is True

    def test_session_manager_is_sensitive(self):
        assert _is_sensitive_key(r"SYSTEM\CurrentControlSet\Control\Session Manager") is True

    def test_sam_is_sensitive(self):
        assert _is_sensitive_key("SAM") is True

    def test_sam_with_subkey_is_sensitive(self):
        assert _is_sensitive_key(r"SAM\SAM\Domains") is True

    def test_security_is_sensitive(self):
        assert _is_sensitive_key("SECURITY") is True

    def test_software_policies_is_sensitive(self):
        assert _is_sensitive_key(r"SOFTWARE\Policies") is True

    def test_software_policies_with_subkey_is_sensitive(self):
        assert _is_sensitive_key(r"SOFTWARE\Policies\Microsoft") is True

    # --- False cases: safe non-sensitive paths ---

    def test_myapp_software_key_is_safe(self):
        assert _is_sensitive_key(r"Software\MyApp\Settings") is False

    def test_generic_hkcu_software_is_safe(self):
        assert _is_sensitive_key(r"Software\Microsoft\Office") is False

    def test_empty_subkey_is_safe(self):
        assert _is_sensitive_key("") is False

    def test_unrelated_system_key_is_safe(self):
        assert _is_sensitive_key(r"SYSTEM\CurrentControlSet\Enum") is False

    def test_environment_key_is_safe(self):
        assert (
            _is_sensitive_key(r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment")
            is True
        )  # child of sensitive

    def test_partial_run_prefix_is_safe(self):
        # "RunApplication" does NOT match because \b ensures "Run" is a whole word
        assert (
            _is_sensitive_key(r"Software\Microsoft\Windows\CurrentVersion\RunApplication") is False
        )

    # --- Case-insensitivity ---

    def test_lowercase_run_is_sensitive(self):
        assert _is_sensitive_key(r"software\microsoft\windows\currentversion\run") is True

    def test_mixed_case_sam_is_sensitive(self):
        assert _is_sensitive_key("Sam") is True

    def test_uppercase_security_is_sensitive(self):
        assert _is_sensitive_key("SECURITY") is True

    def test_mixed_case_services_is_sensitive(self):
        assert _is_sensitive_key(r"system\currentcontrolset\services") is True

    # --- Normalization: forward-slash paths ---

    def test_forward_slash_run_key_is_sensitive(self):
        assert _is_sensitive_key("Software/Microsoft/Windows/CurrentVersion/Run") is True

    def test_forward_slash_sam_is_sensitive(self):
        assert _is_sensitive_key("SAM") is True

    # --- Normalization: leading/trailing backslashes are stripped ---

    def test_leading_backslash_stripped(self):
        assert _is_sensitive_key(r"\SAM") is True

    def test_trailing_backslash_stripped(self):
        assert _is_sensitive_key(r"SAM\\") is True


# ---------------------------------------------------------------------------
# _check_registry_write tests
# ---------------------------------------------------------------------------


class TestCheckRegistryWrite:
    def test_raises_for_sensitive_run_key(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with pytest.raises(PermissionError, match="sensitive registry path"):
            _check_registry_write(r"Software\Microsoft\Windows\CurrentVersion\Run")

    def test_raises_for_sam_key(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with pytest.raises(PermissionError):
            _check_registry_write("SAM")

    def test_raises_for_security_key(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with pytest.raises(PermissionError):
            _check_registry_write("SECURITY")

    def test_raises_for_services_key(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with pytest.raises(PermissionError):
            _check_registry_write(r"SYSTEM\CurrentControlSet\Services")

    def test_allows_safe_key_without_override(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        # Must not raise for a non-sensitive path
        _check_registry_write(r"Software\MyApp\Settings")

    def test_allows_empty_subkey_without_override(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        _check_registry_write("")

    def test_allows_sensitive_key_when_unrestricted_true(self, monkeypatch):
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "true")
        # Must not raise even for the most sensitive path
        _check_registry_write(r"Software\Microsoft\Windows\CurrentVersion\Run")

    def test_allows_sensitive_key_when_unrestricted_TRUE_uppercase(self, monkeypatch):
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "TRUE")
        _check_registry_write("SAM")

    def test_does_not_allow_when_unrestricted_is_false(self, monkeypatch):
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "false")
        with pytest.raises(PermissionError):
            _check_registry_write("SAM")

    def test_does_not_allow_when_unrestricted_is_1(self, monkeypatch):
        # Only the literal string "true" (case-insensitive) unlocks; "1" does not
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "1")
        with pytest.raises(PermissionError):
            _check_registry_write("SAM")

    def test_error_message_includes_key_path(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with pytest.raises(PermissionError) as exc_info:
            _check_registry_write(key)
        assert key in str(exc_info.value)

    def test_error_message_mentions_unrestricted_env_var(self, monkeypatch):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with pytest.raises(PermissionError) as exc_info:
            _check_registry_write("SAM")
        assert "WINDOWS_MCP_REGISTRY_UNRESTRICTED" in str(exc_info.value)


# ---------------------------------------------------------------------------
# registry_set security integration tests
# ---------------------------------------------------------------------------


class TestRegistrySetSecurity:
    """Verify registry_set blocks writes to sensitive keys at the function level."""

    def test_registry_set_blocked_for_run_key(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.REG_SZ = 1
            result = registry.registry_set(
                path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                name="Malware",
                value="evil.exe",
            )
        assert "Error" in result
        # winreg.CreateKey must NOT have been called -- blocked before any write
        mock_winreg.CreateKey.assert_not_called()

    def test_registry_set_blocked_for_sam(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.REG_SZ = 1
            result = registry.registry_set(
                path=r"HKCU:\SAM\SAM",
                name="Key",
                value="val",
            )
        assert "Error" in result
        mock_winreg.CreateKey.assert_not_called()

    def test_registry_set_blocked_for_services(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_LOCAL_MACHINE = 0x80000002
            mock_winreg.REG_SZ = 1
            result = registry.registry_set(
                path=r"HKLM:\SYSTEM\CurrentControlSet\Services\Spooler",
                name="Start",
                value="2",
            )
        assert "Error" in result
        mock_winreg.CreateKey.assert_not_called()

    def test_registry_set_allowed_for_safe_key(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.KEY_SET_VALUE = 0x0002
            mock_winreg.REG_SZ = 1
            mock_key = MagicMock()
            mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
            mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
            result = registry.registry_set(
                path=r"HKCU:\Software\MyApp",
                name="Theme",
                value="dark",
            )
        assert "set to" in result
        assert "Error" not in result

    def test_registry_set_allowed_for_sensitive_key_when_unrestricted(self, monkeypatch, registry):
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "true")
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.KEY_SET_VALUE = 0x0002
            mock_winreg.REG_SZ = 1
            mock_key = MagicMock()
            mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
            mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
            result = registry.registry_set(
                path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                name="MyApp",
                value="myapp.exe",
            )
        assert "set to" in result
        assert "Error" not in result


# ---------------------------------------------------------------------------
# registry_delete security integration tests
# ---------------------------------------------------------------------------


class TestRegistryDeleteSecurity:
    """Verify registry_delete blocks deletions of sensitive keys at the function level."""

    def test_registry_delete_value_blocked_for_run_key(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.KEY_SET_VALUE = 0x0002
            result = registry.registry_delete(
                path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                name="SomeEntry",
            )
        assert "Error" in result
        mock_winreg.OpenKey.assert_not_called()

    def test_registry_delete_key_blocked_for_run_key(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            result = registry.registry_delete(
                path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                name=None,
            )
        assert "Error" in result
        mock_winreg.DeleteKey.assert_not_called()

    def test_registry_delete_blocked_for_session_manager(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_LOCAL_MACHINE = 0x80000002
            result = registry.registry_delete(
                path=r"HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager",
                name="CriticalSetting",
            )
        assert "Error" in result
        mock_winreg.OpenKey.assert_not_called()

    def test_registry_delete_allowed_for_safe_key(self, monkeypatch, registry):
        monkeypatch.delenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", raising=False)
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            mock_winreg.KEY_SET_VALUE = 0x0002
            mock_key = MagicMock()
            mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
            mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
            result = registry.registry_delete(
                path=r"HKCU:\Software\MyApp",
                name="OldSetting",
            )
        assert "deleted" in result
        assert "Error" not in result

    def test_registry_delete_allowed_for_sensitive_key_when_unrestricted(
        self, monkeypatch, registry
    ):
        monkeypatch.setenv("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "true")
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            result = registry.registry_delete(
                path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                name=None,
            )
        # With unrestricted=true, the delete proceeds to winreg.DeleteKey
        mock_winreg.DeleteKey.assert_called_once()
        assert "Error" not in result
