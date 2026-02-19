from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.desktop.service import Desktop
from windows_mcp.registry.service import RegistryService


@pytest.fixture
def registry():
    return RegistryService()


class TestPsQuote:
    def test_simple_string(self):
        assert Desktop._ps_quote("hello") == "'hello'"

    def test_single_quote_escaping(self):
        assert Desktop._ps_quote("it's") == "'it''s'"

    def test_double_quotes_not_escaped(self):
        assert Desktop._ps_quote('say "hi"') == """'say "hi"'"""

    def test_dollar_sign_not_expanded(self):
        assert Desktop._ps_quote("$env:PATH") == "'$env:PATH'"

    def test_empty_string(self):
        assert Desktop._ps_quote("") == "''"

    def test_registry_path(self):
        result = Desktop._ps_quote("HKCU:\\Software\\Test")
        assert result == "'HKCU:\\Software\\Test'"


class TestParseRegPath:
    def test_hkcu_abbreviation(self, registry):
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            hive, subkey = registry._parse_reg_path("HKCU:\\Software\\Test")
            assert hive == 0x80000001
            assert subkey == "Software\\Test"

    def test_hklm_full_name(self, registry):
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_LOCAL_MACHINE = 0x80000002
            hive, subkey = registry._parse_reg_path("HKEY_LOCAL_MACHINE\\SOFTWARE\\Test")
            assert hive == 0x80000002
            assert subkey == "SOFTWARE\\Test"

    def test_unknown_hive_raises(self, registry):
        with pytest.raises(ValueError, match="Unknown registry hive"):
            registry._parse_reg_path("HKBOGUS:\\Software\\Test")

    def test_no_subkey(self, registry):
        with patch("windows_mcp.registry.service.winreg") as mock_winreg:
            mock_winreg.HKEY_CURRENT_USER = 0x80000001
            hive, subkey = registry._parse_reg_path("HKCU:")
            assert subkey == ""


class TestRegistryGet:
    @patch("windows_mcp.registry.service.winreg")
    def test_success(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
        mock_winreg.QueryValueEx.return_value = (42, 1)

        result = registry.registry_get(path="HKCU:\\Software\\Test", name="MyValue")
        assert "MyValue" in result
        assert "42" in result
        assert "Error" not in result

    @patch("windows_mcp.registry.service.winreg")
    def test_failure(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.OpenKey.side_effect = OSError("Property not found")

        result = registry.registry_get(path="HKCU:\\Software\\Test", name="Missing")
        assert "Error reading registry" in result

    def test_invalid_hive(self, registry):
        result = registry.registry_get(path="HKBOGUS:\\Software\\Test", name="Key")
        assert "Error reading registry" in result


class TestRegistrySet:
    @patch("windows_mcp.registry.service.winreg")
    def test_success(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_winreg.REG_SZ = 1
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)

        result = registry.registry_set(path="HKCU:\\Software\\Test", name="MyKey", value="hello")
        assert "set to" in result
        assert '"hello"' in result

    @patch("windows_mcp.registry.service.winreg")
    def test_failure(self, mock_winreg, registry):
        mock_winreg.HKEY_LOCAL_MACHINE = 0x80000002
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_winreg.REG_SZ = 1
        mock_winreg.CreateKey.return_value = None
        mock_winreg.OpenKey.side_effect = OSError("Access denied")

        result = registry.registry_set(path="HKLM:\\Software\\Test", name="Key", value="val")
        assert "Error writing registry" in result

    def test_invalid_type(self, registry):
        result = registry.registry_set(
            path="HKCU:\\Test", name="Key", value="val", reg_type="Invalid"
        )
        assert "Error: invalid registry type" in result
        assert "Invalid" in result

    @patch("windows_mcp.registry.service.winreg")
    def test_all_valid_types(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_winreg.REG_SZ = 1
        mock_winreg.REG_EXPAND_SZ = 2
        mock_winreg.REG_BINARY = 3
        mock_winreg.REG_DWORD = 4
        mock_winreg.REG_MULTI_SZ = 7
        mock_winreg.REG_QWORD = 11
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)

        for reg_type in ("String", "ExpandString", "Binary", "DWord", "MultiString", "QWord"):
            # Use valid values for each type
            value = "0" if reg_type in ("DWord", "QWord") else "00" if reg_type == "Binary" else "V"
            result = registry.registry_set(
                path="HKCU:\\Test", name="K", value=value, reg_type=reg_type
            )
            assert "Error" not in result, f"Failed for type {reg_type}: {result}"

    @patch("windows_mcp.registry.service.winreg")
    def test_creates_key(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_winreg.REG_SZ = 1
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)

        registry.registry_set(path="HKCU:\\Software\\NewKey", name="Val", value="1")
        mock_winreg.CreateKey.assert_called_once()


class TestRegistryDelete:
    @patch("windows_mcp.registry.service.winreg")
    def test_delete_value(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)

        result = registry.registry_delete(path="HKCU:\\Software\\Test", name="MyValue")
        assert "deleted" in result
        assert '"MyValue"' in result

    @patch("windows_mcp.registry.service.winreg")
    def test_delete_key(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001

        result = registry.registry_delete(path="HKCU:\\Software\\Test", name=None)
        assert "key" in result.lower()
        assert "deleted" in result
        mock_winreg.DeleteKey.assert_called_once()

    @patch("windows_mcp.registry.service.winreg")
    def test_delete_value_failure(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.KEY_SET_VALUE = 0x0002
        mock_winreg.OpenKey.side_effect = OSError("Not found")

        result = registry.registry_delete(path="HKCU:\\Software\\Test", name="Missing")
        assert "Error deleting registry value" in result

    @patch("windows_mcp.registry.service.winreg")
    def test_delete_key_failure(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.DeleteKey.side_effect = OSError("Access denied")

        result = registry.registry_delete(path="HKCU:\\Software\\Protected")
        assert "Error deleting registry key" in result


class TestRegistryList:
    @patch("windows_mcp.registry.service.winreg")
    def test_success(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
        # EnumValue returns (name, data, type) then raises OSError
        mock_winreg.EnumValue.side_effect = [
            ("MyKey", "hello", 1),
            OSError("no more"),
        ]
        mock_winreg.EnumKey.side_effect = [
            "Child1",
            OSError("no more"),
        ]

        result = registry.registry_list(path="HKCU:\\Software\\Test")
        assert "MyKey" in result
        assert "hello" in result
        assert "Child1" in result

    @patch("windows_mcp.registry.service.winreg")
    def test_failure(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_winreg.OpenKey.side_effect = OSError("Path not found")

        result = registry.registry_list(path="HKCU:\\Software\\Missing")
        assert "Error listing registry" in result

    @patch("windows_mcp.registry.service.winreg")
    def test_empty(self, mock_winreg, registry):
        mock_winreg.HKEY_CURRENT_USER = 0x80000001
        mock_key = MagicMock()
        mock_winreg.OpenKey.return_value.__enter__ = MagicMock(return_value=mock_key)
        mock_winreg.OpenKey.return_value.__exit__ = MagicMock(return_value=False)
        mock_winreg.EnumValue.side_effect = OSError("no more")
        mock_winreg.EnumKey.side_effect = OSError("no more")

        result = registry.registry_list(path="HKCU:\\Software\\Empty")
        assert "No values or sub-keys found" in result
