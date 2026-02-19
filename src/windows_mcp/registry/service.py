"""Windows Registry operations using winreg stdlib.

Accepts PowerShell-style paths (HKCU:\\, HKLM:\\) and long-form paths
(HKEY_CURRENT_USER\\...). Stateless service -- no constructor dependencies.
"""

import winreg


class RegistryService:
    """CRUD operations on the Windows registry."""

    _REG_HIVE_MAP = {
        "HKCU": "HKEY_CURRENT_USER",
        "HKEY_CURRENT_USER": "HKEY_CURRENT_USER",
        "HKLM": "HKEY_LOCAL_MACHINE",
        "HKEY_LOCAL_MACHINE": "HKEY_LOCAL_MACHINE",
        "HKCR": "HKEY_CLASSES_ROOT",
        "HKEY_CLASSES_ROOT": "HKEY_CLASSES_ROOT",
        "HKU": "HKEY_USERS",
        "HKEY_USERS": "HKEY_USERS",
        "HKCC": "HKEY_CURRENT_CONFIG",
        "HKEY_CURRENT_CONFIG": "HKEY_CURRENT_CONFIG",
    }

    _REG_TYPE_MAP = {
        "String": "REG_SZ",
        "ExpandString": "REG_EXPAND_SZ",
        "Binary": "REG_BINARY",
        "DWord": "REG_DWORD",
        "QWord": "REG_QWORD",
        "MultiString": "REG_MULTI_SZ",
    }

    def _parse_reg_path(self, path: str) -> tuple:
        """Parse a PowerShell-style registry path into (hive_key, subkey).

        Accepts paths like ``HKCU:\\Software\\MyApp`` or ``HKEY_LOCAL_MACHINE\\SOFTWARE\\Key``.
        """
        normalized = path.replace(":/", "\\").replace(":\\", "\\").replace("/", "\\")
        normalized = normalized.rstrip(":")
        parts = normalized.split("\\", 1)
        hive_name = parts[0].upper()
        subkey = parts[1] if len(parts) > 1 else ""

        hive_full = self._REG_HIVE_MAP.get(hive_name)
        if not hive_full:
            raise ValueError(f"Unknown registry hive: {parts[0]}")

        hive_key = getattr(winreg, hive_full)
        return hive_key, subkey

    def registry_get(self, path: str, name: str) -> str:
        try:
            hive, subkey = self._parse_reg_path(path)
            with winreg.OpenKey(hive, subkey) as key:
                value, _ = winreg.QueryValueEx(key, name)
                return f'Registry value [{path}] "{name}" = {value}'
        except OSError as e:
            return f"Error reading registry: {e}"
        except ValueError as e:
            return f"Error reading registry: {e}"

    def registry_set(self, path: str, name: str, value: str, reg_type: str = "String") -> str:
        allowed_types = {"String", "ExpandString", "Binary", "DWord", "MultiString", "QWord"}
        if reg_type not in allowed_types:
            return (
                f"Error: invalid registry type '{reg_type}'. "
                f"Allowed: {', '.join(sorted(allowed_types))}"
            )

        type_const_name = self._REG_TYPE_MAP[reg_type]
        reg_type_const = getattr(winreg, type_const_name)

        if reg_type in ("DWord", "QWord"):
            typed_value = int(value)
        elif reg_type == "Binary":
            typed_value = bytes.fromhex(value)
        elif reg_type == "MultiString":
            typed_value = value.split("\\0")
        else:
            typed_value = value

        try:
            hive, subkey = self._parse_reg_path(path)
            winreg.CreateKey(hive, subkey)
            with winreg.OpenKey(hive, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, name, 0, reg_type_const, typed_value)
            return f'Registry value [{path}] "{name}" set to "{value}" (type: {reg_type}).'
        except OSError as e:
            return f"Error writing registry: {e}"
        except ValueError as e:
            return f"Error writing registry: {e}"

    def registry_delete(self, path: str, name: str | None = None) -> str:
        try:
            hive, subkey = self._parse_reg_path(path)
            if name:
                with winreg.OpenKey(hive, subkey, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, name)
                return f'Registry value [{path}] "{name}" deleted.'
            else:
                winreg.DeleteKey(hive, subkey)
                return f"Registry key [{path}] deleted."
        except OSError as e:
            return f"Error deleting registry {'value' if name else 'key'}: {e}"
        except ValueError as e:
            return f"Error deleting registry: {e}"

    def registry_list(self, path: str) -> str:
        try:
            hive, subkey = self._parse_reg_path(path)
            with winreg.OpenKey(hive, subkey) as key:
                values = []
                i = 0
                while True:
                    try:
                        vname, vdata, vtype = winreg.EnumValue(key, i)
                        values.append(f"  {vname} = {vdata}")
                        i += 1
                    except OSError:
                        break

                subkeys = []
                i = 0
                while True:
                    try:
                        sk_name = winreg.EnumKey(key, i)
                        subkeys.append(sk_name)
                        i += 1
                    except OSError:
                        break

            parts = []
            if values:
                parts.append("Values:\n" + "\n".join(values))
            if subkeys:
                parts.append("Sub-Keys:\n" + "\n".join(f"  {sk}" for sk in subkeys))
            if not parts:
                parts.append("No values or sub-keys found.")
            return f"Registry key [{path}]:\n" + "\n\n".join(parts)
        except OSError as e:
            return f"Error listing registry: {e}"
        except ValueError as e:
            return f"Error listing registry: {e}"
