"""Windows Registry operations using winreg stdlib.

Accepts PowerShell-style paths (HKCU:\\, HKLM:\\) and long-form paths
(HKEY_CURRENT_USER\\...). Stateless service -- no constructor dependencies.

Security: Write/delete operations to sensitive registry paths (Run, RunOnce,
Services, Policies, SAM, Security) are blocked by default.  Set
WINDOWS_MCP_REGISTRY_UNRESTRICTED=true to bypass.
"""

import os
import re
import winreg

# Regex patterns for security-sensitive registry subkeys
# Matched case-insensitively against the normalized subkey portion
_SENSITIVE_KEY_PATTERNS = [
    re.compile(p, re.IGNORECASE)
    for p in [
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\Run\b",
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce\b",
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\RunServices\b",
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\Policies\b",
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders\b",
        r"^Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\b",
        r"^SYSTEM\\CurrentControlSet\\Services\b",
        r"^SYSTEM\\CurrentControlSet\\Control\\Session Manager\b",
        r"^SAM\b",
        r"^SECURITY\b",
        r"^SOFTWARE\\Policies\b",
    ]
]


def _is_sensitive_key(subkey: str) -> bool:
    """Return True if the subkey matches a sensitive registry path."""
    normalized = subkey.replace("/", "\\").strip("\\")
    return any(pat.search(normalized) for pat in _SENSITIVE_KEY_PATTERNS)


def _check_registry_write(subkey: str) -> None:
    """Raise PermissionError if writing to a sensitive key without override."""
    unrestricted = os.environ.get("WINDOWS_MCP_REGISTRY_UNRESTRICTED", "").lower() == "true"
    if unrestricted:
        return
    if _is_sensitive_key(subkey):
        raise PermissionError(
            f"Write/delete to sensitive registry path '{subkey}' is blocked. "
            "Set WINDOWS_MCP_REGISTRY_UNRESTRICTED=true to bypass."
        )


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

        try:
            if reg_type in ("DWord", "QWord"):
                typed_value = int(value)
            elif reg_type == "Binary":
                typed_value = bytes.fromhex(value)
            elif reg_type == "MultiString":
                typed_value = value.split("\\0")
            else:
                typed_value = value

            hive, subkey = self._parse_reg_path(path)
            _check_registry_write(subkey)
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
            _check_registry_write(subkey)
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
