"""API key management with DPAPI encryption for local HTTP transport auth."""

import logging
import os
import secrets
from pathlib import Path

logger = logging.getLogger(__name__)

# DPAPI storage location
_APP_DATA_DIR = Path(os.environ.get("APPDATA", "")) / "windows-mcp"
_KEY_FILE = _APP_DATA_DIR / "auth.key"


def _dpapi_encrypt(data: bytes) -> bytes:
    """Encrypt data using Windows DPAPI (tied to current user account)."""
    import win32crypt

    return win32crypt.CryptProtectData(data, "windows-mcp-auth", None, None, None, 0)


def _dpapi_decrypt(encrypted: bytes) -> bytes:
    """Decrypt DPAPI-protected data."""
    import win32crypt

    _, decrypted = win32crypt.CryptUnprotectData(encrypted, None, None, None, 0)
    return decrypted


class AuthKeyManager:
    """Manages API keys for local HTTP transport authentication.

    Keys are encrypted with Windows DPAPI and stored in %APPDATA%/windows-mcp/auth.key.
    DPAPI ties encryption to the current Windows user account.
    """

    @staticmethod
    def generate_key() -> str:
        """Generate a new 32-byte random API key and store it encrypted.

        Returns the key as a hex string for the user to copy.
        """
        key = secrets.token_hex(32)
        _APP_DATA_DIR.mkdir(parents=True, exist_ok=True)

        encrypted = _dpapi_encrypt(key.encode("utf-8"))
        _KEY_FILE.write_bytes(encrypted)
        logger.info("API key generated and stored in %s", _KEY_FILE)
        return key

    @staticmethod
    def load_key() -> str | None:
        """Load and decrypt the stored API key, or return None if not found."""
        if not _KEY_FILE.exists():
            return None
        try:
            encrypted = _KEY_FILE.read_bytes()
            decrypted = _dpapi_decrypt(encrypted)
            return decrypted.decode("utf-8")
        except Exception as e:
            logger.warning("Failed to load stored API key: %s", e)
            return None

    @staticmethod
    def rotate_key() -> str:
        """Generate a new key, replacing the existing one.

        Returns the new key as a hex string.
        """
        return AuthKeyManager.generate_key()

    @staticmethod
    def validate_key(provided: str, stored: str) -> bool:
        """Constant-time comparison of provided key against stored key."""
        return secrets.compare_digest(provided, stored)

    @staticmethod
    def has_stored_key() -> bool:
        """Check if an encrypted key file exists."""
        return _KEY_FILE.exists()
