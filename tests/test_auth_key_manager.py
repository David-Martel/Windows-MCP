"""Tests for AuthKeyManager -- DPAPI key generation, storage, and validation."""

import secrets
import sys
from types import ModuleType
from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.auth.key_manager import AuthKeyManager


class TestGenerateKey:
    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._APP_DATA_DIR")
    @patch("windows_mcp.auth.key_manager._dpapi_encrypt")
    def test_generates_64_char_hex_key(self, mock_encrypt, mock_dir, mock_file):
        mock_encrypt.return_value = b"encrypted"
        mock_dir.mkdir = MagicMock()
        mock_file.write_bytes = MagicMock()

        key = AuthKeyManager.generate_key()
        assert len(key) == 64
        # Verify it's valid hex
        int(key, 16)

    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._APP_DATA_DIR")
    @patch("windows_mcp.auth.key_manager._dpapi_encrypt")
    def test_creates_app_dir(self, mock_encrypt, mock_dir, mock_file):
        mock_encrypt.return_value = b"encrypted"
        mock_file.write_bytes = MagicMock()

        AuthKeyManager.generate_key()
        mock_dir.mkdir.assert_called_once_with(parents=True, exist_ok=True)

    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._APP_DATA_DIR")
    @patch("windows_mcp.auth.key_manager._dpapi_encrypt")
    def test_encrypts_and_writes_key(self, mock_encrypt, mock_dir, mock_file):
        mock_encrypt.return_value = b"encrypted_bytes"
        mock_dir.mkdir = MagicMock()
        mock_file.write_bytes = MagicMock()

        key = AuthKeyManager.generate_key()
        mock_encrypt.assert_called_once_with(key.encode("utf-8"))
        mock_file.write_bytes.assert_called_once_with(b"encrypted_bytes")


class TestLoadKey:
    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._dpapi_decrypt")
    def test_loads_existing_key(self, mock_decrypt, mock_file):
        mock_file.exists.return_value = True
        mock_file.read_bytes.return_value = b"encrypted"
        mock_decrypt.return_value = b"mykey123"

        result = AuthKeyManager.load_key()
        assert result == "mykey123"

    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    def test_returns_none_if_no_file(self, mock_file):
        mock_file.exists.return_value = False
        assert AuthKeyManager.load_key() is None

    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._dpapi_decrypt")
    def test_returns_none_on_decrypt_error(self, mock_decrypt, mock_file):
        mock_file.exists.return_value = True
        mock_file.read_bytes.return_value = b"bad"
        mock_decrypt.side_effect = Exception("DPAPI decrypt failed")

        assert AuthKeyManager.load_key() is None


class TestValidateKey:
    def test_valid_key_matches(self):
        key = secrets.token_hex(32)
        assert AuthKeyManager.validate_key(key, key) is True

    def test_invalid_key_rejected(self):
        assert AuthKeyManager.validate_key("wrong", "correct") is False

    def test_empty_keys_match(self):
        assert AuthKeyManager.validate_key("", "") is True

    def test_timing_safe(self):
        # Validate that secrets.compare_digest is being used (constant-time)
        with patch("windows_mcp.auth.key_manager.secrets.compare_digest") as mock_compare:
            mock_compare.return_value = True
            AuthKeyManager.validate_key("a", "b")
            mock_compare.assert_called_once_with("a", "b")


class TestHasStoredKey:
    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    def test_returns_true_when_file_exists(self, mock_file):
        mock_file.exists.return_value = True
        assert AuthKeyManager.has_stored_key() is True

    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    def test_returns_false_when_no_file(self, mock_file):
        mock_file.exists.return_value = False
        assert AuthKeyManager.has_stored_key() is False


class TestRotateKey:
    @patch("windows_mcp.auth.key_manager._KEY_FILE")
    @patch("windows_mcp.auth.key_manager._APP_DATA_DIR")
    @patch("windows_mcp.auth.key_manager._dpapi_encrypt")
    def test_returns_new_key(self, mock_encrypt, mock_dir, mock_file):
        mock_encrypt.return_value = b"encrypted"
        mock_dir.mkdir = MagicMock()
        mock_file.write_bytes = MagicMock()

        key = AuthKeyManager.rotate_key()
        assert len(key) == 64
        int(key, 16)  # Valid hex


class TestDpapiEncrypt:
    """Test the _dpapi_encrypt function directly via a mocked win32crypt module."""

    def test_encrypt_calls_crypt_protect_data(self):
        """_dpapi_encrypt should call CryptProtectData with expected arguments."""
        from windows_mcp.auth import key_manager

        mock_win32crypt = ModuleType("win32crypt")
        mock_win32crypt.CryptProtectData = MagicMock(return_value=b"encrypted_output")

        with patch.dict(sys.modules, {"win32crypt": mock_win32crypt}):
            result = key_manager._dpapi_encrypt(b"plaintext")

        mock_win32crypt.CryptProtectData.assert_called_once_with(
            b"plaintext", "windows-mcp-auth", None, None, None, 0
        )
        assert result == b"encrypted_output"

    def test_encrypt_propagates_win32crypt_error(self):
        """_dpapi_encrypt should let win32crypt exceptions bubble up."""
        from windows_mcp.auth import key_manager

        mock_win32crypt = ModuleType("win32crypt")
        mock_win32crypt.CryptProtectData = MagicMock(side_effect=OSError("DPAPI unavailable"))

        with patch.dict(sys.modules, {"win32crypt": mock_win32crypt}):
            with pytest.raises(OSError, match="DPAPI unavailable"):
                key_manager._dpapi_encrypt(b"data")


class TestDpapiDecrypt:
    """Test the _dpapi_decrypt function directly via a mocked win32crypt module."""

    def test_decrypt_calls_crypt_unprotect_data(self):
        """_dpapi_decrypt should call CryptUnprotectData and return the second tuple element."""
        from windows_mcp.auth import key_manager

        mock_win32crypt = ModuleType("win32crypt")
        mock_win32crypt.CryptUnprotectData = MagicMock(return_value=(None, b"decrypted_bytes"))

        with patch.dict(sys.modules, {"win32crypt": mock_win32crypt}):
            result = key_manager._dpapi_decrypt(b"ciphertext")

        mock_win32crypt.CryptUnprotectData.assert_called_once_with(
            b"ciphertext", None, None, None, 0
        )
        assert result == b"decrypted_bytes"

    def test_decrypt_propagates_win32crypt_error(self):
        """_dpapi_decrypt should let win32crypt exceptions bubble up to the caller."""
        from windows_mcp.auth import key_manager

        mock_win32crypt = ModuleType("win32crypt")
        mock_win32crypt.CryptUnprotectData = MagicMock(
            side_effect=OSError("Decryption failed: wrong user")
        )

        with patch.dict(sys.modules, {"win32crypt": mock_win32crypt}):
            with pytest.raises(OSError, match="Decryption failed"):
                key_manager._dpapi_decrypt(b"bad_cipher")

    def test_load_key_returns_none_when_decrypt_raises_os_error(self):
        """load_key wraps _dpapi_decrypt errors and returns None instead of raising."""
        mock_win32crypt = ModuleType("win32crypt")
        mock_win32crypt.CryptUnprotectData = MagicMock(side_effect=OSError("DPAPI error"))

        with patch("windows_mcp.auth.key_manager._KEY_FILE") as mock_file:
            mock_file.exists.return_value = True
            mock_file.read_bytes.return_value = b"garbage"

            with patch.dict(sys.modules, {"win32crypt": mock_win32crypt}):
                result = AuthKeyManager.load_key()

        assert result is None
