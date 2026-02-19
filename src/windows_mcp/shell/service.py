"""Shell command execution with safety blocklist.

Executes PowerShell commands via subprocess with configurable blocklist
filtering. The blocklist prevents accidental execution of destructive commands
(format, diskpart, etc.) and can be customized via the
WINDOWS_MCP_SHELL_BLOCKLIST environment variable.
"""

import base64
import logging
import os
import re
import subprocess
from locale import getpreferredencoding

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
#  Shell sandboxing: blocked command patterns (configurable via env var)
# ---------------------------------------------------------------------------
# Default blocklist: commands that can cause irreversible damage.
# Set WINDOWS_MCP_SHELL_BLOCKLIST to a comma-separated list to override.
# Set WINDOWS_MCP_SHELL_BLOCKLIST="" (empty) to disable the blocklist.
_DEFAULT_SHELL_BLOCKLIST = [
    r"\bformat\b.*[a-zA-Z]:",  # format C:, format D:, etc.
    r"\brm\s+-rf\s+[/\\]",  # rm -rf /
    r"\bRemove-Item\b.*-Recurse.*[/\\]\s*$",  # Remove-Item -Recurse C:\
    r"\bdel\s+/[sS]\s+[/\\]",  # del /s \
    r"\brd\s+/[sS]\s+[/\\]",  # rd /s \
    r"\bclear-disk\b",  # Clear-Disk
    r"\bstop-computer\b",  # Stop-Computer (shutdown)
    r"\brestart-computer\b",  # Restart-Computer
    r"\bdiskpart\b",  # diskpart
    r"\bbcdedit\b",  # bcdedit (boot config)
    r"\bsfc\s+/scannow\b",  # sfc /scannow (system file checker)
    r"\breg\s+delete\s+HK",  # reg delete HKLM\...
    r"\bnet\s+user\b.*/add\b",  # net user <x> /add
    r"\bnet\s+localgroup\s+administrators\b.*/add\b",  # privilege escalation
    r"\bInvoke-Expression\b.*\bDownloadString\b",  # IEX(downloadstring) cradle
    r"\biex\b.*\bNet\.WebClient\b",  # IEX cradle variant
]

_shell_blocklist_patterns: list[re.Pattern] | None = None


def _get_shell_blocklist() -> list[re.Pattern]:
    """Lazily compile and cache the shell blocklist patterns."""
    global _shell_blocklist_patterns
    if _shell_blocklist_patterns is not None:
        return _shell_blocklist_patterns

    env_val = os.environ.get("WINDOWS_MCP_SHELL_BLOCKLIST")
    if env_val is not None:
        if env_val.strip() == "":
            _shell_blocklist_patterns = []
        else:
            raw = [p.strip() for p in env_val.split(",") if p.strip()]
            _shell_blocklist_patterns = [re.compile(p, re.IGNORECASE) for p in raw]
    else:
        _shell_blocklist_patterns = [re.compile(p, re.IGNORECASE) for p in _DEFAULT_SHELL_BLOCKLIST]
    return _shell_blocklist_patterns


class ShellService:
    """PowerShell command execution with safety blocklist."""

    def __init__(self) -> None:
        self.encoding = getpreferredencoding()

    @staticmethod
    def ps_quote(value: str) -> str:
        """Wrap a value in a PowerShell single-quoted string literal.

        Single-quoted strings in PowerShell are truly literal -- they do NOT
        expand variables ($env:X), subexpressions ($(...)), or escape sequences.
        The only character that needs escaping is the single quote itself,
        which is doubled ('').
        """
        return "'" + value.replace("'", "''") + "'"

    @staticmethod
    def check_blocklist(command: str) -> str | None:
        """Check a command against the shell blocklist.

        Returns the matched pattern description if blocked, or None if allowed.
        """
        for pattern in _get_shell_blocklist():
            if pattern.search(command):
                return pattern.pattern
        return None

    def execute(self, command: str, timeout: int = 10) -> tuple[str, int]:
        """Execute a PowerShell command and return (output, return_code)."""
        blocked = self.check_blocklist(command)
        if blocked:
            logger.warning("Shell command blocked by safety filter: %s", blocked)
            return (
                f"Command blocked by safety filter (matched pattern: {blocked}). "
                "Set WINDOWS_MCP_SHELL_BLOCKLIST env var to customize.",
                1,
            )
        try:
            encoded = base64.b64encode(command.encode("utf-16le")).decode("ascii")
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-OutputFormat",
                    "Text",
                    "-EncodedCommand",
                    encoded,
                ],
                capture_output=True,
                timeout=timeout,
                cwd=os.path.expanduser(path="~"),
                env=os.environ.copy(),
            )
            stdout = result.stdout
            stderr = result.stderr
            if isinstance(stdout, bytes):
                stdout = stdout.decode(self.encoding, errors="ignore")
            if isinstance(stderr, bytes):
                stderr = stderr.decode(self.encoding, errors="ignore")
            return (stdout or stderr, result.returncode)
        except subprocess.TimeoutExpired:
            return ("Command execution timed out", 1)
        except Exception as e:
            return (f"Command execution failed: {type(e).__name__}: {e}", 1)
