"""Process listing, termination, and safety controls.

Stateless service providing process enumeration with fuzzy name matching
and kill/terminate with a protected-process blocklist.
"""

import logging
import re
from typing import Literal

import psutil

logger = logging.getLogger(__name__)

# Processes that must never be terminated -- critical Windows services.
_PROTECTED_PROCESSES: set[str] = {
    "csrss.exe",
    "lsass.exe",
    "services.exe",
    "smss.exe",
    "svchost.exe",
    "wininit.exe",
    "winlogon.exe",
    "msmpeng.exe",
    "system",
    "registry",
    "memory compression",
}

# Pattern for common system-critical process families
_PROTECTED_PATTERN = re.compile(
    r"^(csrss|lsass|services|smss|svchost|wininit|winlogon|msmpeng|system idle process)\.exe$",
    re.IGNORECASE,
)


class ProcessService:
    """Process enumeration and termination with safety controls."""

    @staticmethod
    def is_protected(name: str) -> bool:
        """Return True if the process name is in the protected blocklist."""
        lower = name.lower()
        return lower in _PROTECTED_PROCESSES or bool(_PROTECTED_PATTERN.match(lower))

    def list_processes(
        self,
        name: str | None = None,
        sort_by: Literal["memory", "cpu", "name"] = "memory",
        limit: int = 20,
    ) -> str:
        """List running processes with optional name filter, sorted and truncated."""
        from tabulate import tabulate

        procs = []
        for p in psutil.process_iter(["pid", "name", "cpu_percent", "memory_info"]):
            try:
                info = p.info
                mem_mb = info["memory_info"].rss / (1024 * 1024) if info["memory_info"] else 0
                procs.append(
                    {
                        "pid": info["pid"],
                        "name": info["name"] or "Unknown",
                        "cpu": info["cpu_percent"] or 0,
                        "mem_mb": round(mem_mb, 1),
                    }
                )
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if name:
            from thefuzz import fuzz

            procs = [p for p in procs if fuzz.partial_ratio(name.lower(), p["name"].lower()) > 60]

        sort_key = {
            "memory": lambda x: x["mem_mb"],
            "cpu": lambda x: x["cpu"],
            "name": lambda x: x["name"].lower(),
        }
        procs.sort(key=sort_key.get(sort_by, sort_key["memory"]), reverse=(sort_by != "name"))
        procs = procs[: max(1, limit)]

        if not procs:
            return f"No processes found{f' matching {name}' if name else ''}."

        table = tabulate(
            [[p["pid"], p["name"], f"{p['cpu']:.1f}%", f"{p['mem_mb']:.1f} MB"] for p in procs],
            headers=["PID", "Name", "CPU%", "Memory"],
            tablefmt="simple",
        )
        return f"Processes ({len(procs)} shown):\n{table}"

    def kill_process(
        self, name: str | None = None, pid: int | None = None, force: bool = False
    ) -> str:
        """Terminate or kill a process by PID or name.

        Protected system processes (csrss, lsass, svchost, etc.) are blocked
        from termination to prevent system instability.
        """
        if pid is None and name is None:
            return "Error: Provide either pid or name parameter for kill mode."

        killed = []
        if pid is not None:
            try:
                p = psutil.Process(pid)
                pname = p.name()
                if self.is_protected(pname):
                    return (
                        f"Refused: {pname} (PID {pid}) is a protected system process "
                        f"and cannot be terminated."
                    )
                if force:
                    p.kill()
                else:
                    p.terminate()
                killed.append(f"{pname} (PID {pid})")
            except psutil.NoSuchProcess:
                return f"No process with PID {pid} found."
            except psutil.AccessDenied:
                return f"Access denied to kill PID {pid}. Try running as administrator."
        else:
            for p in psutil.process_iter(["pid", "name"]):
                try:
                    if p.info["name"] and p.info["name"].lower() == name.lower():
                        if self.is_protected(p.info["name"]):
                            return (
                                f"Refused: {p.info['name']} is a protected system process "
                                f"and cannot be terminated."
                            )
                        if force:
                            p.kill()
                        else:
                            p.terminate()
                        killed.append(f"{p.info['name']} (PID {p.info['pid']})")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

        if not killed:
            return f'No process matching "{name}" found or access denied.'
        return f"{'Force killed' if force else 'Terminated'}: {', '.join(killed)}"
