"""Subprocess IPC wrapper for the wmcp-worker binary.

Spawns a long-lived Rust worker process that communicates via JSON-RPC
over stdin/stdout.  Provides COM-isolated operations with crash isolation.

Usage::

    from windows_mcp.native_worker import NativeWorker

    worker = NativeWorker()
    await worker.start()
    info = await worker.call("system_info")
    tree = await worker.call("capture_tree", handles=[12345], max_depth=10)
    await worker.stop()
"""

import asyncio
import json
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)


def _find_worker_exe() -> Path | None:
    """Search for wmcp-worker.exe in known locations."""
    candidates = [
        # In venv Scripts
        Path("C:/codedev/windows-mcp/.venv/Scripts/wmcp-worker.exe"),
        # Shared Cargo target
        Path("T:/RustCache/cargo-target/release/wmcp-worker.exe"),
        # Local build
        Path("native/target/release/wmcp-worker.exe"),
    ]

    env_path = os.environ.get("WMCP_WORKER_EXE")
    if env_path:
        candidates.insert(0, Path(env_path))

    for p in candidates:
        if p.exists():
            return p
    return None


class NativeWorker:
    """Async wrapper around the wmcp-worker subprocess.

    The worker process is spawned once and reused for the session.
    Each ``call()`` sends a JSON-RPC request and awaits the response.
    """

    def __init__(self, exe_path: str | Path | None = None, verbose: bool = False):
        self._exe_path = Path(exe_path) if exe_path else _find_worker_exe()
        self._verbose = verbose
        self._process: asyncio.subprocess.Process | None = None
        self._request_id = 0
        self._lock = asyncio.Lock()

    @property
    def is_running(self) -> bool:
        return self._process is not None and self._process.returncode is None

    async def start(self):
        """Spawn the worker process."""
        if self._exe_path is None:
            raise FileNotFoundError(
                "wmcp-worker.exe not found. Build with "
                "`cargo build --release -p wmcp-cli` or set WMCP_WORKER_EXE env var."
            )

        if self.is_running:
            return

        args = [str(self._exe_path)]
        if self._verbose:
            args.append("--verbose")

        self._process = await asyncio.create_subprocess_exec(
            *args,
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        logger.info("Started wmcp-worker (PID %d)", self._process.pid)

    async def stop(self):
        """Terminate the worker process."""
        if self._process and self._process.returncode is None:
            self._process.stdin.close()
            try:
                await asyncio.wait_for(self._process.wait(), timeout=5.0)
            except asyncio.TimeoutError:
                self._process.kill()
                await self._process.wait()
            logger.info("Stopped wmcp-worker")
        self._process = None

    async def call(self, method: str, **params) -> dict | list | str | int | float | bool | None:
        """Send a JSON-RPC request and return the result.

        Args:
            method: The RPC method name (e.g. "system_info", "capture_tree").
            **params: Keyword arguments passed as the ``params`` object.

        Returns:
            The deserialized result value.

        Raises:
            RuntimeError: If the worker returns an error or is not running.
        """
        if not self.is_running:
            raise RuntimeError("Worker not running. Call start() first.")

        async with self._lock:
            self._request_id += 1
            request = {
                "id": self._request_id,
                "method": method,
                "params": params if params else {},
            }

            line = json.dumps(request) + "\n"
            self._process.stdin.write(line.encode("utf-8"))
            await self._process.stdin.drain()

            response_line = await self._process.stdout.readline()
            if not response_line:
                raise RuntimeError("Worker process closed stdout unexpectedly")

            response = json.loads(response_line.decode("utf-8"))

            if response.get("error"):
                raise RuntimeError(f"Worker error: {response['error']}")

            return response.get("result")

    async def __aenter__(self):
        await self.start()
        return self

    async def __aexit__(self, *exc):
        await self.stop()
