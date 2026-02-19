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
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

# Default timeout for a single RPC call (seconds).
DEFAULT_CALL_TIMEOUT = 30.0


def _find_worker_exe() -> Path | None:
    """Search for wmcp-worker.exe in known locations."""
    candidates = []

    # Environment variable override
    env_path = os.environ.get("WMCP_WORKER_EXE")
    if env_path:
        candidates.append(Path(env_path))

    # In venv Scripts (derived from sys.prefix, not hardcoded)
    venv_exe = Path(sys.prefix) / "Scripts" / "wmcp-worker.exe"
    candidates.append(venv_exe)

    # Shared Cargo target (via environment variable)
    cargo_target = os.environ.get("CARGO_TARGET_DIR")
    if cargo_target:
        candidates.append(Path(cargo_target) / "release" / "wmcp-worker.exe")

    for p in candidates:
        if p.exists():
            return p
    return None


class NativeWorker:
    """Async wrapper around the wmcp-worker subprocess.

    The worker process is spawned once and reused for the session.
    Each ``call()`` sends a JSON-RPC request and awaits the response.
    """

    def __init__(
        self,
        exe_path: str | Path | None = None,
        verbose: bool = False,
        call_timeout: float = DEFAULT_CALL_TIMEOUT,
    ):
        self._exe_path = Path(exe_path) if exe_path else _find_worker_exe()
        self._verbose = verbose
        self._call_timeout = call_timeout
        self._process: asyncio.subprocess.Process | None = None
        self._request_id = 0
        self._lock = asyncio.Lock()
        self._stderr_task: asyncio.Task | None = None

    @property
    def is_running(self) -> bool:
        return self._process is not None and self._process.returncode is None

    async def _drain_stderr(self):
        """Background task that reads and logs worker stderr."""
        assert self._process is not None
        assert self._process.stderr is not None
        try:
            while True:
                line = await self._process.stderr.readline()
                if not line:
                    break
                logger.debug(
                    "wmcp-worker stderr: %s", line.decode("utf-8", errors="replace").rstrip()
                )
        except asyncio.CancelledError:
            pass

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
        # Start background stderr drain to prevent pipe buffer from filling
        self._stderr_task = asyncio.create_task(self._drain_stderr())
        logger.info("Started wmcp-worker (PID %d)", self._process.pid)

    async def stop(self):
        """Terminate the worker process."""
        if self._stderr_task and not self._stderr_task.done():
            self._stderr_task.cancel()
            try:
                await self._stderr_task
            except asyncio.CancelledError:
                pass
            self._stderr_task = None

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
            TimeoutError: If the worker doesn't respond within the timeout.
        """
        if not self.is_running:
            raise RuntimeError("Worker not running. Call start() first.")

        async with self._lock:
            self._request_id += 1
            request_id = self._request_id
            request = {
                "id": request_id,
                "method": method,
                "params": params if params else {},
            }

            line = json.dumps(request) + "\n"
            self._process.stdin.write(line.encode("utf-8"))
            await self._process.stdin.drain()

            try:
                response_line = await asyncio.wait_for(
                    self._process.stdout.readline(),
                    timeout=self._call_timeout,
                )
            except asyncio.TimeoutError:
                # Worker is in an undefined state after timeout -- the sent
                # request may produce a belated response that would corrupt
                # the next call.  Kill and mark unhealthy so the next call
                # gets a clear error rather than a response-ID mismatch.
                logger.warning(
                    "Worker call '%s' timed out after %ss, killing worker",
                    method, self._call_timeout,
                )
                try:
                    self._process.kill()
                    await self._process.wait()
                except Exception:
                    pass
                self._process = None
                raise TimeoutError(
                    f"Worker call '{method}' timed out after {self._call_timeout}s. "
                    "Worker has been killed; call start() to respawn."
                ) from None

            if not response_line:
                raise RuntimeError("Worker process closed stdout unexpectedly")

            try:
                response = json.loads(response_line.decode("utf-8"))
            except json.JSONDecodeError as exc:
                raise RuntimeError(f"Worker returned non-JSON output: {response_line!r}") from exc

            # Validate response ID matches request
            resp_id = response.get("id")
            if resp_id != request_id:
                raise RuntimeError(f"Response ID mismatch: expected {request_id}, got {resp_id}")

            if response.get("error"):
                raise RuntimeError(f"Worker error: {response['error']}")

            return response.get("result")

    async def __aenter__(self):
        await self.start()
        return self

    async def __aexit__(self, *exc):
        await self.stop()
