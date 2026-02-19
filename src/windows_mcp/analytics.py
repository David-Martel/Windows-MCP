import asyncio
import logging
import os
import tempfile
import threading
import time
import traceback
from collections import deque
from functools import wraps
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, Protocol, TypeVar

import posthog
from fastmcp import Context
from uuid_extensions import uuid7str

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# Add handler only if none configured (avoids duplicates on reimport)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter("[%(levelname)s] %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

# Security audit logger: writes tool invocations to a local file
# Enabled by setting WINDOWS_MCP_AUDIT_LOG to a file path
_audit_logger: logging.Logger | None = None
_audit_log_path = os.environ.get("WINDOWS_MCP_AUDIT_LOG", "").strip()
if _audit_log_path:
    _audit_logger = logging.getLogger("windows_mcp.audit")
    _audit_logger.setLevel(logging.INFO)
    _audit_logger.propagate = False
    try:
        Path(_audit_log_path).parent.mkdir(parents=True, exist_ok=True)
        _fh = logging.FileHandler(_audit_log_path, encoding="utf-8")
        _fh.setFormatter(logging.Formatter("%(asctime)s\t%(message)s", datefmt="%Y-%m-%dT%H:%M:%S"))
        _audit_logger.addHandler(_fh)
        logger.info("Audit logging enabled: %s", _audit_log_path)
    except Exception as e:
        logger.warning("Failed to set up audit log at %s: %s", _audit_log_path, e)
        _audit_logger = None

T = TypeVar("T")


# ---------------------------------------------------------------------------
# RateLimiter
# ---------------------------------------------------------------------------

# Per-tool rate limit overrides shipped with the project.
# Format: tool_name -> (max_calls, window_seconds)
_DEFAULT_RATE_LIMITS: dict[str, tuple[int, int]] = {
    "Shell": (10, 60),
    "Registry-Set": (5, 60),
    "Registry-Delete": (5, 60),
}

# Global fallback: 60 calls per 60-second window.
_DEFAULT_RATE_LIMIT_CALLS = 60
_DEFAULT_RATE_LIMIT_WINDOW = 60


def _parse_rate_limits_env(raw: str) -> dict[str, tuple[int, int]]:
    """Parse the ``WINDOWS_MCP_RATE_LIMITS`` environment variable.

    Expected format::

        Tool-A:30:60;Tool-B:5:120

    Each segment is ``<tool_name>:<max_calls>:<window_seconds>``.
    Malformed segments are skipped with a warning.

    Returns:
        Mapping of tool name to ``(max_calls, window_seconds)``.
    """
    result: dict[str, tuple[int, int]] = {}
    for segment in raw.split(";"):
        segment = segment.strip()
        if not segment:
            continue
        parts = segment.split(":")
        if len(parts) != 3:
            logger.warning(
                "RateLimiter: ignoring malformed segment %r (expected tool:limit:window)", segment
            )
            continue
        tool_name, limit_str, window_str = parts
        tool_name = tool_name.strip()
        try:
            limit = int(limit_str.strip())
            window = int(window_str.strip())
        except ValueError:
            logger.warning(
                "RateLimiter: ignoring segment %r -- limit and window must be integers", segment
            )
            continue
        if limit <= 0 or window <= 0:
            logger.warning(
                "RateLimiter: ignoring segment %r -- limit and window must be positive", segment
            )
            continue
        result[tool_name] = (limit, window)
    return result


# Merge defaults with any env-var overrides at module load time.
_env_rate_limits = _parse_rate_limits_env(os.environ.get("WINDOWS_MCP_RATE_LIMITS", "").strip())
_EFFECTIVE_RATE_LIMITS: dict[str, tuple[int, int]] = {**_DEFAULT_RATE_LIMITS, **_env_rate_limits}


class RateLimitExceededError(Exception):
    """Raised when a tool call exceeds its configured rate limit."""


class RateLimiter:
    """Sliding-window rate limiter for MCP tool calls.

    Each tool gets its own deque of call timestamps and a dedicated
    ``threading.Lock`` so concurrent calls never corrupt state.

    Args:
        limits: Mapping of tool name to ``(max_calls, window_seconds)``.
            Tools not present in the mapping use the global defaults.
        default_calls: Global default maximum calls per window.
        default_window: Global default window duration in seconds.

    Example::

        limiter = RateLimiter({"Shell": (10, 60)})
        limiter.check("Shell")   # raises RateLimitExceededError on breach
    """

    def __init__(
        self,
        limits: dict[str, tuple[int, int]] | None = None,
        *,
        default_calls: int = _DEFAULT_RATE_LIMIT_CALLS,
        default_window: int = _DEFAULT_RATE_LIMIT_WINDOW,
    ) -> None:
        self._limits: dict[str, tuple[int, int]] = limits if limits is not None else {}
        self._default_calls = default_calls
        self._default_window = default_window
        # Per-tool state: deque of timestamps + a lock.
        self._timestamps: dict[str, deque[float]] = {}
        self._locks: dict[str, threading.Lock] = {}
        self._meta_lock = threading.Lock()

    def _get_tool_state(self, tool_name: str) -> tuple[deque[float], threading.Lock]:
        """Return (timestamps_deque, lock) for *tool_name*, creating lazily."""
        with self._meta_lock:
            if tool_name not in self._timestamps:
                self._timestamps[tool_name] = deque()
                self._locks[tool_name] = threading.Lock()
            return self._timestamps[tool_name], self._locks[tool_name]

    def _resolve_limit(self, tool_name: str) -> tuple[int, int]:
        """Return ``(max_calls, window_seconds)`` for the given tool."""
        return self._limits.get(tool_name, (self._default_calls, self._default_window))

    def check(self, tool_name: str) -> None:
        """Record a call attempt and raise if the rate limit is exceeded.

        This method is thread-safe.  It slides the observation window by
        evicting timestamps that are older than ``window_seconds`` before
        comparing against ``max_calls``.

        Args:
            tool_name: The MCP tool name being invoked.

        Raises:
            RateLimitExceededError: When the call count in the current window
                exceeds the configured maximum.
        """
        max_calls, window_seconds = self._resolve_limit(tool_name)
        timestamps, lock = self._get_tool_state(tool_name)

        now = time.monotonic()
        cutoff = now - window_seconds

        with lock:
            # Evict timestamps outside the sliding window.
            while timestamps and timestamps[0] < cutoff:
                timestamps.popleft()

            if len(timestamps) >= max_calls:
                # Calculate how long the caller must wait before the oldest
                # call ages out of the window.
                oldest = timestamps[0]
                retry_after = window_seconds - (now - oldest)
                raise RateLimitExceededError(
                    f"Rate limit exceeded for tool '{tool_name}': "
                    f"{max_calls} calls per {window_seconds}s allowed. "
                    f"Retry after {retry_after:.1f}s."
                )

            timestamps.append(now)


# Module-level singleton populated from project defaults + env overrides.
_rate_limiter = RateLimiter(limits=_EFFECTIVE_RATE_LIMITS)


_USE_MODULE_RATE_LIMITER = object()  # sentinel: resolve _rate_limiter via sys.modules at call time


# ---------------------------------------------------------------------------
# Tool permission manifest (allow/deny lists)
# ---------------------------------------------------------------------------


class ToolNotAllowedError(Exception):
    """Raised when a tool call is blocked by the permission manifest."""


def _parse_tool_list(raw: str) -> set[str]:
    """Parse a comma-separated tool name list (case-insensitive)."""
    return {t.strip().lower() for t in raw.split(",") if t.strip()}


# WINDOWS_MCP_ALLOW: if set, only listed tools are available (allowlist).
# WINDOWS_MCP_DENY: if set, listed tools are blocked (denylist).
# If both are set, ALLOW takes precedence (only allowed tools minus denied).
_allow_tools: set[str] | None = None
_deny_tools: set[str] = set()

_raw_allow = os.environ.get("WINDOWS_MCP_ALLOW", "").strip()
_raw_deny = os.environ.get("WINDOWS_MCP_DENY", "").strip()
if _raw_allow:
    _allow_tools = _parse_tool_list(_raw_allow)
    logger.info("Tool allowlist active: %s", _allow_tools)
if _raw_deny:
    _deny_tools = _parse_tool_list(_raw_deny)
    logger.info("Tool denylist active: %s", _deny_tools)


def check_tool_permission(tool_name: str) -> None:
    """Check whether ``tool_name`` is allowed by the permission manifest.

    Raises:
        ToolNotAllowedError: If the tool is blocked by allow/deny rules.
    """
    name_lower = tool_name.lower()

    if _allow_tools is not None and name_lower not in _allow_tools:
        raise ToolNotAllowedError(
            f"Tool '{tool_name}' is not in the allowlist (WINDOWS_MCP_ALLOW). "
            f"Allowed: {', '.join(sorted(_allow_tools))}."
        )

    if name_lower in _deny_tools:
        raise ToolNotAllowedError(
            f"Tool '{tool_name}' is blocked by the denylist (WINDOWS_MCP_DENY)."
        )


# ---------------------------------------------------------------------------
# Analytics protocol and PostHog implementation
# ---------------------------------------------------------------------------


class Analytics(Protocol):
    async def track_tool(self, tool_name: str, result: Dict[str, Any]) -> None:
        """Tracks the execution of a tool."""
        ...

    async def track_error(self, error: Exception, context: Dict[str, Any]) -> None:
        """Tracks an error that occurred during the execution of a tool."""
        ...

    async def is_feature_enabled(self, feature: str) -> bool:
        """Checks if a feature flag is enabled."""
        ...

    async def close(self) -> None:
        """Closes the analytics client."""
        ...


class PostHogAnalytics:
    TEMP_FOLDER = Path(tempfile.gettempdir())
    _DEFAULT_API_KEY = "phc_uxdCItyVTjXNU0sMPr97dq3tcz39scQNt3qjTYw5vLV"
    HOST = "https://us.i.posthog.com"

    def __init__(self):
        api_key = os.environ.get("POSTHOG_API_KEY", self._DEFAULT_API_KEY)
        self.client = posthog.Posthog(
            api_key,
            host=self.HOST,
            disable_geoip=True,
            enable_exception_autocapture=False,
            debug=False,
        )
        self._user_id = None
        self.mcp_interaction_id = f"mcp_{int(time.time() * 1000)}_{os.getpid()}"
        self.mode = os.getenv("MODE", "local").lower()

        if self.client:
            logger.debug(
                "Initialized with user ID: %s and session ID: %s",
                self.user_id,
                self.mcp_interaction_id,
            )

    @property
    def user_id(self) -> str:
        if self._user_id:
            return self._user_id

        user_id_file = self.TEMP_FOLDER / ".windows-mcp-user-id"
        if user_id_file.exists():
            self._user_id = user_id_file.read_text(encoding="utf-8").strip()
        else:
            self._user_id = uuid7str()
            try:
                user_id_file.write_text(self._user_id, encoding="utf-8")
            except Exception as e:
                logger.warning("Could not persist user ID: %s", e)

        return self._user_id

    async def track_tool(self, tool_name: str, result: Dict[str, Any]) -> None:
        if self.client:
            self.client.capture(
                distinct_id=self.user_id,
                event="tool_executed",
                properties={
                    "tool_name": tool_name,
                    "session_id": self.mcp_interaction_id,
                    "mode": self.mode,
                    "process_person_profile": True,
                    **result,
                },
            )

        duration = result.get("duration_ms", 0)
        success_mark = "SUCCESS" if result.get("success") else "FAILED"
        logger.info("%s: %s (%dms)", tool_name, success_mark, duration)

    async def track_error(self, error: Exception, context: Dict[str, Any]) -> None:
        if self.client:
            tb_str = "".join(traceback.format_exception(type(error), error, error.__traceback__))
            self.client.capture(
                distinct_id=self.user_id,
                event="exception",
                properties={
                    "exception": str(error),
                    "traceback": tb_str,
                    "session_id": self.mcp_interaction_id,
                    "mode": self.mode,
                    "process_person_profile": True,
                    **context,
                },
            )

        logger.error("ERROR in %s: %s", context.get("tool_name"), error)

    async def is_feature_enabled(self, feature: str) -> bool:
        if not self.client:
            return False
        return self.client.is_feature_enabled(feature, self.user_id)

    async def close(self) -> None:
        if self.client:
            self.client.shutdown()
            logger.debug("Closed analytics")


# ---------------------------------------------------------------------------
# with_analytics decorator
# ---------------------------------------------------------------------------


def with_analytics(
    analytics_instance: "Callable[[], Analytics | None] | Analytics | None",
    tool_name: str,
    *,
    rate_limiter: "RateLimiter | None | object" = _USE_MODULE_RATE_LIMITER,
):
    """
    Decorator to wrap tool functions with analytics tracking and rate limiting.

    ``analytics_instance`` may be:
    - A zero-argument callable (e.g. ``lambda: analytics``) whose return value is
      resolved at each call.  Use this when the analytics object is assigned after
      decoration time (the common case in ``__main__.py``).
    - An ``Analytics`` instance resolved before decoration.
    - ``None`` to disable tracking entirely.

    ``rate_limiter`` defaults to the module-level ``_rate_limiter`` singleton,
    resolved at each call so that test patches to ``_rate_limiter`` are respected.
    Pass ``rate_limiter=None`` to disable rate limiting for a specific tool.
    Pass an explicit ``RateLimiter`` instance to use a custom limiter.
    """

    def decorator(func: Callable[..., Awaitable[T]]) -> Callable[..., Awaitable[T]]:
        @wraps(func)
        async def wrapper(*args, **kwargs) -> T:
            # Resolve the rate limiter at call time.
            # When the sentinel is present, look up the module-level
            # ``_rate_limiter`` via sys.modules so that both test patches and
            # importlib.reload() side-effects are always visible.
            # When rate_limiter is an explicit RateLimiter or None, use it as-is.
            if rate_limiter is _USE_MODULE_RATE_LIMITER or not isinstance(
                rate_limiter, (RateLimiter, type(None))
            ):
                import sys

                mod = sys.modules.get(__name__)
                effective_limiter = getattr(mod, "_rate_limiter", None) if mod else None
            else:
                effective_limiter = rate_limiter  # type: ignore[assignment]

            # Enforce rate limit before doing any work.
            if effective_limiter is not None:
                effective_limiter.check(tool_name)

            # Enforce tool permission manifest.
            check_tool_permission(tool_name)

            # Resolve the analytics instance at call time so that late
            # assignment (e.g. inside an async lifespan) is picked up.
            # Objects that implement the Analytics protocol (have track_tool)
            # are used directly.  A plain callable without track_tool is
            # treated as a provider/factory and called to obtain the instance.
            if analytics_instance is None or hasattr(analytics_instance, "track_tool"):
                instance = analytics_instance
            else:
                instance = analytics_instance()

            start = time.time()

            # Capture client info from Context passed as argument
            client_data = {}
            try:
                ctx = next((arg for arg in args if isinstance(arg, Context)), None)
                if not ctx:
                    ctx = next(
                        (val for val in kwargs.values() if isinstance(val, Context)),
                        None,
                    )

                if (
                    ctx
                    and ctx.session
                    and ctx.session.client_params
                    and ctx.session.client_params.clientInfo
                ):
                    info = ctx.session.client_params.clientInfo
                    client_data["client_name"] = info.name
                    client_data["client_version"] = info.version
            except Exception:
                pass

            try:
                if asyncio.iscoroutinefunction(func):
                    result = await func(*args, **kwargs)
                else:
                    # Run sync function in thread to avoid blocking loop
                    result = await asyncio.to_thread(func, *args, **kwargs)

                duration_ms = int((time.time() - start) * 1000)

                if instance:
                    try:
                        await instance.track_tool(
                            tool_name,
                            {"duration_ms": duration_ms, "success": True, **client_data},
                        )
                    except Exception:
                        logger.debug("Analytics track_tool failed for %s", tool_name)

                if _audit_logger:
                    _audit_logger.info("OK\t%s\t%dms", tool_name, duration_ms)

                return result
            except Exception as error:
                duration_ms = int((time.time() - start) * 1000)
                if instance:
                    try:
                        await instance.track_error(
                            error,
                            {
                                "tool_name": tool_name,
                                "duration_ms": duration_ms,
                                **client_data,
                            },
                        )
                    except Exception:
                        logger.debug("Analytics track_error failed for %s", tool_name)

                if _audit_logger:
                    _audit_logger.info(
                        "ERR\t%s\t%dms\t%s", tool_name, duration_ms, type(error).__name__
                    )

                raise error

        return wrapper

    return decorator
