import asyncio
import logging
import os
import tempfile
import time
import traceback
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

T = TypeVar("T")


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
    API_KEY = "phc_uxdCItyVTjXNU0sMPr97dq3tcz39scQNt3qjTYw5vLV"
    HOST = "https://us.i.posthog.com"

    def __init__(self):
        self.client = posthog.Posthog(
            self.API_KEY,
            host=self.HOST,
            disable_geoip=False,
            enable_exception_autocapture=True,
            debug=False,
        )
        self._user_id = None
        self.mcp_interaction_id = f"mcp_{int(time.time() * 1000)}_{os.getpid()}"
        self.mode = os.getenv("MODE", "local").lower()

        if self.client:
            logger.debug(
                "Initialized with user ID: %s and session ID: %s",
                self.user_id, self.mcp_interaction_id,
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


def with_analytics(
    analytics_instance: "Callable[[], Analytics | None] | Analytics | None",
    tool_name: str,
):
    """
    Decorator to wrap tool functions with analytics tracking.

    ``analytics_instance`` may be:
    - A zero-argument callable (e.g. ``lambda: analytics``) whose return value is
      resolved at each call.  Use this when the analytics object is assigned after
      decoration time (the common case in ``__main__.py``).
    - An ``Analytics`` instance resolved before decoration.
    - ``None`` to disable tracking entirely.
    """

    def decorator(func: Callable[..., Awaitable[T]]) -> Callable[..., Awaitable[T]]:
        @wraps(func)
        async def wrapper(*args, **kwargs) -> T:
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
                raise error

        return wrapper

    return decorator
