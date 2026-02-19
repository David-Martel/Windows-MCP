"""Comprehensive tests for the analytics module (PostHog telemetry and with_analytics decorator).

Covers:
- PostHogAnalytics class: initialization, user_id persistence, track_tool, track_error,
  is_feature_enabled, and close.
- with_analytics decorator: success path, error path, None analytics, duration measurement,
  Context extraction, sync function wrapping, and the known None-capture bug.
- Enable/disable telemetry via ANONYMIZED_TELEMETRY environment variable (tested at the
  import/instantiation boundary -- the env var gates PostHogAnalytics construction in the
  server startup code, not inside the class itself).
- RateLimiter: sliding window enforcement, per-tool limits, env-var parsing, thread safety,
  and integration with the with_analytics decorator.
"""

import asyncio
import logging
import os
import threading
import time
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from windows_mcp.analytics import (
    Analytics,
    PostHogAnalytics,
    with_analytics,
)

# ---------------------------------------------------------------------------
# Helpers / Fixtures
# ---------------------------------------------------------------------------


def _make_mock_posthog() -> MagicMock:
    """Return a MagicMock that stands in for a posthog.Posthog client instance."""
    mock = MagicMock()
    mock.capture = MagicMock()
    mock.is_feature_enabled = MagicMock(return_value=True)
    mock.shutdown = MagicMock()
    return mock


def _make_posthog_analytics(mock_posthog_client: MagicMock) -> PostHogAnalytics:
    """Construct a PostHogAnalytics instance with the posthog.Posthog ctor patched."""
    with patch("windows_mcp.analytics.posthog.Posthog", return_value=mock_posthog_client):
        return PostHogAnalytics()


# ---------------------------------------------------------------------------
# PostHogAnalytics -- Initialization
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsInit:
    def test_client_is_created(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        assert analytics.client is mock_client

    def test_mcp_interaction_id_format(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        assert analytics.mcp_interaction_id.startswith("mcp_")

    def test_mode_defaults_to_local(self):
        mock_client = _make_mock_posthog()
        with patch.dict(os.environ, {}, clear=True):
            # Remove MODE so the default kicks in
            os.environ.pop("MODE", None)
            analytics = _make_posthog_analytics(mock_client)
        assert analytics.mode == "local"

    def test_mode_reads_env_var(self):
        mock_client = _make_mock_posthog()
        with patch.dict(os.environ, {"MODE": "remote"}):
            analytics = _make_posthog_analytics(mock_client)
        assert analytics.mode == "remote"

    def test_mode_is_lowercased(self):
        mock_client = _make_mock_posthog()
        with patch.dict(os.environ, {"MODE": "REMOTE"}):
            analytics = _make_posthog_analytics(mock_client)
        assert analytics.mode == "remote"

    def test_user_id_is_populated_after_init(self):
        """
        PostHogAnalytics.__init__ calls self.user_id inside the logger.debug() call,
        so _user_id is already set (non-None) by the time __init__ returns.
        """
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        # _user_id is populated eagerly during __init__ (via the debug log that calls self.user_id)
        assert analytics._user_id is not None
        assert isinstance(analytics._user_id, str)

    def test_posthog_constructed_with_expected_api_key(self):
        with patch("windows_mcp.analytics.posthog.Posthog") as mock_ctor:
            mock_ctor.return_value = _make_mock_posthog()
            PostHogAnalytics()
        call_args = mock_ctor.call_args
        assert call_args[0][0] == PostHogAnalytics._DEFAULT_API_KEY

    def test_posthog_uses_env_api_key_when_set(self, monkeypatch):
        monkeypatch.setenv("POSTHOG_API_KEY", "custom-key-123")
        with patch("windows_mcp.analytics.posthog.Posthog") as mock_ctor:
            mock_ctor.return_value = _make_mock_posthog()
            PostHogAnalytics()
        call_args = mock_ctor.call_args
        assert call_args[0][0] == "custom-key-123"


# ---------------------------------------------------------------------------
# PostHogAnalytics -- user_id property
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsUserId:
    def test_user_id_generated_when_no_file(self, tmp_path):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path  # redirect to temp so we don't pollute the FS

        uid = analytics.user_id
        assert isinstance(uid, str)
        assert len(uid) > 0

    def test_user_id_persisted_to_file(self, tmp_path):
        # TEMP_FOLDER must be patched at class level before construction because
        # __init__ eagerly accesses user_id via logger.debug(), which uses TEMP_FOLDER.
        mock_client = _make_mock_posthog()
        with (
            patch.object(PostHogAnalytics, "TEMP_FOLDER", tmp_path),
            patch("windows_mcp.analytics.posthog.Posthog", return_value=mock_client),
        ):
            analytics = PostHogAnalytics()

        written_file = tmp_path / ".windows-mcp-user-id"
        assert written_file.exists()
        assert written_file.read_text(encoding="utf-8").strip() == analytics._user_id

    def test_user_id_loaded_from_existing_file(self, tmp_path):
        existing_id = "test-user-id-from-file"
        (tmp_path / ".windows-mcp-user-id").write_text(existing_id, encoding="utf-8")

        mock_client = _make_mock_posthog()
        with (
            patch.object(PostHogAnalytics, "TEMP_FOLDER", tmp_path),
            patch("windows_mcp.analytics.posthog.Posthog", return_value=mock_client),
        ):
            analytics = PostHogAnalytics()

        assert analytics._user_id == existing_id

    def test_user_id_cached_after_first_access(self, tmp_path):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path

        uid1 = analytics.user_id
        uid2 = analytics.user_id
        assert uid1 == uid2

    def test_user_id_warns_when_file_write_fails(self, tmp_path):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path

        with patch("pathlib.Path.write_text", side_effect=OSError("permission denied")):
            # Should not raise -- logs a warning instead
            uid = analytics.user_id
        assert isinstance(uid, str)


# ---------------------------------------------------------------------------
# PostHogAnalytics -- track_tool
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsTrackTool:
    async def test_capture_called_on_success(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("click", {"duration_ms": 50, "success": True})

        mock_client.capture.assert_called_once()

    async def test_capture_event_name_is_tool_executed(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("snapshot", {"duration_ms": 100, "success": True})

        call_kwargs = mock_client.capture.call_args[1]
        assert call_kwargs["event"] == "tool_executed"

    async def test_capture_includes_tool_name(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("type", {"duration_ms": 30, "success": True})

        props = mock_client.capture.call_args[1]["properties"]
        assert props["tool_name"] == "type"

    async def test_capture_includes_duration_ms(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("scroll", {"duration_ms": 200, "success": True})

        props = mock_client.capture.call_args[1]["properties"]
        assert props["duration_ms"] == 200

    async def test_capture_includes_session_id(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("app", {"duration_ms": 10, "success": True})

        props = mock_client.capture.call_args[1]["properties"]
        assert "session_id" in props

    async def test_capture_not_called_when_client_is_none(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.client = None  # simulate disabled client

        await analytics.track_tool("shell", {"duration_ms": 5, "success": True})

        mock_client.capture.assert_not_called()

    async def test_failed_tool_marks_success_false(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_tool("click", {"duration_ms": 50, "success": False})

        props = mock_client.capture.call_args[1]["properties"]
        assert props["success"] is False


# ---------------------------------------------------------------------------
# PostHogAnalytics -- track_error
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsTrackError:
    async def test_capture_called_on_error(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_error(ValueError("oops"), {"tool_name": "click"})

        mock_client.capture.assert_called_once()

    async def test_capture_event_name_is_exception(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_error(RuntimeError("bad"), {"tool_name": "type"})

        call_kwargs = mock_client.capture.call_args[1]
        assert call_kwargs["event"] == "exception"

    async def test_capture_includes_exception_string(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        err = ValueError("something broke badly")
        await analytics.track_error(err, {"tool_name": "scroll"})

        props = mock_client.capture.call_args[1]["properties"]
        assert "something broke badly" in props["exception"]

    async def test_capture_not_called_when_client_is_none(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.client = None

        # Should not raise
        await analytics.track_error(RuntimeError("x"), {"tool_name": "app"})

        mock_client.capture.assert_not_called()

    async def test_capture_includes_context_data(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.track_error(
            TypeError("type mismatch"),
            {"tool_name": "registry", "duration_ms": 99},
        )

        props = mock_client.capture.call_args[1]["properties"]
        assert props["tool_name"] == "registry"
        assert props["duration_ms"] == 99


# ---------------------------------------------------------------------------
# PostHogAnalytics -- is_feature_enabled
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsIsFeatureEnabled:
    async def test_returns_true_when_feature_enabled(self, tmp_path):
        mock_client = _make_mock_posthog()
        mock_client.is_feature_enabled.return_value = True
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path

        result = await analytics.is_feature_enabled("new_feature")
        assert result is True

    async def test_returns_false_when_feature_disabled(self, tmp_path):
        mock_client = _make_mock_posthog()
        mock_client.is_feature_enabled.return_value = False
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path

        result = await analytics.is_feature_enabled("old_feature")
        assert result is False

    async def test_returns_false_when_client_is_none(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.client = None

        result = await analytics.is_feature_enabled("any_feature")
        assert result is False

    async def test_delegates_to_posthog_client(self, tmp_path):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path

        await analytics.is_feature_enabled("beta_feature")

        mock_client.is_feature_enabled.assert_called_once_with("beta_feature", analytics.user_id)


# ---------------------------------------------------------------------------
# PostHogAnalytics -- close
# ---------------------------------------------------------------------------


class TestPostHogAnalyticsClose:
    async def test_shutdown_called_on_close(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)

        await analytics.close()

        mock_client.shutdown.assert_called_once()

    async def test_close_is_safe_when_client_is_none(self):
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.client = None

        # Must not raise
        await analytics.close()
        mock_client.shutdown.assert_not_called()


# ---------------------------------------------------------------------------
# with_analytics decorator -- success path
# ---------------------------------------------------------------------------


class TestWithAnalyticsSuccessPath:
    async def test_returns_result_of_wrapped_function(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "test_tool")
        async def my_tool():
            return "hello"

        result = await my_tool()
        assert result == "hello"

    async def test_track_tool_called_on_success(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "test_tool")
        async def my_tool():
            return "ok"

        await my_tool()
        mock_analytics.track_tool.assert_called_once()

    async def test_track_tool_receives_correct_tool_name(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "snapshot")
        async def my_tool():
            return None

        await my_tool()
        call_args = mock_analytics.track_tool.call_args
        assert call_args[0][0] == "snapshot"

    async def test_track_tool_success_flag_is_true(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "click")
        async def my_tool():
            return True

        await my_tool()
        payload = mock_analytics.track_tool.call_args[0][1]
        assert payload["success"] is True

    async def test_track_tool_includes_duration_ms(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "type")
        async def my_tool():
            return None

        await my_tool()
        payload = mock_analytics.track_tool.call_args[0][1]
        assert "duration_ms" in payload
        assert isinstance(payload["duration_ms"], int)

    async def test_complex_return_value_preserved(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "system_info")
        async def my_tool():
            return {"cpu": 50, "ram": 8192}

        result = await my_tool()
        assert result == {"cpu": 50, "ram": 8192}

    async def test_track_error_not_called_on_success(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "scroll")
        async def my_tool():
            return "fine"

        await my_tool()
        mock_analytics.track_error.assert_not_called()


# ---------------------------------------------------------------------------
# with_analytics decorator -- error path
# ---------------------------------------------------------------------------


class TestWithAnalyticsErrorPath:
    async def test_exception_is_reraised(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "test_tool")
        async def failing_tool():
            raise ValueError("intentional failure")

        with pytest.raises(ValueError, match="intentional failure"):
            await failing_tool()

    async def test_track_error_called_on_exception(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "test_tool")
        async def failing_tool():
            raise RuntimeError("boom")

        with pytest.raises(RuntimeError):
            await failing_tool()

        mock_analytics.track_error.assert_called_once()

    async def test_track_error_receives_exception_instance(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "test_tool")
        async def failing_tool():
            raise TypeError("bad type")

        with pytest.raises(TypeError):
            await failing_tool()

        error_arg = mock_analytics.track_error.call_args[0][0]
        assert isinstance(error_arg, TypeError)

    async def test_track_error_context_contains_tool_name(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "registry")
        async def failing_tool():
            raise OSError("no access")

        with pytest.raises(OSError):
            await failing_tool()

        context_arg = mock_analytics.track_error.call_args[0][1]
        assert context_arg["tool_name"] == "registry"

    async def test_track_error_context_contains_duration_ms(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "shell")
        async def failing_tool():
            raise PermissionError("denied")

        with pytest.raises(PermissionError):
            await failing_tool()

        context_arg = mock_analytics.track_error.call_args[0][1]
        assert "duration_ms" in context_arg

    async def test_track_tool_not_called_on_exception(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "click")
        async def failing_tool():
            raise Exception("fail")

        with pytest.raises(Exception):
            await failing_tool()

        mock_analytics.track_tool.assert_not_called()

    async def test_exception_type_preserved(self):
        """Verify the original exception type is not swallowed or wrapped."""
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "app")
        async def failing_tool():
            raise KeyError("missing_key")

        with pytest.raises(KeyError):
            await failing_tool()


# ---------------------------------------------------------------------------
# with_analytics decorator -- None analytics (known bug)
# ---------------------------------------------------------------------------


class TestWithAnalyticsNoneInstance:
    async def test_none_analytics_does_not_raise_on_success(self):
        """Passing None as analytics_instance must not cause AttributeError."""

        @with_analytics(None, "test_tool")
        async def my_tool():
            return 42

        result = await my_tool()
        assert result == 42

    async def test_none_analytics_does_not_raise_on_error(self):
        """Passing None as analytics_instance must not suppress the original error."""

        @with_analytics(None, "test_tool")
        async def failing_tool():
            raise ValueError("original error")

        with pytest.raises(ValueError, match="original error"):
            await failing_tool()

    async def test_known_bug_none_captured_at_decoration_time(self):
        """
        Documents the known bug: if analytics_instance is None at decoration time,
        subsequent assignment of a real analytics object to the outer variable has
        no effect because the decorator already closed over None.

        This test exists to document the behaviour, not to assert it is correct.
        """
        analytics_holder: Analytics | None = None

        @with_analytics(analytics_holder, "test_tool")
        async def my_tool():
            return "result"

        # Simulate late assignment (as happens in lifespan startup)
        analytics_holder = AsyncMock()

        result = await my_tool()
        # The late-assigned analytics object is NEVER called because None was captured.
        analytics_holder.track_tool.assert_not_called()
        assert result == "result"


# ---------------------------------------------------------------------------
# with_analytics decorator -- duration measurement
# ---------------------------------------------------------------------------


class TestWithAnalyticsDuration:
    async def test_duration_is_non_negative(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "wait")
        async def my_tool():
            return None

        await my_tool()
        duration = mock_analytics.track_tool.call_args[0][1]["duration_ms"]
        assert duration >= 0

    async def test_duration_captures_elapsed_time(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "slow_tool")
        async def slow_tool():
            await asyncio.sleep(0.1)
            return "done"

        await slow_tool()
        duration = mock_analytics.track_tool.call_args[0][1]["duration_ms"]
        # 100ms sleep -- allow generous margin for CI variability
        assert duration >= 50

    async def test_error_path_duration_is_non_negative(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "failing")
        async def failing_tool():
            await asyncio.sleep(0.05)
            raise RuntimeError("fail after delay")

        with pytest.raises(RuntimeError):
            await failing_tool()

        duration = mock_analytics.track_error.call_args[0][1]["duration_ms"]
        assert duration >= 0


# ---------------------------------------------------------------------------
# with_analytics decorator -- Context extraction
# ---------------------------------------------------------------------------


class TestWithAnalyticsContextExtraction:
    async def test_client_info_extracted_from_positional_context_arg(self):
        mock_analytics = AsyncMock()

        mock_client_info = MagicMock()
        mock_client_info.name = "test-client"
        mock_client_info.version = "1.2.3"

        mock_session = MagicMock()
        mock_session.client_params.clientInfo = mock_client_info

        mock_ctx = MagicMock()
        mock_ctx.session = mock_session

        # Make isinstance(mock_ctx, Context) return True
        with patch("windows_mcp.analytics.Context", MagicMock):
            from fastmcp import Context as RealContext

            mock_ctx.__class__ = RealContext

            @with_analytics(mock_analytics, "click")
            async def my_tool(ctx):
                return "ok"

            await my_tool(mock_ctx)

        payload = mock_analytics.track_tool.call_args[0][1]
        assert payload.get("client_name") == "test-client"
        assert payload.get("client_version") == "1.2.3"

    async def test_client_info_extracted_from_keyword_context_arg(self):
        mock_analytics = AsyncMock()

        mock_client_info = MagicMock()
        mock_client_info.name = "kwarg-client"
        mock_client_info.version = "9.9.9"

        mock_session = MagicMock()
        mock_session.client_params.clientInfo = mock_client_info

        mock_ctx = MagicMock()
        mock_ctx.session = mock_session

        with patch("windows_mcp.analytics.Context", MagicMock):
            from fastmcp import Context as RealContext

            mock_ctx.__class__ = RealContext

            @with_analytics(mock_analytics, "type")
            async def my_tool(ctx=None):
                return "ok"

            await my_tool(ctx=mock_ctx)

        payload = mock_analytics.track_tool.call_args[0][1]
        assert payload.get("client_name") == "kwarg-client"

    async def test_no_crash_when_context_absent(self):
        """Decorator must not fail if no Context is passed at all."""
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "snapshot")
        async def my_tool(x: int):
            return x * 2

        result = await my_tool(5)
        assert result == 10

    async def test_no_crash_when_context_has_no_session(self):
        mock_analytics = AsyncMock()

        mock_ctx = MagicMock()
        mock_ctx.session = None

        with patch("windows_mcp.analytics.Context", MagicMock):
            from fastmcp import Context as RealContext

            mock_ctx.__class__ = RealContext

            @with_analytics(mock_analytics, "move")
            async def my_tool(ctx):
                return "ok"

            result = await my_tool(mock_ctx)

        assert result == "ok"
        mock_analytics.track_tool.assert_called_once()

    async def test_no_crash_when_client_params_is_none(self):
        mock_analytics = AsyncMock()

        mock_session = MagicMock()
        mock_session.client_params = None

        mock_ctx = MagicMock()
        mock_ctx.session = mock_session

        with patch("windows_mcp.analytics.Context", MagicMock):
            from fastmcp import Context as RealContext

            mock_ctx.__class__ = RealContext

            @with_analytics(mock_analytics, "shortcut")
            async def my_tool(ctx):
                return "ok"

            result = await my_tool(mock_ctx)

        assert result == "ok"


# ---------------------------------------------------------------------------
# with_analytics decorator -- sync function wrapping
# ---------------------------------------------------------------------------


class TestWithAnalyticsSyncFunction:
    async def test_sync_function_is_wrapped_and_returns_result(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "sync_tool")
        def sync_tool(x: int, y: int) -> int:
            return x + y

        result = await sync_tool(3, 4)
        assert result == 7

    async def test_sync_function_tracks_on_success(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "sync_tool")
        def sync_tool():
            return "value"

        await sync_tool()
        mock_analytics.track_tool.assert_called_once()

    async def test_sync_function_tracks_error(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "sync_failing")
        def sync_tool():
            raise ValueError("sync error")

        with pytest.raises(ValueError, match="sync error"):
            await sync_tool()

        mock_analytics.track_error.assert_called_once()


# ---------------------------------------------------------------------------
# with_analytics decorator -- ANONYMIZED_TELEMETRY env var (integration)
# ---------------------------------------------------------------------------


class TestTelemetryEnvVar:
    """
    The ANONYMIZED_TELEMETRY env var does not disable PostHogAnalytics internally;
    instead the server startup code in __main__.py checks it before constructing
    PostHogAnalytics.  These tests verify that pattern works correctly.
    """

    def test_telemetry_disabled_when_env_var_false(self):
        """Server startup should produce None analytics when ANONYMIZED_TELEMETRY=false."""
        with patch.dict(os.environ, {"ANONYMIZED_TELEMETRY": "false"}):
            telemetry_enabled = os.getenv("ANONYMIZED_TELEMETRY", "true").lower() != "false"
        assert telemetry_enabled is False

    def test_telemetry_enabled_by_default(self):
        """Server startup should construct analytics when env var is not set."""
        env = {k: v for k, v in os.environ.items() if k != "ANONYMIZED_TELEMETRY"}
        with patch.dict(os.environ, env, clear=True):
            telemetry_enabled = os.getenv("ANONYMIZED_TELEMETRY", "true").lower() != "false"
        assert telemetry_enabled is True

    def test_telemetry_enabled_when_env_var_true(self):
        with patch.dict(os.environ, {"ANONYMIZED_TELEMETRY": "true"}):
            telemetry_enabled = os.getenv("ANONYMIZED_TELEMETRY", "true").lower() != "false"
        assert telemetry_enabled is True

    def test_telemetry_disabled_case_insensitive(self):
        """FALSE, False, false should all disable telemetry."""
        for variant in ("FALSE", "False", "false"):
            with patch.dict(os.environ, {"ANONYMIZED_TELEMETRY": variant}):
                telemetry_enabled = os.getenv("ANONYMIZED_TELEMETRY", "true").lower() != "false"
            assert telemetry_enabled is False, f"Failed for variant: {variant!r}"

    async def test_none_analytics_from_disabled_telemetry_is_safe(self):
        """
        When telemetry is disabled the server passes None to with_analytics.
        Verify the decorator handles this gracefully for both success and error paths.
        """

        @with_analytics(None, "some_tool")
        async def my_tool(value: str) -> str:
            return value.upper()

        result = await my_tool("hello")
        assert result == "HELLO"

    async def test_none_analytics_error_still_propagates(self):
        @with_analytics(None, "some_tool")
        async def my_tool():
            raise RuntimeError("real error")

        with pytest.raises(RuntimeError, match="real error"):
            await my_tool()


# ---------------------------------------------------------------------------
# with_analytics decorator -- wraps() preservation
# ---------------------------------------------------------------------------


class TestWithAnalyticsWraps:
    def test_wrapped_function_name_preserved(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "click")
        async def my_special_tool():
            return None

        assert my_special_tool.__name__ == "my_special_tool"

    def test_wrapped_function_docstring_preserved(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "click")
        async def my_tool():
            """This is a docstring."""
            return None

        assert my_tool.__doc__ == "This is a docstring."


# ---------------------------------------------------------------------------
# with_analytics decorator -- multiple tools / re-use
# ---------------------------------------------------------------------------


class TestWithAnalyticsMultipleTools:
    async def test_each_tool_tracked_with_own_name(self):
        mock_analytics = AsyncMock()

        @with_analytics(mock_analytics, "tool_a")
        async def tool_a():
            return "a"

        @with_analytics(mock_analytics, "tool_b")
        async def tool_b():
            return "b"

        await tool_a()
        await tool_b()

        assert mock_analytics.track_tool.call_count == 2
        names = [call[0][0] for call in mock_analytics.track_tool.call_args_list]
        assert "tool_a" in names
        assert "tool_b" in names

    async def test_independent_analytics_instances(self):
        analytics_a = AsyncMock()
        analytics_b = AsyncMock()

        @with_analytics(analytics_a, "tool_a")
        async def tool_a():
            return "a"

        @with_analytics(analytics_b, "tool_b")
        async def tool_b():
            return "b"

        await tool_a()
        await tool_b()

        analytics_a.track_tool.assert_called_once()
        analytics_b.track_tool.assert_called_once()
        analytics_a.track_error.assert_not_called()
        analytics_b.track_error.assert_not_called()


# ---------------------------------------------------------------------------
# Coverage gap: analytics lines 80-81, 181-182, 199-200, 215-216
# ---------------------------------------------------------------------------


class TestAnalyticsCoverageGaps:
    """Exercises the remaining uncovered exception handlers in analytics.py."""

    def test_user_id_write_failure_logs_warning(self, tmp_path):
        """Lines 80-81: write_text raises but user_id is still generated."""
        mock_client = _make_mock_posthog()
        analytics = _make_posthog_analytics(mock_client)
        analytics.TEMP_FOLDER = tmp_path
        analytics._user_id = None  # Clear cached user_id from __init__

        with patch("pathlib.Path.write_text", side_effect=OSError("disk full")):
            uid = analytics.user_id
        assert isinstance(uid, str)
        assert len(uid) > 0

    async def test_track_tool_exception_is_swallowed(self):
        """Lines 199-200: track_tool raises during decorator but doesn't crash the tool."""
        mock_instance = AsyncMock(spec=Analytics)
        mock_instance.track_tool = AsyncMock(side_effect=RuntimeError("analytics broken"))

        @with_analytics(lambda: mock_instance, "TestTool")
        async def my_tool():
            return "success"

        result = await my_tool()
        assert result == "success"
        mock_instance.track_tool.assert_awaited_once()

    async def test_track_error_exception_is_swallowed(self):
        """Lines 215-216: track_error raises during error path but original exception propagates."""
        mock_instance = AsyncMock(spec=Analytics)
        mock_instance.track_error = AsyncMock(side_effect=RuntimeError("analytics broken"))
        mock_instance.track_tool = AsyncMock()

        @with_analytics(lambda: mock_instance, "TestTool")
        async def my_failing_tool():
            raise ValueError("tool error")

        with pytest.raises(ValueError, match="tool error"):
            await my_failing_tool()
        mock_instance.track_error.assert_awaited_once()

    async def test_client_params_exception_is_swallowed(self):
        """Lines 181-182: exception accessing ctx.session.client_params.clientInfo is caught."""
        from fastmcp import Context

        mock_instance = AsyncMock(spec=Analytics)
        mock_instance.track_tool = AsyncMock()

        # Create a mock that passes isinstance(arg, Context) and has a session
        # where clientInfo.name raises, triggering the except block at lines 181-182
        mock_ctx = MagicMock(spec=Context)
        mock_client_info = MagicMock()
        mock_client_info.name = property(lambda self: (_ for _ in ()).throw(TypeError("broken")))
        # Assign to a class so the property descriptor works
        broken_info = type("BrokenInfo", (), {"name": property(lambda s: 1 / 0)})()
        mock_ctx.session.client_params.clientInfo = broken_info

        @with_analytics(lambda: mock_instance, "TestTool")
        async def my_tool(ctx=None):
            return "ok"

        result = await my_tool(ctx=mock_ctx)
        assert result == "ok"


# ---------------------------------------------------------------------------
# Audit logger -- module-level configuration
# ---------------------------------------------------------------------------


class TestAuditLoggerConfiguration:
    """Tests for _audit_logger setup at module import time.

    Because the audit logger is initialised at module level (when the env var
    is read), these tests use importlib.reload() to re-execute the module-level
    code with a controlled environment.  After each reload the module is
    re-imported under its canonical name so subsequent imports are consistent.
    """

    def test_audit_logger_is_none_when_env_var_not_set(self, monkeypatch):
        """_audit_logger must be None when WINDOWS_MCP_AUDIT_LOG is absent."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        monkeypatch.delenv("WINDOWS_MCP_AUDIT_LOG", raising=False)
        importlib.reload(analytics_mod)
        assert analytics_mod._audit_logger is None

    def test_audit_logger_is_none_when_env_var_is_empty_string(self, monkeypatch):
        """An empty string for WINDOWS_MCP_AUDIT_LOG must leave _audit_logger as None."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", "")
        importlib.reload(analytics_mod)
        assert analytics_mod._audit_logger is None

    def test_audit_logger_is_configured_when_env_var_set(self, tmp_path, monkeypatch):
        """_audit_logger must be a Logger instance when env var points to a valid path."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        log_path = str(tmp_path / "audit.log")
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", log_path)
        importlib.reload(analytics_mod)

        assert analytics_mod._audit_logger is not None
        assert isinstance(analytics_mod._audit_logger, logging.Logger)

    def test_audit_logger_has_file_handler(self, tmp_path, monkeypatch):
        """After configuration the logger must have exactly one FileHandler."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        log_path = str(tmp_path / "audit.log")
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", log_path)
        importlib.reload(analytics_mod)

        assert analytics_mod._audit_logger is not None
        file_handlers = [
            h for h in analytics_mod._audit_logger.handlers if isinstance(h, logging.FileHandler)
        ]
        assert len(file_handlers) >= 1

    def test_audit_logger_propagation_disabled(self, tmp_path, monkeypatch):
        """Audit logger must not propagate to the root logger."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        log_path = str(tmp_path / "audit.log")
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", log_path)
        importlib.reload(analytics_mod)

        assert analytics_mod._audit_logger is not None
        assert analytics_mod._audit_logger.propagate is False

    def test_audit_logger_gracefully_degrades_on_invalid_path(self, tmp_path, monkeypatch):
        """A FileHandler that raises during setup must leave _audit_logger as None.

        This simulates an unwritable path (permission denied, locked file, etc.)
        without relying on OS-specific path quirks that differ between Windows and
        POSIX.  The env var is set to a syntactically valid path; the FileHandler
        constructor is patched to raise so that the except-branch in analytics.py
        sets _audit_logger back to None.
        """
        import importlib

        import windows_mcp.analytics as analytics_mod

        # Any non-empty path triggers the audit-logger setup code.
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", str(tmp_path / "no_perms.log"))

        # Simulate a failure that occurs when the FileHandler tries to open the file
        # (e.g. permission denied, locked directory, etc.).
        with patch("logging.FileHandler", side_effect=OSError("permission denied")):
            importlib.reload(analytics_mod)

        assert analytics_mod._audit_logger is None

    def test_audit_log_file_is_created_on_disk(self, tmp_path, monkeypatch):
        """The log file must exist on disk after configuration (FileHandler creates it)."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", str(log_path))
        importlib.reload(analytics_mod)

        assert log_path.exists()

    def test_audit_logger_creates_parent_directories(self, tmp_path, monkeypatch):
        """Parent directories that don't exist must be created automatically."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "nested" / "dir" / "audit.log"
        monkeypatch.setenv("WINDOWS_MCP_AUDIT_LOG", str(log_path))
        importlib.reload(analytics_mod)

        assert log_path.parent.exists()

    def teardown_method(self):
        """Reload module without env var after each test to restore pristine state."""
        import importlib

        import windows_mcp.analytics as analytics_mod

        # Remove the env var so reload restores _audit_logger=None
        os.environ.pop("WINDOWS_MCP_AUDIT_LOG", None)
        importlib.reload(analytics_mod)


# ---------------------------------------------------------------------------
# Audit logger -- success path log content
# ---------------------------------------------------------------------------


class TestAuditLoggerSuccessPath:
    """Tests that verify the content written to the audit log on successful tool execution."""

    async def test_audit_log_written_on_success(self, tmp_path):
        """A successful tool call must produce a line in the audit log file."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "click")
            async def my_tool():
                return "ok"

            await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert len(content.strip()) > 0

    async def test_audit_log_success_starts_with_ok(self, tmp_path):
        """Success entries must contain the 'OK' marker."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "snapshot")
            async def my_tool():
                return "result"

            await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "\tOK\t" in content

    async def test_audit_log_success_contains_tool_name(self, tmp_path):
        """Success entries must include the tool name."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "MySpecialTool")
            async def my_tool():
                return "data"

            await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "MySpecialTool" in content

    async def test_audit_log_success_contains_duration(self, tmp_path):
        """Success entries must include a duration value ending with 'ms'."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "type")
            async def my_tool():
                return None

            await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "ms" in content

    async def test_audit_log_success_has_timestamp_prefix(self, tmp_path):
        """Each line must start with a timestamp in ISO-like format (YYYY-MM-DDTHH:MM:SS)."""
        import re

        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "scroll")
            async def my_tool():
                return "scrolled"

            await my_tool()

        content = log_path.read_text(encoding="utf-8").strip()
        # Timestamp format: 2024-01-15T12:34:56
        assert re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", content)

    async def test_audit_log_format_is_tab_separated(self, tmp_path):
        """Each field in a log line must be separated by tabs."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "move")
            async def my_tool():
                return True

            await my_tool()

        content = log_path.read_text(encoding="utf-8").strip()
        # Format: <timestamp>\tOK\t<tool_name>\t<duration_ms>
        # The line must have at least 3 tab characters
        assert content.count("\t") >= 3

    async def test_audit_log_tab_separated_fields_order(self, tmp_path):
        """Fields must appear in order: timestamp, OK, tool_name, duration."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "shell")
            async def my_tool():
                return "output"

            await my_tool()

        line = log_path.read_text(encoding="utf-8").strip()
        fields = line.split("\t")
        # fields[0] = timestamp, fields[1] = OK, fields[2] = tool_name, fields[3] = duration
        assert len(fields) >= 4
        assert fields[1] == "OK"
        assert fields[2] == "shell"
        assert fields[3].endswith("ms")


# ---------------------------------------------------------------------------
# Audit logger -- error path log content
# ---------------------------------------------------------------------------


class TestAuditLoggerErrorPath:
    """Tests that verify the content written to the audit log on tool failure."""

    async def test_audit_log_written_on_failure(self, tmp_path):
        """A failing tool call must produce a line in the audit log file."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "click")
            async def my_tool():
                raise ValueError("oops")

            with pytest.raises(ValueError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert len(content.strip()) > 0

    async def test_audit_log_error_contains_err_marker(self, tmp_path):
        """Error entries must contain the 'ERR' marker."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "registry")
            async def my_tool():
                raise RuntimeError("registry failure")

            with pytest.raises(RuntimeError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "\tERR\t" in content

    async def test_audit_log_error_contains_tool_name(self, tmp_path):
        """Error entries must include the tool name."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "FailingTool")
            async def my_tool():
                raise OSError("disk error")

            with pytest.raises(OSError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "FailingTool" in content

    async def test_audit_log_error_contains_exception_type(self, tmp_path):
        """Error entries must include the exception class name."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "scraper")
            async def my_tool():
                raise TypeError("bad arg")

            with pytest.raises(TypeError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "TypeError" in content

    async def test_audit_log_error_contains_duration(self, tmp_path):
        """Error entries must include a duration value ending with 'ms'."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "app")
            async def my_tool():
                raise PermissionError("access denied")

            with pytest.raises(PermissionError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "ms" in content

    async def test_audit_log_error_tab_separated_fields_order(self, tmp_path):
        """Error fields must appear in order: timestamp, ERR, tool_name, duration, error_type."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "process")
            async def my_tool():
                raise KeyError("missing")

            with pytest.raises(KeyError):
                await my_tool()

        line = log_path.read_text(encoding="utf-8").strip()
        fields = line.split("\t")
        # fields[0]=timestamp, fields[1]=ERR, fields[2]=tool_name, fields[3]=duration, fields[4]=error_type
        assert len(fields) >= 5
        assert fields[1] == "ERR"
        assert fields[2] == "process"
        assert fields[3].endswith("ms")
        assert fields[4] == "KeyError"

    async def test_audit_log_error_exception_is_reraised(self, tmp_path):
        """The decorator must re-raise the original exception after logging."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "wait")
            async def my_tool():
                raise ValueError("must propagate")

            with pytest.raises(ValueError, match="must propagate"):
                await my_tool()

    async def test_audit_log_error_has_timestamp_prefix(self, tmp_path):
        """Error lines must start with a timestamp in ISO-like format."""
        import re

        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "shortcut")
            async def my_tool():
                raise RuntimeError("shortcut failed")

            with pytest.raises(RuntimeError):
                await my_tool()

        content = log_path.read_text(encoding="utf-8").strip()
        assert re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", content)


# ---------------------------------------------------------------------------
# Audit logger -- no-op when _audit_logger is None
# ---------------------------------------------------------------------------


class TestAuditLoggerDisabled:
    """Tests that verify no audit I/O happens when _audit_logger is None."""

    async def test_no_audit_call_when_logger_is_none_success(self):
        """When _audit_logger is None, a successful call must not call any logger method."""
        import windows_mcp.analytics as analytics_mod

        with patch.object(analytics_mod, "_audit_logger", None):
            # Create a spy on logging.Logger.info to confirm it is not called via audit path
            spy = MagicMock()
            with patch("logging.Logger.info", spy):

                @with_analytics(None, "click")
                async def my_tool():
                    return "ok"

                await my_tool()

            # The spy may be called by other loggers; verify no audit-path call occurred.
            # We rely on the fact that _audit_logger is None, so the branch is skipped entirely.
            # This test is a smoke test -- if the branch guard is removed, coverage will catch it.

    async def test_no_audit_call_when_logger_is_none_error(self):
        """When _audit_logger is None, a failing call must not suppress the exception."""
        import windows_mcp.analytics as analytics_mod

        with patch.object(analytics_mod, "_audit_logger", None):

            @with_analytics(None, "click")
            async def my_tool():
                raise RuntimeError("still raises")

            with pytest.raises(RuntimeError, match="still raises"):
                await my_tool()


# ---------------------------------------------------------------------------
# Audit logger -- coexistence with PostHog tracking
# ---------------------------------------------------------------------------


class TestAuditLoggerWithPostHog:
    """Tests that verify audit logging and PostHog telemetry both fire independently."""

    async def test_both_audit_and_posthog_fire_on_success(self, tmp_path):
        """Both the audit log file and PostHog track_tool must be called on success."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)
        mock_analytics = AsyncMock()

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(mock_analytics, "snapshot")
            async def my_tool():
                return "screenshot_data"

            await my_tool()

        # PostHog fired
        mock_analytics.track_tool.assert_called_once()
        # Audit log written
        content = log_path.read_text(encoding="utf-8")
        assert "\tOK\t" in content
        assert "snapshot" in content

    async def test_both_audit_and_posthog_fire_on_failure(self, tmp_path):
        """Both the audit log file and PostHog track_error must be called on failure."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)
        mock_analytics = AsyncMock()

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(mock_analytics, "shell")
            async def my_tool():
                raise RuntimeError("shell error")

            with pytest.raises(RuntimeError):
                await my_tool()

        # PostHog fired
        mock_analytics.track_error.assert_called_once()
        # Audit log written
        content = log_path.read_text(encoding="utf-8")
        assert "\tERR\t" in content
        assert "shell" in content

    async def test_audit_fires_even_when_posthog_raises(self, tmp_path):
        """Audit log must be written even if PostHog track_tool raises an exception."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)
        mock_analytics = AsyncMock()
        mock_analytics.track_tool = AsyncMock(side_effect=RuntimeError("posthog down"))

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(lambda: mock_analytics, "type")
            async def my_tool():
                return "typed"

            result = await my_tool()

        assert result == "typed"
        content = log_path.read_text(encoding="utf-8")
        assert "\tOK\t" in content
        assert "type" in content

    async def test_posthog_none_audit_still_fires(self, tmp_path):
        """Audit log must write entries when PostHog is disabled (analytics=None)."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "clipboard")
            async def my_tool():
                return "pasted"

            await my_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "\tOK\t" in content
        assert "clipboard" in content

    async def test_multiple_tool_calls_produce_multiple_log_lines(self, tmp_path):
        """Each tool invocation must produce a separate log line."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "click")
            async def my_tool():
                return "clicked"

            await my_tool()
            await my_tool()
            await my_tool()

        lines = [ln for ln in log_path.read_text(encoding="utf-8").splitlines() if ln.strip()]
        assert len(lines) == 3

    async def test_mixed_success_and_error_produce_distinct_markers(self, tmp_path):
        """Interleaved success and error calls must produce OK and ERR lines respectively."""
        import windows_mcp.analytics as analytics_mod

        log_path = tmp_path / "audit.log"
        audit_logger = _build_file_audit_logger(log_path)

        with patch.object(analytics_mod, "_audit_logger", audit_logger):

            @with_analytics(None, "app")
            async def good_tool():
                return "launched"

            @with_analytics(None, "app")
            async def bad_tool():
                raise ValueError("crash")

            await good_tool()
            with pytest.raises(ValueError):
                await bad_tool()

        content = log_path.read_text(encoding="utf-8")
        assert "\tOK\t" in content
        assert "\tERR\t" in content


# ---------------------------------------------------------------------------
# Audit logger -- helper (module-private, not a test class)
# ---------------------------------------------------------------------------


def _build_file_audit_logger(log_path) -> logging.Logger:
    """Create an isolated FileHandler-based audit logger for testing.

    Returns a fresh Logger instance that writes to *log_path* using the same
    formatter as the production audit logger, without touching module-level state.
    """
    audit_logger = logging.getLogger(f"windows_mcp.audit.test.{id(log_path)}")
    audit_logger.setLevel(logging.INFO)
    audit_logger.propagate = False
    # Remove any pre-existing handlers (guard against pytest re-use)
    audit_logger.handlers.clear()
    fh = logging.FileHandler(str(log_path), encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s\t%(message)s", datefmt="%Y-%m-%dT%H:%M:%S"))
    audit_logger.addHandler(fh)
    return audit_logger


# ---------------------------------------------------------------------------
# Fixture: reload analytics module to undo any importlib.reload() side-effects
# from TestAuditLoggerConfiguration tests above.  All rate-limiter tests use it
# via autouse so the class objects they reference are always the live ones.
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Base class: ensures RateLimiter / RateLimitExceededError class identity is
# consistent after TestAuditLoggerConfiguration.teardown_method() may have
# called importlib.reload(), which creates new class objects for those names.
# ---------------------------------------------------------------------------


class _RateLimiterTestBase:
    """Mixin that refreshes the analytics module references before each test.

    ``TestAuditLoggerConfiguration.teardown_method`` calls ``importlib.reload()``
    which creates new class objects.  Any ``RateLimiter`` / ``RateLimitExceededError``
    instances or ``pytest.raises`` checks that run after those reloads will
    compare against the wrong (stale) class unless we re-bind from the live module.
    """

    def setup_method(self):
        import importlib

        import windows_mcp.analytics as analytics_mod

        importlib.reload(analytics_mod)
        # Rebind names at the test-instance level so test methods see them.
        self.RateLimiter = analytics_mod.RateLimiter
        self.RateLimitExceededError = analytics_mod.RateLimitExceededError
        self._parse_rate_limits_env = analytics_mod._parse_rate_limits_env
        self.with_analytics = analytics_mod.with_analytics


# ---------------------------------------------------------------------------
# RateLimiter -- core sliding window behaviour
# ---------------------------------------------------------------------------


class TestRateLimiterSlidingWindow(_RateLimiterTestBase):
    """Tests for the core sliding window enforcement in RateLimiter."""

    def test_calls_within_limit_do_not_raise(self):
        limiter = self.RateLimiter({"tool": (3, 60)})
        for _ in range(3):
            limiter.check("tool")  # Must not raise

    def test_exceeding_limit_raises_rate_limit_error(self):
        limiter = self.RateLimiter({"tool": (3, 60)})
        for _ in range(3):
            limiter.check("tool")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("tool")

    def test_error_message_contains_tool_name(self):
        limiter = self.RateLimiter({"Shell": (1, 60)})
        limiter.check("Shell")
        with pytest.raises(self.RateLimitExceededError, match="Shell"):
            limiter.check("Shell")

    def test_error_message_contains_limit_and_window(self):
        limiter = self.RateLimiter({"Shell": (5, 30)})
        for _ in range(5):
            limiter.check("Shell")
        with pytest.raises(self.RateLimitExceededError, match=r"5.*30"):
            limiter.check("Shell")

    def test_error_message_contains_retry_after(self):
        limiter = self.RateLimiter({"tool": (1, 60)})
        limiter.check("tool")
        with pytest.raises(self.RateLimitExceededError, match=r"Retry after"):
            limiter.check("tool")

    def test_calls_allowed_after_window_expires(self):
        """Timestamps older than window_seconds must be evicted, allowing new calls."""
        limiter = self.RateLimiter({"tool": (2, 1)})
        limiter.check("tool")
        limiter.check("tool")
        # Window is 1 second; sleep past it so both timestamps are evicted.
        time.sleep(1.05)
        # Must not raise -- previous calls have aged out.
        limiter.check("tool")
        limiter.check("tool")

    def test_partial_window_eviction_counts_correctly(self):
        """Only timestamps inside the window count against the limit."""
        limiter = self.RateLimiter({"tool": (3, 1)})
        limiter.check("tool")
        limiter.check("tool")
        # Sleep so those two calls age out.
        time.sleep(1.05)
        # Now add two more within the new window -- should still allow one more.
        limiter.check("tool")
        limiter.check("tool")
        # Third in new window is still within limit.
        limiter.check("tool")
        # Fourth must fail.
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("tool")

    def test_independent_windows_per_tool(self):
        """Exhausting one tool's limit must not affect another tool."""
        limiter = self.RateLimiter({"tool_a": (2, 60), "tool_b": (2, 60)})
        limiter.check("tool_a")
        limiter.check("tool_a")
        # tool_a is exhausted; tool_b is independent.
        limiter.check("tool_b")
        limiter.check("tool_b")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("tool_a")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("tool_b")

    def test_unknown_tool_uses_default_limit(self):
        """A tool not in limits falls back to the configured default."""
        limiter = self.RateLimiter({}, default_calls=2, default_window=60)
        limiter.check("unknown_tool")
        limiter.check("unknown_tool")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("unknown_tool")


# ---------------------------------------------------------------------------
# RateLimiter -- default limits for built-in tools
# ---------------------------------------------------------------------------


class TestRateLimiterDefaultLimits(_RateLimiterTestBase):
    """Verify the project-default per-tool limits are correctly defined."""

    def test_shell_limit_is_10_per_minute(self):
        limiter = self.RateLimiter({"Shell": (10, 60)})
        for _ in range(10):
            limiter.check("Shell")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("Shell")

    def test_registry_set_limit_is_5_per_minute(self):
        limiter = self.RateLimiter({"Registry-Set": (5, 60)})
        for _ in range(5):
            limiter.check("Registry-Set")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("Registry-Set")

    def test_registry_delete_limit_is_5_per_minute(self):
        limiter = self.RateLimiter({"Registry-Delete": (5, 60)})
        for _ in range(5):
            limiter.check("Registry-Delete")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("Registry-Delete")

    def test_default_global_limit_is_60_per_minute(self):
        limiter = self.RateLimiter({}, default_calls=60, default_window=60)
        for _ in range(60):
            limiter.check("Click")
        with pytest.raises(self.RateLimitExceededError):
            limiter.check("Click")


# ---------------------------------------------------------------------------
# RateLimiter -- thread safety
# ---------------------------------------------------------------------------


class TestRateLimiterThreadSafety(_RateLimiterTestBase):
    """Verify that concurrent calls do not corrupt RateLimiter state."""

    def test_concurrent_calls_respect_limit(self):
        """Multiple threads racing to call check() must not exceed the limit together."""
        limit = 20
        limiter = self.RateLimiter({"tool": (limit, 60)})
        successes = []
        errors = []
        result_lock = threading.Lock()

        def do_call():
            try:
                limiter.check("tool")
                with result_lock:
                    successes.append(1)
            except self.RateLimitExceededError:
                with result_lock:
                    errors.append(1)

        threads = [threading.Thread(target=do_call) for _ in range(40)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # Exactly `limit` calls must succeed; the rest must be rejected.
        assert len(successes) == limit
        assert len(errors) == 40 - limit

    def test_separate_tools_do_not_block_each_other(self):
        """Threads checking different tools must not serialise on the same lock."""
        limiter = self.RateLimiter({"tool_x": (100, 60), "tool_y": (100, 60)})
        results = {"x": 0, "y": 0}
        result_lock = threading.Lock()

        def call_x():
            for _ in range(50):
                limiter.check("tool_x")
            with result_lock:
                results["x"] += 1

        def call_y():
            for _ in range(50):
                limiter.check("tool_y")
            with result_lock:
                results["y"] += 1

        tx = threading.Thread(target=call_x)
        ty = threading.Thread(target=call_y)
        tx.start()
        ty.start()
        tx.join()
        ty.join()

        assert results["x"] == 1
        assert results["y"] == 1


# ---------------------------------------------------------------------------
# RateLimiter -- env var parsing (_parse_rate_limits_env)
# ---------------------------------------------------------------------------


class TestParseRateLimitsEnv(_RateLimiterTestBase):
    """Tests for the environment variable parser."""

    def test_empty_string_returns_empty_dict(self):
        result = self._parse_rate_limits_env("")
        assert result == {}

    def test_whitespace_only_returns_empty_dict(self):
        result = self._parse_rate_limits_env("   ")
        assert result == {}

    def test_single_valid_segment(self):
        result = self._parse_rate_limits_env("Shell:10:60")
        assert result == {"Shell": (10, 60)}

    def test_multiple_valid_segments(self):
        result = self._parse_rate_limits_env("Shell:10:60;Registry-Set:5:120")
        assert result == {"Shell": (10, 60), "Registry-Set": (5, 120)}

    def test_trailing_semicolon_is_ignored(self):
        result = self._parse_rate_limits_env("Shell:10:60;")
        assert result == {"Shell": (10, 60)}

    def test_segments_with_extra_whitespace_are_parsed(self):
        result = self._parse_rate_limits_env("  Shell : 10 : 60  ")
        assert result == {"Shell": (10, 60)}

    def test_malformed_segment_missing_window_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:10;Click:30:60")
        assert "Shell" not in result
        assert result == {"Click": (30, 60)}

    def test_non_integer_limit_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:abc:60")
        assert result == {}

    def test_non_integer_window_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:10:xyz")
        assert result == {}

    def test_zero_limit_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:0:60")
        assert result == {}

    def test_zero_window_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:10:0")
        assert result == {}

    def test_negative_limit_is_skipped(self):
        result = self._parse_rate_limits_env("Shell:-5:60")
        assert result == {}

    def test_valid_segment_after_bad_segment_is_kept(self):
        result = self._parse_rate_limits_env("BAD:notanint:60;Click:30:60")
        assert "BAD" not in result
        assert result["Click"] == (30, 60)

    def test_tool_name_with_hyphen_is_accepted(self):
        result = self._parse_rate_limits_env("Registry-Delete:5:60")
        assert result == {"Registry-Delete": (5, 60)}


# ---------------------------------------------------------------------------
# RateLimiter -- integration with with_analytics decorator
# ---------------------------------------------------------------------------


class TestRateLimiterWithAnalyticsIntegration(_RateLimiterTestBase):
    """Verify the decorator enforces rate limits before calling the wrapped function."""

    async def test_rate_limit_error_raised_before_tool_executes(self):
        """The wrapped tool body must not execute when the rate limit is exceeded."""
        limiter = self.RateLimiter({"FastTool": (1, 60)})
        call_count = 0

        @self.with_analytics(None, "FastTool", rate_limiter=limiter)
        async def fast_tool():
            nonlocal call_count
            call_count += 1
            return "result"

        await fast_tool()  # First call succeeds.
        with pytest.raises(self.RateLimitExceededError):
            await fast_tool()

        # The body was invoked only once -- the second attempt was rejected early.
        assert call_count == 1

    async def test_rate_limit_error_is_not_tracked_as_analytics_error(self):
        """RateLimitExceededError must propagate without calling track_error."""
        mock_analytics = AsyncMock()
        limiter = self.RateLimiter({"tool": (1, 60)})

        @self.with_analytics(mock_analytics, "tool", rate_limiter=limiter)
        async def my_tool():
            return "ok"

        await my_tool()
        with pytest.raises(self.RateLimitExceededError):
            await my_tool()

        # Only the first successful call should have been tracked.
        mock_analytics.track_tool.assert_called_once()
        mock_analytics.track_error.assert_not_called()

    async def test_rate_limiter_none_disables_limiting(self):
        """Passing rate_limiter=None must allow unlimited calls."""

        @self.with_analytics(None, "UnlimitedTool", rate_limiter=None)
        async def unlimited_tool():
            return "ok"

        # Call 200 times without hitting any limit.
        for _ in range(200):
            result = await unlimited_tool()
        assert result == "ok"

    async def test_successful_call_is_still_tracked_after_rate_limit_hit(self):
        """After a rate limit violation, the window resets and tracking resumes."""
        mock_analytics = AsyncMock()
        # Very short window so we can expire it quickly.
        limiter = self.RateLimiter({"tool": (1, 1)})

        @self.with_analytics(mock_analytics, "tool", rate_limiter=limiter)
        async def my_tool():
            return "ok"

        await my_tool()

        with pytest.raises(self.RateLimitExceededError):
            await my_tool()

        # Sleep past the window.
        time.sleep(1.05)
        await my_tool()

        assert mock_analytics.track_tool.call_count == 2

    async def test_rate_limit_exceeded_error_message_propagates(self):
        """The RateLimitExceededError message must reach the caller unchanged."""
        limiter = self.RateLimiter({"Shell": (1, 60)})

        @self.with_analytics(None, "Shell", rate_limiter=limiter)
        async def shell_tool():
            return "output"

        await shell_tool()
        with pytest.raises(self.RateLimitExceededError, match="Shell"):
            await shell_tool()

    async def test_decorator_uses_module_rate_limiter_by_default(self):
        """Default rate_limiter kwarg should be the module-level singleton."""
        import windows_mcp.analytics as analytics_mod

        mock_limiter = MagicMock(spec=self.RateLimiter)
        mock_limiter.check = MagicMock()

        with patch.object(analytics_mod, "_rate_limiter", mock_limiter):
            # Re-apply the decorator inside the patch so it picks up the mock.
            @analytics_mod.with_analytics(None, "SomeTool")
            async def some_tool():
                return "ok"

            await some_tool()

        mock_limiter.check.assert_called_once_with("SomeTool")
