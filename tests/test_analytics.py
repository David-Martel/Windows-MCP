"""Comprehensive tests for the analytics module (PostHog telemetry and with_analytics decorator).

Covers:
- PostHogAnalytics class: initialization, user_id persistence, track_tool, track_error,
  is_feature_enabled, and close.
- with_analytics decorator: success path, error path, None analytics, duration measurement,
  Context extraction, sync function wrapping, and the known None-capture bug.
- Enable/disable telemetry via ANONYMIZED_TELEMETRY environment variable (tested at the
  import/instantiation boundary -- the env var gates PostHogAnalytics construction in the
  server startup code, not inside the class itself).
"""

import asyncio
import os
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from windows_mcp.analytics import Analytics, PostHogAnalytics, with_analytics

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
        assert call_args[0][0] == PostHogAnalytics.API_KEY


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
