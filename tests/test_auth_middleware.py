"""Tests for BearerAuthMiddleware -- token validation for HTTP transport."""

from unittest.mock import AsyncMock, MagicMock

import pytest

from windows_mcp.auth.middleware import BearerAuthMiddleware


@pytest.fixture
def middleware():
    return BearerAuthMiddleware(api_key="test-secret-key-12345")


@pytest.fixture
def mock_call_next():
    return AsyncMock(return_value="success")


def make_context_with_headers(auth_header=None):
    """Create a mock MiddlewareContext with HTTP request headers."""
    ctx = MagicMock()
    ctx.fastmcp_context = None  # No fastmcp context

    if auth_header is not None:
        request = MagicMock()
        request.headers = {"authorization": auth_header}
        ctx.request = request
    else:
        ctx.request = MagicMock()
        ctx.request.headers = {}

    return ctx


class TestBearerAuthMiddleware:
    async def test_valid_token_passes(self, middleware, mock_call_next):
        ctx = make_context_with_headers("Bearer test-secret-key-12345")
        result = await middleware(ctx, mock_call_next)
        assert result == "success"
        mock_call_next.assert_called_once()

    async def test_missing_token_raises(self, middleware, mock_call_next):
        ctx = make_context_with_headers(None)
        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)
        mock_call_next.assert_not_called()

    async def test_wrong_token_raises(self, middleware, mock_call_next):
        ctx = make_context_with_headers("Bearer wrong-key")
        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)
        mock_call_next.assert_not_called()

    async def test_empty_bearer_raises(self, middleware, mock_call_next):
        ctx = make_context_with_headers("Bearer ")
        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)
        mock_call_next.assert_not_called()

    async def test_non_bearer_scheme_raises(self, middleware, mock_call_next):
        ctx = make_context_with_headers("Basic dXNlcjpwYXNz")
        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)
        mock_call_next.assert_not_called()

    async def test_no_auth_header_at_all_raises(self, middleware, mock_call_next):
        ctx = MagicMock()
        ctx.fastmcp_context = None
        ctx.request = MagicMock()
        ctx.request.headers = {}
        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)

    async def test_token_from_fastmcp_context_meta_authorization_key(
        self, middleware, mock_call_next
    ):
        """Token extracted from fastmcp_context session client_params meta 'authorization'."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = {
            "authorization": "Bearer test-secret-key-12345"
        }
        # No request-level auth needed -- fastmcp_context path wins
        ctx.request = None

        result = await middleware(ctx, mock_call_next)
        assert result == "success"
        mock_call_next.assert_called_once()

    async def test_token_from_fastmcp_context_meta_Authorization_key(
        self, middleware, mock_call_next
    ):
        """Token extracted from fastmcp_context meta using capital-A 'Authorization' key."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = {
            "Authorization": "Bearer test-secret-key-12345"
        }
        ctx.request = None

        result = await middleware(ctx, mock_call_next)
        assert result == "success"
        mock_call_next.assert_called_once()

    async def test_wrong_token_in_fastmcp_context_meta_raises(self, middleware, mock_call_next):
        """Wrong token in fastmcp_context meta should be rejected."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = {"authorization": "Bearer wrong-key"}
        ctx.request = None

        with pytest.raises(PermissionError, match="Unauthorized"):
            await middleware(ctx, mock_call_next)
        mock_call_next.assert_not_called()

    async def test_fastmcp_context_meta_not_dict_falls_through_to_request(
        self, middleware, mock_call_next
    ):
        """When meta is not a dict the code falls through to request-level auth."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = "not-a-dict"

        request = MagicMock()
        request.headers = {"authorization": "Bearer test-secret-key-12345"}
        ctx.request = request

        result = await middleware(ctx, mock_call_next)
        assert result == "success"

    async def test_fastmcp_context_no_meta_falls_through_to_request(
        self, middleware, mock_call_next
    ):
        """When meta attribute is None the code falls through to request-level auth."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = None

        request = MagicMock()
        request.headers = {"authorization": "Bearer test-secret-key-12345"}
        ctx.request = request

        result = await middleware(ctx, mock_call_next)
        assert result == "success"

    async def test_fastmcp_context_no_session_falls_through_to_request(
        self, middleware, mock_call_next
    ):
        """When session is None the code falls through to request-level auth."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = None

        request = MagicMock()
        request.headers = {"authorization": "Bearer test-secret-key-12345"}
        ctx.request = request

        result = await middleware(ctx, mock_call_next)
        assert result == "success"

    async def test_fastmcp_context_bearer_not_in_meta_falls_through(
        self, middleware, mock_call_next
    ):
        """Meta dict present but no Bearer prefix -- falls through to request-level auth."""
        ctx = MagicMock()
        ctx.fastmcp_context = MagicMock()
        ctx.fastmcp_context.session = MagicMock()
        ctx.fastmcp_context.session.client_params = MagicMock()
        ctx.fastmcp_context.session.client_params.meta = {"authorization": "Basic abc123"}

        request = MagicMock()
        request.headers = {"authorization": "Bearer test-secret-key-12345"}
        ctx.request = request

        result = await middleware(ctx, mock_call_next)
        assert result == "success"
