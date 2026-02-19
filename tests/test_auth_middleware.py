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
