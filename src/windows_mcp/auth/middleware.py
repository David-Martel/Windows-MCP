"""Bearer token authentication middleware for FastMCP HTTP/SSE transports."""

import logging

from fastmcp.server.middleware import Middleware, MiddlewareContext

logger = logging.getLogger(__name__)


class BearerAuthMiddleware(Middleware):
    """FastMCP middleware that validates Bearer token on every request.

    Only active for HTTP/SSE transports. Stdio transport bypasses this
    since it's local-only by design.
    """

    def __init__(self, api_key: str):
        self._api_key = api_key

    async def __call__(self, context: MiddlewareContext, call_next):
        """Validate the Bearer token from request metadata."""
        import secrets

        # Try to extract auth token from request context
        token = None
        if hasattr(context, "fastmcp_context") and context.fastmcp_context:
            ctx = context.fastmcp_context
            if ctx.session and ctx.session.client_params:
                # Check for auth in client params or request headers
                meta = getattr(ctx.session.client_params, "meta", None)
                if meta and isinstance(meta, dict):
                    auth_header = meta.get("authorization", meta.get("Authorization", ""))
                    if auth_header.startswith("Bearer "):
                        token = auth_header[7:]

        if token is None:
            # For HTTP transport, also check request-level metadata
            if hasattr(context, "request") and context.request:
                req = context.request
                if hasattr(req, "headers"):
                    auth_header = req.headers.get("authorization", "")
                    if auth_header.startswith("Bearer "):
                        token = auth_header[7:]

        if token is None or not secrets.compare_digest(token, self._api_key):
            logger.warning("Unauthorized request: invalid or missing Bearer token")
            raise PermissionError("Unauthorized: invalid or missing API key")

        return await call_next(context)
