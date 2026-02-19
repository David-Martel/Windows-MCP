import logging
import os
import sys
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from enum import Enum
from textwrap import dedent

import click
import pyautogui as pg
from dotenv import load_dotenv
from fastmcp import FastMCP
from fastmcp.client.transports import StreamableHttpTransport
from fastmcp.server.proxy import ProxyClient

from windows_mcp.analytics import PostHogAnalytics
from windows_mcp.auth import AuthClient, AuthKeyManager, BearerAuthMiddleware
from windows_mcp.desktop.service import Desktop
from windows_mcp.tools import _state, register_all_tools
from windows_mcp.tools._helpers import _coerce_bool as _coerce_bool  # noqa: F401 -- re-export
from windows_mcp.tools._helpers import _validate_loc as _validate_loc  # noqa: F401 -- re-export
from windows_mcp.watchdog.service import WatchDog

load_dotenv()

logger = logging.getLogger("windows_mcp")


@dataclass
class Config:
    mode: str
    sandbox_id: str = field(default="")
    api_key: str = field(default="")


pg.FAILSAFE = False
pg.PAUSE = 0.05


instructions = dedent("""
Windows MCP server provides tools to interact directly with the Windows desktop,
thus enabling to operate the desktop on the user's behalf.
""")


@asynccontextmanager
async def lifespan(app: FastMCP):
    """Runs initialization code before the server starts and cleanup code after it shuts down."""
    # Initialize components and publish to shared state module
    if os.getenv("ANONYMIZED_TELEMETRY", "true").lower() != "false":
        _state.analytics = PostHogAnalytics()

    _state.desktop = Desktop()
    watchdog = WatchDog()
    _state.screen_size = _state.desktop.get_screen_size()
    watchdog.set_focus_callback(_state.desktop.tree._on_focus_change)

    try:
        watchdog.start()
        yield
    finally:
        if watchdog:
            watchdog.stop()
        if _state.analytics:
            await _state.analytics.close()


mcp = FastMCP(name="windows-mcp", instructions=instructions, lifespan=lifespan)
register_all_tools(mcp)


class Transport(Enum):
    STDIO = "stdio"
    SSE = "sse"
    STREAMABLE_HTTP = "streamable-http"

    def __str__(self):
        return self.value


class Mode(Enum):
    LOCAL = "local"
    REMOTE = "remote"

    def __str__(self):
        return self.value


@click.command()
@click.option(
    "--transport",
    help="The transport layer used by the MCP server.",
    type=click.Choice(
        [Transport.STDIO.value, Transport.SSE.value, Transport.STREAMABLE_HTTP.value]
    ),
    default="stdio",
)
@click.option(
    "--host",
    help="Host to bind the SSE/Streamable HTTP server.",
    default="localhost",
    type=str,
    show_default=True,
)
@click.option(
    "--port",
    help="Port to bind the SSE/Streamable HTTP server.",
    default=8000,
    type=int,
    show_default=True,
)
@click.option(
    "--api-key",
    help="API key for authenticating HTTP/SSE clients. If not set, loads from DPAPI store.",
    default=None,
    type=str,
)
@click.option(
    "--generate-key",
    help="Generate a new API key, store it encrypted (DPAPI), and exit.",
    is_flag=True,
    default=False,
)
@click.option(
    "--rotate-key",
    help="Rotate the stored API key and exit.",
    is_flag=True,
    default=False,
)
def main(transport, host, port, api_key, generate_key, rotate_key):
    logger = logging.getLogger("windows_mcp")

    # Handle key management commands
    if generate_key:
        key = AuthKeyManager.generate_key()
        click.echo(f"API key generated and encrypted with DPAPI.\nKey: {key}")
        click.echo("Save this key -- it will not be shown again.")
        click.echo("Use: windows-mcp --transport sse --api-key <key>")
        sys.exit(0)

    if rotate_key:
        key = AuthKeyManager.rotate_key()
        click.echo(f"API key rotated.\nNew key: {key}")
        click.echo("Update your client configurations with the new key.")
        sys.exit(0)

    # Resolve API key for HTTP transports
    resolved_key = api_key
    if transport in (Transport.SSE.value, Transport.STREAMABLE_HTTP.value):
        if not resolved_key:
            resolved_key = AuthKeyManager.load_key()

        if resolved_key:
            mcp.add_middleware(BearerAuthMiddleware(resolved_key))
            logger.info("Bearer token authentication enabled for %s transport", transport)
        else:
            # Safety: no auth configured -- bind to localhost only
            if host != "localhost" and host != "127.0.0.1":
                logger.warning(
                    "No API key configured. Refusing to bind to %s. "
                    "Use --api-key or --generate-key, or bind to localhost.",
                    host,
                )
                click.echo(
                    f"Error: Cannot bind to {host} without authentication.\n"
                    "Run 'windows-mcp --generate-key' first, or use --host localhost.",
                    err=True,
                )
                sys.exit(1)
            logger.warning("No API key configured. Server accessible without authentication.")

    config = Config(
        mode=os.getenv("MODE", Mode.LOCAL.value).lower(),
        sandbox_id=os.getenv("SANDBOX_ID", ""),
        api_key=os.getenv("API_KEY", ""),
    )
    match config.mode:
        case Mode.LOCAL.value:
            match transport:
                case Transport.STDIO.value:
                    mcp.run(transport=Transport.STDIO.value, show_banner=False)
                case Transport.SSE.value | Transport.STREAMABLE_HTTP.value:
                    mcp.run(transport=transport, host=host, port=port, show_banner=False)
                case _:
                    raise ValueError(f"Invalid transport: {transport}")
        case Mode.REMOTE.value:
            if not config.sandbox_id:
                raise ValueError("SANDBOX_ID is required for MODE: remote")
            if not config.api_key:
                raise ValueError("API_KEY is required for MODE: remote")
            client = AuthClient(api_key=config.api_key, sandbox_id=config.sandbox_id)
            client.authenticate()
            backend = StreamableHttpTransport(url=client.proxy_url, headers=client.proxy_headers)
            proxy_mcp = FastMCP.as_proxy(ProxyClient(backend), name="windows-mcp")
            match transport:
                case Transport.STDIO.value:
                    proxy_mcp.run(transport=Transport.STDIO.value, show_banner=False)
                case Transport.SSE.value | Transport.STREAMABLE_HTTP.value:
                    proxy_mcp.run(transport=transport, host=host, port=port, show_banner=False)
                case _:
                    raise ValueError(f"Invalid transport: {transport}")
        case _:
            raise ValueError(f"Invalid mode: {config.mode}")


if __name__ == "__main__":
    main()
