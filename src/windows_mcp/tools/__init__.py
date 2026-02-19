"""MCP tool registration package.

Call ``register_all_tools(mcp)`` after creating the FastMCP instance
to register all 24 tool handlers.
"""

from windows_mcp.tools import input_tools, state_tools, system_tools


def register_all_tools(mcp):
    """Register all tool handlers on *mcp*."""
    input_tools.register(mcp)
    state_tools.register(mcp)
    system_tools.register(mcp)
