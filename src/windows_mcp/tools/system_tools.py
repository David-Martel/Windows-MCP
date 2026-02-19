"""System / management MCP tools.

Registers: App, Shell, File, Scrape, Process, SystemInfo,
Notification, LockScreen, Clipboard, Registry (10 tools).
"""

import os
from typing import Literal

from fastmcp import Context
from mcp.types import ToolAnnotations

from windows_mcp import filesystem
from windows_mcp.analytics import with_analytics
from windows_mcp.tools import _state
from windows_mcp.tools._helpers import _coerce_bool


def register(mcp):  # noqa: C901
    """Register system/management tools on *mcp*."""

    @mcp.tool(
        name="App",
        description="Manages Windows applications with three modes: 'launch' (opens the prescribed application), 'resize' (adjusts active window size/position), 'switch' (brings specific window into focus).",
        annotations=ToolAnnotations(
            title="App",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "App-Tool")
    def app_tool(
        mode: Literal["launch", "resize", "switch"] = "launch",
        name: str | None = None,
        window_loc: list[int] | None = None,
        window_size: list[int] | None = None,
        ctx: Context = None,
    ):
        loc = tuple(window_loc) if window_loc else None
        size = tuple(window_size) if window_size else None
        return _state.desktop.app(mode, name, loc, size)

    @mcp.tool(
        name="Shell",
        description="A comprehensive system tool for executing any PowerShell commands. Use it to navigate the file system, manage files and processes, and execute system-level operations. Capable of accessing web content (e.g., via Invoke-WebRequest), interacting with network resources, and performing complex administrative tasks. This tool provides full access to the underlying operating system capabilities, making it the primary interface for system automation, scripting, and deep system interaction.",
        annotations=ToolAnnotations(
            title="Shell",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=True,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Powershell-Tool")
    def powershell_tool(command: str, timeout: int = 30, ctx: Context = None) -> str:
        if not command or not command.strip():
            return "Error: command must not be empty.\nStatus Code: 1"
        timeout = min(max(timeout, 1), 300)
        try:
            response, status_code = _state.desktop.execute_command(command, timeout)
            return f"Response: {response}\nStatus Code: {status_code}"
        except Exception as e:
            return f"Error executing command: {str(e)}\nStatus Code: 1"

    @mcp.tool(
        name="File",
        description="Manages file system operations with eight modes: 'read' (read text file contents with optional line offset/limit), 'write' (create or overwrite a file, set append=True to append), 'copy' (copy file or directory to destination), 'move' (move or rename file/directory), 'delete' (delete file or directory, set recursive=True for non-empty dirs), 'list' (list directory contents with optional pattern filter), 'search' (find files matching a glob pattern), 'info' (get file/directory metadata like size, dates, type). Relative paths are resolved from the user's Desktop folder. Use absolute paths to access other locations.",
        annotations=ToolAnnotations(
            title="File",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "File-Tool")
    def file_tool(
        mode: Literal["read", "write", "copy", "move", "delete", "list", "search", "info"],
        path: str,
        destination: str | None = None,
        content: str | None = None,
        pattern: str | None = None,
        recursive: bool | str = False,
        append: bool | str = False,
        overwrite: bool | str = False,
        offset: int | None = None,
        limit: int | None = None,
        encoding: str = "utf-8",
        show_hidden: bool | str = False,
        ctx: Context = None,
    ) -> str:
        try:
            from platformdirs import user_desktop_dir

            default_dir = user_desktop_dir()
            if not os.path.isabs(path):
                path = os.path.join(default_dir, path)
            if destination and not os.path.isabs(destination):
                destination = os.path.join(default_dir, destination)

            recursive = _coerce_bool(recursive)
            append = _coerce_bool(append)
            overwrite = _coerce_bool(overwrite)
            show_hidden = _coerce_bool(show_hidden)

            match mode:
                case "read":
                    return filesystem.read_file(path, offset=offset, limit=limit, encoding=encoding)
                case "write":
                    if content is None:
                        return "Error: content parameter is required for write mode."
                    return filesystem.write_file(path, content, append=append, encoding=encoding)
                case "copy":
                    if destination is None:
                        return "Error: destination parameter is required for copy mode."
                    return filesystem.copy_path(path, destination, overwrite=overwrite)
                case "move":
                    if destination is None:
                        return "Error: destination parameter is required for move mode."
                    return filesystem.move_path(path, destination, overwrite=overwrite)
                case "delete":
                    return filesystem.delete_path(path, recursive=recursive)
                case "list":
                    return filesystem.list_directory(
                        path, pattern=pattern, recursive=recursive, show_hidden=show_hidden
                    )
                case "search":
                    if pattern is None:
                        return "Error: pattern parameter is required for search mode."
                    return filesystem.search_files(path, pattern, recursive=recursive)
                case "info":
                    return filesystem.get_file_info(path)
                case _:
                    return (
                        f'Error: Unknown mode "{mode}". '
                        "Use: read, write, copy, move, delete, list, search, info."
                    )
        except Exception as e:
            return f"Error in File tool: {str(e)}"

    @mcp.tool(
        name="Scrape",
        description="Fetch content from a URL or the active browser tab. By default (use_dom=False), performs a lightweight HTTP request to the URL and returns markdown content of complete webpage. Note: Some websites may block automated HTTP requests. If this fails, open the page in a browser and retry with use_dom=True to extract visible text from the active tab's DOM within the viewport using the accessibility tree data.",
        annotations=ToolAnnotations(
            title="Scrape",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=True,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Scrape-Tool")
    def scrape_tool(url: str, use_dom: bool | str = False, ctx: Context = None) -> str:
        use_dom = _coerce_bool(use_dom)
        if not use_dom:
            content = _state.desktop.scrape(url)
            return f"URL:{url}\nContent:\n{content}"

        desktop_state = _state.desktop.get_state(use_vision=False, use_dom=use_dom)
        tree_state = desktop_state.tree_state
        if not tree_state or not tree_state.dom_node:
            return f"No DOM information found. Please open {url} in browser first."
        dom_node = tree_state.dom_node
        vertical_scroll_percent = dom_node.vertical_scroll_percent
        content = "\n".join([node.text for node in tree_state.dom_informative_nodes])
        header_status = "Reached top" if vertical_scroll_percent <= 0 else "Scroll up to see more"
        footer_status = (
            "Reached bottom" if vertical_scroll_percent >= 100 else "Scroll down to see more"
        )
        return f"URL:{url}\nContent:\n{header_status}\n{content}\n{footer_status}"

    @mcp.tool(
        name="Clipboard",
        description='Manages Windows clipboard operations. Use mode="get" to read current clipboard content, mode="set" to set clipboard text.',
        annotations=ToolAnnotations(
            title="Clipboard",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Clipboard-Tool")
    def clipboard_tool(
        mode: Literal["get", "set"], text: str | None = None, ctx: Context = None
    ) -> str:
        try:
            import win32clipboard

            if mode == "get":
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        return f"Clipboard content:\n{data}"
                    else:
                        return "Clipboard is empty or contains non-text data."
                finally:
                    win32clipboard.CloseClipboard()
            elif mode == "set":
                if text is None:
                    return "Error: text parameter required for set mode."
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
                    return f"Clipboard set to: {text[:100]}{'...' if len(text) > 100 else ''}"
                finally:
                    win32clipboard.CloseClipboard()
            else:
                return 'Error: mode must be either "get" or "set".'
        except Exception as e:
            return f"Error managing clipboard: {str(e)}"

    @mcp.tool(
        name="Process",
        description='Manages system processes. Use mode="list" to list running processes with filtering and sorting options. Use mode="kill" to terminate processes by PID or name.',
        annotations=ToolAnnotations(
            title="Process",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Process-Tool")
    def process_tool(
        mode: Literal["list", "kill"],
        name: str | None = None,
        pid: int | None = None,
        sort_by: Literal["memory", "cpu", "name"] = "memory",
        limit: int = 20,
        force: bool | str = False,
        ctx: Context = None,
    ) -> str:
        try:
            if mode == "list":
                return _state.desktop.list_processes(name=name, sort_by=sort_by, limit=limit)
            elif mode == "kill":
                force = _coerce_bool(force)
                return _state.desktop.kill_process(name=name, pid=pid, force=force)
            else:
                return 'Error: mode must be either "list" or "kill".'
        except Exception as e:
            return f"Error managing processes: {str(e)}"

    @mcp.tool(
        name="SystemInfo",
        description="Returns system information including CPU usage, memory usage, disk space, network stats, and uptime. Useful for monitoring system health remotely.",
        annotations=ToolAnnotations(
            title="SystemInfo",
            readOnlyHint=True,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "SystemInfo-Tool")
    def system_info_tool(ctx: Context = None) -> str:
        try:
            return _state.desktop.get_system_info()
        except Exception as e:
            return f"Error getting system info: {str(e)}"

    @mcp.tool(
        name="Notification",
        description="Sends a Windows toast notification with a title and message. Useful for alerting the user remotely.",
        annotations=ToolAnnotations(
            title="Notification",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Notification-Tool")
    def notification_tool(title: str, message: str, ctx: Context = None) -> str:
        try:
            return _state.desktop.send_notification(title, message)
        except Exception as e:
            return f"Error sending notification: {str(e)}"

    @mcp.tool(
        name="LockScreen",
        description="Locks the Windows workstation. Requires the user to enter their password to unlock.",
        annotations=ToolAnnotations(
            title="LockScreen",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "LockScreen-Tool")
    def lock_screen_tool(ctx: Context = None) -> str:
        try:
            return _state.desktop.lock_screen()
        except Exception as e:
            return f"Error locking screen: {str(e)}"

    @mcp.tool(
        name="Registry",
        description='Accesses the Windows Registry. Use mode="get" to read a value, mode="set" to create/update a value, mode="delete" to remove a value or key, mode="list" to list values and sub-keys under a path. Paths use PowerShell format (e.g. "HKCU:\\Software\\MyApp", "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion").',
        annotations=ToolAnnotations(
            title="Registry",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
    )
    @with_analytics(lambda: _state.analytics, "Registry-Tool")
    def registry_tool(
        mode: Literal["get", "set", "delete", "list"],
        path: str,
        name: str | None = None,
        value: str | None = None,
        type: Literal[
            "String", "DWord", "QWord", "Binary", "MultiString", "ExpandString"
        ] = "String",
        ctx: Context = None,
    ) -> str:
        try:
            if mode == "get":
                if name is None:
                    return "Error: name parameter is required for get mode."
                return _state.desktop.registry_get(path=path, name=name)
            elif mode == "set":
                if name is None:
                    return "Error: name parameter is required for set mode."
                if value is None:
                    return "Error: value parameter is required for set mode."
                return _state.desktop.registry_set(path=path, name=name, value=value, reg_type=type)
            elif mode == "delete":
                return _state.desktop.registry_delete(path=path, name=name)
            elif mode == "list":
                return _state.desktop.registry_list(path=path)
            else:
                return 'Error: mode must be "get", "set", "delete", or "list".'
        except Exception as e:
            return f"Error accessing registry: {str(e)}"
