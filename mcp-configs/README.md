# Windows-MCP Client Configurations

MCP client config files for connecting AI agents to the Windows-MCP server at `C:/codedev/windows-mcp`.

## Files

| File | Target client | Deploy location |
|------|--------------|-----------------|
| `claude-desktop.json` | Claude Desktop | `%APPDATA%\Claude\claude_desktop_config.json` |
| `claude-code.json` | Claude Code | `.claude/mcp.json` in any project |
| `pc-ai-agent.json` | PC-AI agent | `C:\Users\david\PC_AI\mcp.json` (or equivalent) |
| `gemini-cli.json` | Gemini CLI | `%USERPROFILE%\.gemini\settings.json` |
| `codex-cli.toml` | OpenAI Codex CLI | `%USERPROFILE%\.codex\config.toml` |

## Usage

### Claude Desktop

Copy `claude-desktop.json` content into `%APPDATA%\Claude\claude_desktop_config.json` and restart Claude Desktop.

### Claude Code

Copy `claude-code.json` to `.claude/mcp.json` in your project root, or merge its `mcpServers` block into an existing `mcp.json`.

The config includes both a stdio entry (starts the server automatically) and an SSE entry (connects to a server you start manually).

### PC-AI Agent

Copy `pc-ai-agent.json` to the PC-AI config directory and reference it in the agent's MCP server list. All 23 tools are enumerated in the `tools` array for explicit capability declaration.

### Gemini CLI

Merge `gemini-cli.json` content into `%USERPROFILE%\.gemini\settings.json`. Create the file if it does not exist.

## Transport Modes

### stdio (default)

The server starts automatically as a subprocess. No manual server management required.

```bash
# Runs automatically via the config above
uv run --directory C:/codedev/windows-mcp windows-mcp
```

### SSE (HTTP streaming)

Start the server manually, then connect clients to the SSE URL.

```bash
# Without auth (localhost only)
uv run --directory C:/codedev/windows-mcp windows-mcp --transport sse --port 8765

# With auth
uv run --directory C:/codedev/windows-mcp windows-mcp --transport sse --port 8765 --api-key <key>
```

SSE endpoint: `http://localhost:8765/sse`

### Streamable HTTP

```bash
uv run --directory C:/codedev/windows-mcp windows-mcp --transport http --port 8765
```

HTTP endpoint: `http://localhost:8765/mcp`

## Auth Setup

Generate a persistent API key (stored encrypted via DPAPI):

```bash
cd C:/codedev/windows-mcp
uv run windows-mcp --generate-key
```

Rotate an existing key:

```bash
uv run windows-mcp --rotate-key
```

Pass the key at server start with `--api-key <key>`. Clients must send it as a Bearer token in the `Authorization` header. The server refuses to bind to non-localhost addresses without auth configured.

## Tool Reference (23 tools)

### Desktop Automation

| Tool | Purpose |
|------|---------|
| `App` | Launch, resize, maximize, minimize, close, or switch focus to a window |
| `Shell` | Execute PowerShell commands with captured output |
| `Snapshot` | Capture accessibility tree + optional screenshot of the active window |
| `Click` | Left/right/double-click at coordinates or on a named element |
| `Type` | Type text into the focused element; supports clear and Enter |
| `Scroll` | Scroll a window or element up/down/left/right |
| `Move` | Move the mouse to coordinates, with optional drag |
| `Shortcut` | Send keyboard shortcuts (e.g. `ctrl+c`, `alt+f4`) |
| `Wait` | Sleep for N seconds |
| `WaitFor` | Event-driven wait for an element matching criteria |
| `Scrape` | Fetch and extract text content from a URL |
| `MultiSelect` | Click multiple elements in sequence |
| `MultiEdit` | Apply edits to multiple coordinate targets |
| `Find` | Semantic element search by role, name, or UIA property |
| `Invoke` | Trigger UIA patterns: InvokePattern, ValuePattern, TogglePattern |

### System

| Tool | Purpose |
|------|---------|
| `Clipboard` | Get or set clipboard text content |
| `Process` | List running processes or kill by name/PID |
| `SystemInfo` | CPU, memory, disk, OS, and locale information |
| `Notification` | Send a Windows toast notification |
| `LockScreen` | Lock the Windows session |
| `Registry` | Read, write, or delete registry values (supports HKCU:\, HKLM:\ paths) |

### File

| Tool | Purpose |
|------|---------|
| `File` | Read, write, search files, or list directory trees |

### Vision

| Tool | Purpose |
|------|---------|
| `VisionAnalyze` | Analyze a screenshot with a vision model prompt |
