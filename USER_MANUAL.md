# Windows-MCP User Manual

**Version:** 0.6.2
**Platform:** Windows 7, 8, 8.1, 10, 11
**Python:** 3.13+
**License:** MIT

---

## Table of Contents

1. [What is Windows-MCP?](#1-what-is-windows-mcp)
2. [Installation](#2-installation)
3. [Configuration & Modes](#3-configuration--modes)
4. [Complete Tool Reference](#4-complete-tool-reference)
5. [Playwright Comparison: Browser vs Desktop Automation](#5-playwright-comparison-browser-vs-desktop-automation)
6. [Common Workflows](#6-common-workflows)
7. [Architecture Overview](#7-architecture-overview)
8. [Troubleshooting](#8-troubleshooting)
9. [Security Considerations](#9-security-considerations)
10. [Performance Characteristics](#10-performance-characteristics)
11. [Optimization Roadmap](#11-optimization-roadmap)

---

## 1. What is Windows-MCP?

Windows-MCP is an MCP (Model Context Protocol) server that bridges AI agents with the Windows operating system. It enables LLMs to perform desktop automation tasks -- clicking buttons, typing text, launching applications, reading UI state, managing files, and executing system commands -- all through a standardized tool interface.

**Key differentiator:** Unlike Playwright or Selenium which operate exclusively within web browsers, Windows-MCP operates at the **OS level** using the Windows Accessibility Tree (UIAutomation COM API). This means it can automate **any Windows application** -- native Win32, WPF, UWP, Electron, Qt, and web browsers alike.

---

## 2. Installation

### From PyPI (Recommended)

```bash
pip install uv  # if not installed
uvx windows-mcp
```

### From Source

```bash
git clone https://github.com/CursorTouch/Windows-MCP.git
cd Windows-MCP
uv sync
uv run windows-mcp
```

### Client Configuration

Add to your MCP client's config (Claude Desktop, Cursor, Gemini CLI, etc.):

```json
{
  "mcpServers": {
    "windows-mcp": {
      "command": "uvx",
      "args": ["windows-mcp"]
    }
  }
}
```

---

## 3. Configuration & Modes

### Local Mode (Default)

Runs directly on your Windows machine. The MCP client connects via stdio, SSE, or HTTP.

```bash
# stdio (default -- for Claude Desktop, Cursor, etc.)
uvx windows-mcp

# SSE (network-accessible)
uvx windows-mcp --transport sse --host localhost --port 8000

# Streamable HTTP (recommended for production)
uvx windows-mcp --transport streamable-http --host localhost --port 8000
```

### Remote Mode

Connects to a cloud-hosted Windows VM via windowsmcp.io.

```json
{
  "env": {
    "MODE": "remote",
    "SANDBOX_ID": "your-sandbox-id",
    "API_KEY": "your-api-key"
  }
}
```

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `ANONYMIZED_TELEMETRY` | `true` | Set to `false` to disable PostHog telemetry |
| `MODE` | `local` | `local` or `remote` |
| `SANDBOX_ID` | (none) | Required for remote mode |
| `API_KEY` | (none) | Required for remote mode |

---

## 4. Complete Tool Reference

### State Inspection

#### Snapshot
Captures the complete desktop state. **Always call this first** to understand what's on screen before taking actions.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `use_vision` | bool | `false` | Include screenshot (PNG image) |
| `use_dom` | bool | `false` | Browser-focused mode: extracts web page DOM elements instead of browser UI chrome |

**Returns:**
- System language
- Active desktop and all virtual desktops
- Focused window details
- All open windows
- Interactive elements (buttons, text fields, links, menus) with coordinates
- Scrollable areas with scroll percentages

**Example output (interactive elements):**
```
# id|window|control_type|name|coords|focus
0|Notepad|ButtonControl|Close|(1893,8)|False
1|Notepad|EditControl|Text Editor|(960,540)|True
2|Notepad|MenuBarControl|Application|(400,30)|False
```

#### SystemInfo
Returns CPU, memory, disk, network stats, and uptime.

### Input Simulation

#### Click
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `loc` | [x, y] | required | Screen coordinates |
| `button` | "left"/"right"/"middle" | "left" | Mouse button |
| `clicks` | 0/1/2 | 1 | 0=hover, 1=single, 2=double |

#### Type
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `loc` | [x, y] | required | Where to type |
| `text` | string | required | Text to type |
| `clear` | bool | false | Clear existing text first |
| `press_enter` | bool | false | Press Enter after typing |
| `caret_position` | "start"/"idle"/"end" | "idle" | Move caret before typing |

#### Scroll
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `loc` | [x, y] | null | Scroll location (null = current mouse) |
| `type` | "vertical"/"horizontal" | "vertical" | Scroll axis |
| `direction` | "up"/"down"/"left"/"right" | "down" | Scroll direction |
| `wheel_times` | int | 1 | Scroll amount (1 wheel ~ 3-5 lines) |

#### Move
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `loc` | [x, y] | required | Target coordinates |
| `drag` | bool | false | Drag from current position to target |

#### Shortcut
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `shortcut` | string | required | Key combo, e.g. "ctrl+c", "alt+tab", "win+r" |

#### MultiSelect
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `locs` | [[x,y], ...] | required | List of coordinates to click |
| `press_ctrl` | bool | true | Hold Ctrl while clicking (multi-select) |

#### MultiEdit
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `locs` | [[x,y,text], ...] | required | Coordinate-text pairs to fill multiple fields |

### Application Management

#### App
| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | "launch"/"resize"/"switch" | Operation mode |
| `name` | string | App name (for launch/switch) |
| `window_loc` | [x, y] | Window position (resize mode) |
| `window_size` | [w, h] | Window size (resize mode) |

#### Wait
| Parameter | Type | Description |
|-----------|------|-------------|
| `duration` | int | Seconds to pause |

### File Operations

#### File
| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | "read"/"write"/"copy"/"move"/"delete"/"list"/"search"/"info" | Operation |
| `path` | string | File/directory path (relative = Desktop) |
| `destination` | string | For copy/move |
| `content` | string | For write |
| `pattern` | string | Glob pattern for list/search |
| `recursive` | bool | Recursive operation |
| `append` | bool | Append to file (write mode) |
| `overwrite` | bool | Overwrite existing (copy/move) |
| `offset`/`limit` | int | Line range for read |
| `encoding` | string | File encoding (default: utf-8) |
| `show_hidden` | bool | Show hidden files in list |

### System Operations

#### Shell
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `command` | string | required | PowerShell command to execute |
| `timeout` | int | 30 | Timeout in seconds |

#### Process
| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | "list"/"kill" | Operation |
| `name` | string | Process name filter |
| `pid` | int | Process ID (kill mode) |
| `sort_by` | "memory"/"cpu"/"name" | Sort order (list mode) |
| `limit` | int | Max results (list mode, default 20) |
| `force` | bool | Force kill (SIGKILL vs SIGTERM) |

#### Registry
| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | "get"/"set"/"delete"/"list" | Operation |
| `path` | string | Registry path (PowerShell format, e.g. "HKCU:\Software\MyApp") |
| `name` | string | Value name |
| `value` | string | Value data (set mode) |
| `type` | "String"/"DWord"/"QWord"/"Binary"/"MultiString"/"ExpandString" | Value type |

#### Clipboard
| Parameter | Type | Description |
|-----------|------|-------------|
| `mode` | "get"/"set" | Read or write clipboard |
| `text` | string | Text to set (set mode) |

#### Notification
| Parameter | Type | Description |
|-----------|------|-------------|
| `title` | string | Toast notification title |
| `message` | string | Toast notification body |

#### LockScreen
Locks the Windows workstation. No parameters.

#### Scrape
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `url` | string | required | URL to scrape |
| `use_dom` | bool | false | Extract from active browser tab instead of HTTP request |

---

## 5. Playwright Comparison: Browser vs Desktop Automation

This section answers: **What Playwright-like capabilities does Windows-MCP have, and do they extend to desktop applications or only web browsers?**

### TL;DR

Windows-MCP is to **the entire Windows desktop** what Playwright is to **web browsers**. It operates at the OS accessibility layer, meaning it can automate any application -- including desktop apps, which Playwright fundamentally cannot. However, it trades Playwright's deep browser internals (DOM access, network interception, JS execution) for breadth of coverage across all applications.

### Detailed Capability Comparison

| Capability | Playwright | Windows-MCP | Scope |
|-----------|-----------|-------------|-------|
| **Element Discovery** | DOM selectors (CSS, XPath, text, role) | Windows Accessibility Tree (UIAutomation) | Playwright: browsers only. **Windows-MCP: ALL applications** |
| **Click Elements** | `page.click(selector)` | `Click(loc=[x,y])` | Both. Windows-MCP uses coordinates, not selectors |
| **Type Text** | `page.fill(selector, text)` | `Type(loc=[x,y], text="...")` | Both. Windows-MCP types character-by-character |
| **Keyboard Shortcuts** | `page.keyboard.press("Control+c")` | `Shortcut(shortcut="ctrl+c")` | Both |
| **Screenshots** | `page.screenshot()` | `Snapshot(use_vision=True)` | Both. Windows-MCP captures the entire desktop |
| **Wait for State** | `page.waitForSelector()`, `page.waitForLoadState()` | `Wait(duration=N)` | Playwright has intelligent waits; Windows-MCP uses fixed delays |
| **Scroll** | `page.mouse.wheel()` | `Scroll(loc, direction, wheel_times)` | Both |
| **Drag & Drop** | `page.dragAndDrop(src, dst)` | `Move(loc, drag=True)` | Both |
| **Multi-select** | Manual Ctrl+Click | `MultiSelect(locs, press_ctrl=True)` | Both |
| **Form Filling** | `page.fill()` per field | `MultiEdit(locs=[[x,y,"text"],...])` | Both |
| **DOM Access** | Full DOM tree, `page.content()`, `page.evaluate()` | `Snapshot(use_dom=True)` -- limited to accessibility-exposed content | Playwright: full DOM. **Windows-MCP: a11y subset only** |
| **Network Interception** | `page.route()`, request/response hooks | Not available | Playwright only |
| **JavaScript Execution** | `page.evaluate("js code")` | Not available | Playwright only |
| **Browser Contexts** | Isolated contexts, cookies, storage | Not available (relies on actual browser state) | Playwright only |
| **PDF Generation** | `page.pdf()` | Not available | Playwright only |
| **Video Recording** | Built-in trace recording | Not available | Playwright only |
| **App Launch** | Browser launch only | `App(mode="launch", name="...")` -- **any** application | **Windows-MCP only** |
| **Window Management** | Browser window only | `App(mode="resize/switch")` -- **any** window | **Windows-MCP only** |
| **File Operations** | Not available | `File(mode="read/write/copy/move/delete/...")` | **Windows-MCP only** |
| **Shell Commands** | Not available | `Shell(command="...")` | **Windows-MCP only** |
| **Process Management** | Not available | `Process(mode="list/kill")` | **Windows-MCP only** |
| **Registry Access** | Not available | `Registry(mode="get/set/delete/list")` | **Windows-MCP only** |
| **Clipboard** | Not directly (via JS eval) | `Clipboard(mode="get/set")` | **Windows-MCP only** |
| **System Notifications** | Not available | `Notification(title, message)` | **Windows-MCP only** |
| **Virtual Desktops** | Not available | Tracked in Snapshot output | **Windows-MCP only** |

### How Element Discovery Differs

**Playwright** navigates the browser DOM:
```javascript
await page.click('button.submit');          // CSS selector
await page.click('//button[@id="save"]');   // XPath
await page.getByRole('button', { name: 'Save' });  // Accessibility role
await page.getByText('Click me');           // Text content
```

**Windows-MCP** reads the Windows Accessibility Tree:
```
1. Call Snapshot() to get interactive elements with coordinates
2. Response includes: id|window|control_type|name|coords|focus
3. Use the coordinates to Click/Type at those positions
```

The critical difference: Playwright selectors work within a single page's DOM. Windows-MCP's Accessibility Tree spans **all visible applications on the desktop simultaneously**. One `Snapshot` call returns buttons, text fields, and menus from Notepad, Chrome, File Explorer, and any other open application.

### Browser-Specific Mode: `use_dom=True`

When interacting with web browsers (Chrome, Edge, Firefox), Windows-MCP has a special mode that filters the Accessibility Tree to show only web page content, excluding browser UI chrome:

```
Snapshot(use_dom=True)  --> Returns web page elements only
Scrape(url, use_dom=True) --> Extracts visible text from active tab's DOM
```

This provides a subset of Playwright's DOM access -- you get the text content and interactive elements exposed through the browser's accessibility API, but NOT:
- Raw HTML/CSS
- Shadow DOM
- JavaScript execution
- Network requests
- Cookie/storage access
- Invisible/off-screen elements

**Browser detection is automatic.** Windows-MCP identifies Chrome (`chrome.exe`), Edge (`msedge.exe`), and Firefox (`firefox.exe`) by process name and adjusts behavior accordingly.

### Desktop Applications: What Windows-MCP Can Do That Playwright Cannot

Windows-MCP automates desktop applications through UIAutomation, which exposes:

| UI Element | Examples | Automatable? |
|-----------|---------|--------------|
| Buttons | Save, Cancel, OK dialogs | Yes -- Click at coordinates |
| Text fields | Input boxes in any app | Yes -- Type with clear/append |
| Menus | File, Edit, View menus | Yes -- Click to open, then click items |
| Toolbars | Ribbon controls, icon bars | Yes -- Click at coordinates |
| Trees | File Explorer tree, settings panels | Yes -- Click to expand/select |
| Lists | File lists, dropdown options | Yes -- Click/MultiSelect |
| Tabs | Settings tabs, browser tabs | Yes -- Click at coordinates |
| Scroll areas | Long documents, data grids | Yes -- Scroll tool |
| Context menus | Right-click menus | Yes -- Click(button="right") + Click |
| System tray | Notification area icons | Yes -- Click at coordinates |
| UAC prompts | Elevation dialogs | Partially -- depends on permissions |

**Applications confirmed working:** Notepad, File Explorer, Settings, Calculator, VS Code (Electron), Chrome/Edge/Firefox, Office suite, Task Manager, and any WPF/WinForms/Qt application that exposes UIAutomation.

### Limitations vs Playwright

| Limitation | Impact |
|-----------|--------|
| **Coordinate-based interaction** | No semantic selectors. Must call `Snapshot` first to get element positions. If the window moves, coordinates change. |
| **No element waiting** | Only `Wait(seconds)` available. No "wait for element to appear" like Playwright's `waitForSelector`. |
| **Character-by-character typing** | `Type` uses `pyautogui.typewrite()` -- not suitable for fast IDE coding. Playwright's `fill()` sets values instantly. |
| **No headless mode** | Requires a visible desktop (or RDP/VNC session). Playwright can run headless. |
| **No parallel browser contexts** | Cannot run isolated browser sessions. Playwright excels at parallel testing. |
| **Text selection limitations** | Cannot select specific text within a paragraph (a11y tree limitation). |
| **No network layer** | Cannot intercept/mock API calls. Playwright's network interception is a major testing feature. |
| **Latency** | 0.2-0.9s per action (mouse click to next click). Playwright operates in milliseconds. |

### When to Use Each

| Use Case | Best Tool |
|----------|-----------|
| Web testing/scraping with deep DOM access | Playwright |
| Cross-browser testing automation | Playwright |
| API mocking/network interception | Playwright |
| Desktop application automation | **Windows-MCP** |
| Multi-application workflows (e.g., copy from Excel to web form) | **Windows-MCP** |
| System administration tasks | **Windows-MCP** |
| QA testing native Windows apps | **Windows-MCP** |
| AI agent desktop interaction | **Windows-MCP** |
| Browser + desktop app coordination | **Windows-MCP** (covers both) |

### Extended Framework Comparison

Beyond Playwright, Windows-MCP occupies a unique position in the Windows automation ecosystem:

| Framework | Type | Native Apps? | MCP Interface? | AI Agent Native? | Status |
|-----------|------|-------------|----------------|-----------------|--------|
| **Windows-MCP** | MCP Server | Yes (UIAutomation) | Yes | Yes | Active |
| **Playwright** | Browser automation | No (browser only) | Via playwright-mcp | Via MCP | Active |
| **FlaUI** | .NET testing library | Yes (UIAutomation) | No | No | Active |
| **Power Automate Desktop** | Enterprise RPA | Yes (recorder) | Dataverse MCP (2025) | Via Copilot Studio | Active |
| **WinAppDriver** | WebDriver for Windows | Yes (UIAutomation) | No | No | **Abandoned** (2021) |
| **AutoHotkey** | Scripting tool | Yes (Win32 messages) | No | No | Active |

**Key insight:** Windows-MCP is the only tool that combines UIAutomation-level desktop access with a native MCP interface for AI agent consumption. FlaUI has stronger element interaction patterns (pattern-based invocation vs coordinate clicking), and Playwright has deeper browser capabilities, but neither can be directly driven by an LLM in a conversational session.

**Complementarity with Playwright:** The two tools are almost perfectly complementary. In a future version, Windows-MCP could detect browser windows and delegate to a running `playwright-mcp` instance for DOM tasks, while retaining control for native app interactions. This would give an AI agent the best tool for each context automatically.

---

## 6. Common Workflows

### Launch App and Interact

```
1. App(mode="launch", name="Notepad")
2. Wait(duration=2)
3. Snapshot()                          --> See all interactive elements
4. Type(loc=[500,300], text="Hello!")  --> Type into editor
5. Shortcut(shortcut="ctrl+s")        --> Save
```

### Fill a Web Form

```
1. App(mode="launch", name="Chrome")
2. Wait(duration=3)
3. Snapshot(use_dom=True)              --> See form fields with coordinates
4. MultiEdit(locs=[
     [300, 200, "John"],              --> First name
     [300, 250, "Doe"],               --> Last name
     [300, 300, "john@example.com"]   --> Email
   ])
5. Click(loc=[300, 400])              --> Submit button
```

### File Management

```
1. File(mode="list", path="C:\\Users\\Me\\Documents", pattern="*.pdf")
2. File(mode="copy", path="report.pdf", destination="C:\\Backup\\report.pdf")
3. File(mode="info", path="C:\\Backup\\report.pdf")
```

### System Monitoring

```
1. SystemInfo()                        --> CPU, memory, disk, network
2. Process(mode="list", sort_by="cpu", limit=10)
3. Process(mode="kill", name="hung_app.exe", force=True)
```

### Web Content Extraction

```
# Method 1: Direct HTTP (works if site allows)
Scrape(url="https://example.com")

# Method 2: Via browser DOM (for JS-rendered pages)
1. App(mode="launch", name="Chrome")
2. Click on address bar, Type URL
3. Wait(duration=3)
4. Scrape(url="https://example.com", use_dom=True)
```

---

## 7. Architecture Overview

```
MCP Client (Claude Desktop, Cursor, etc.)
    |
    | MCP Protocol (stdio / SSE / HTTP)
    |
FastMCP Server (__main__.py)
    |
    +-- Desktop Service (desktop/service.py)
    |     +-- Input simulation (pyautogui)
    |     +-- Window management (win32gui)
    |     +-- Process management (psutil)
    |     +-- Registry (PowerShell subprocess)
    |     +-- Web scraping (requests + markdownify)
    |
    +-- Tree Service (tree/service.py)
    |     +-- Accessibility tree traversal
    |     +-- Interactive element discovery
    |     +-- DOM mode for browsers
    |     +-- Element caching (cache_utils.py)
    |
    +-- UIAutomation Wrapper (uia/)
    |     +-- COM API abstraction (comtypes)
    |     +-- Control type handling
    |     +-- Pattern support
    |
    +-- WatchDog Service (watchdog/service.py)
    |     +-- Focus change monitoring
    |     +-- STA thread for COM events
    |
    +-- Virtual Desktop Manager (vdm/core.py)
    |     +-- Win10/11 virtual desktop tracking
    |
    +-- Filesystem Service (filesystem/service.py)
    |     +-- File CRUD operations
    |
    +-- Auth Client (auth/service.py)
    |     +-- Remote mode authentication
    |
    +-- Analytics (analytics.py)
          +-- PostHog telemetry (opt-out via env var)
```

### How the Accessibility Tree Works

Windows-MCP reads the **Windows UIAutomation tree** -- a structured representation of all UI elements exposed by applications. This is the same API used by screen readers (Narrator, JAWS, NVDA).

For each visible window, the Tree Service:
1. Gets the window handle from Win32
2. Creates a UIAutomation element from the handle
3. Recursively traverses all child elements
4. Classifies each as interactive (clickable/typeable), scrollable, or informative
5. Returns coordinates, names, control types, and values

For browsers with `use_dom=True`, it additionally:
1. Detects the browser process (Chrome/Edge/Firefox)
2. Filters the tree to the document content area
3. Excludes browser chrome (tabs, address bar, toolbar)
4. Returns only web page elements

---

## 8. Troubleshooting

### First Run Timeout

The first run installs dependencies and may timeout. Restart the MCP server.

### "No interactive elements found"

- Ensure the target application is in the foreground
- Some apps have poor UIAutomation support (e.g., some games, custom-rendered UIs)
- Try `Snapshot(use_vision=True)` to see a screenshot of what's visible

### Elements at Wrong Coordinates

- Check screen resolution/DPI settings
- Windows-MCP caps screenshots at 1920x1080 and scales coordinates accordingly
- Multi-monitor setups may shift coordinates

### Shell Commands Timeout

- Default timeout is 30 seconds
- Increase with `Shell(command="...", timeout=120)`

### Non-English Windows

- The `App` tool may not work correctly with non-English Start Menu
- Workaround: disable the `App` tool and use `Shell` to launch applications

---

## 9. Security Considerations

**Windows-MCP operates with full system access.** There is no sandboxing or permission model.

### Recommendations

1. **Run in a VM or Windows Sandbox** for untrusted agents
2. **Use stdio transport** (default) rather than SSE/HTTP to limit exposure
3. **Never bind to 0.0.0.0** unless you've added your own authentication layer
4. **Disable telemetry** if operating in sensitive environments: `ANONYMIZED_TELEMETRY=false`
5. **Review SECURITY.md** for comprehensive security guidelines

### What the AI Agent Can Do

With Windows-MCP, a connected AI agent has the ability to:
- Execute any PowerShell command
- Read, write, and delete any file
- Modify the Windows Registry
- Kill any process
- Capture screenshots
- Read clipboard contents
- Lock the workstation

**Treat this server as giving the AI agent the same access level as your user account.**

---

## 10. Performance Characteristics

Understanding latency is important for designing effective automation workflows.

### Current Performance Profile

| Operation | Typical Latency | Bottleneck |
|-----------|----------------|------------|
| `Snapshot()` (no vision) | 0.5-5s | UIAutomation tree traversal (COM round-trips) |
| `Snapshot(use_vision=True)` | 0.6-5.4s | Tree traversal + screenshot capture + PNG encode |
| `Click(loc)` | ~1.1s | `pyautogui.PAUSE = 1.0` (configurable) |
| `Type(loc, text, clear=True)` | 4-7s | Multiple pyautogui calls, each with 1s pause |
| `App(mode="launch")` | 0.5-2s | PowerShell subprocess for Start Menu lookup |
| `Shell(command)` | 0.2-0.5s + exec | PowerShell process initialization overhead |
| `File(mode="read")` | <10ms | Direct filesystem I/O |

### Why Latency Varies

The tree traversal cost depends on desktop complexity:
- **Simple desktop** (1-2 windows, <100 elements): ~500ms
- **Typical desktop** (3-5 windows, 500-2000 elements): ~1-3s
- **Complex desktop** (browser with large DOM, many windows): ~3-5s+

Each UI element requires cross-process COM calls to the target application's UIAutomation provider. The current implementation makes per-node calls that compound across hundreds of elements.

### Tips for Faster Automation

1. **Minimize Snapshot calls** -- Cache element coordinates when the window hasn't changed
2. **Use `use_dom=True` only when needed** -- Browser DOM traversal is the slowest path
3. **Prefer `Shortcut` over `Click`+`Type` chains** -- Keyboard shortcuts bypass coordinate lookup
4. **Close unnecessary windows** -- Fewer visible windows = faster tree traversal

---

## 11. Optimization Roadmap

The following optimizations are planned (see TODO.md for full details):

**Near-term (Python-only):**
- Event-driven `WaitFor` tool replacing fixed-sleep `Wait`
- `ValuePattern.SetValue()` for instant text input (vs character-by-character)
- Reduced `pyautogui.PAUSE` from 1.0s to 0.05s
- Single-shot `TreeScope_Subtree` cache (60-80% tree traversal reduction)
- Elimination of PowerShell subprocess for registry/sysinfo operations

**Medium-term (Caching):**
- Per-window tree cache with WatchDog-based invalidation
- Start Menu app cache with filesystem watcher
- Element coordinate cache reusing stored bounding boxes

**Longer-term (Rust acceleration via PyO3):**
- Rust UIAutomation tree traversal (windows-rs + rayon): 50-200ms target
- DXGI screenshot capture: 1-3ms target
- Playwright-MCP bridge for browser window delegation
- Capability/permission manifest for enterprise deployments
