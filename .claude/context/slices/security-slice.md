# Security Context Slice -- Windows-MCP

## Critical: No Auth on SSE/HTTP Transport
- `__main__.py:662-694` -- Zero authentication when using SSE or HTTP transport
- Any network client can invoke all 19 tools if host is non-localhost

## Critical: Unrestricted Shell Execution
- `desktop/service.py:209-237` -- Arbitrary PowerShell via subprocess.run -EncodedCommand
- `pg.FAILSAFE = False` disables abort mechanism
- Full env inherited (`os.environ.copy()`)

## High: File Path Traversal
- `filesystem/service.py` -- No path scoping. Absolute paths unrestricted.

## High: Registry Access Unrestricted
- `desktop/service.py:1022-1087` -- Can read/write/delete any registry key

## High: Hardcoded PostHog API Key
- `analytics.py:43` -- `phc_uxdCItyVTjXNU0sMPr97dq3tcz39scQNt3qjTYw5vLV`
- GeoIP enabled, exception auto-capture enabled, `**result` forwards unfiltered data

## High: Auth Client HTTP Default
- `auth/service.py:33` -- `http://localhost:3000` (plain HTTP)
- API key in JSON body, no TLS enforcement

## Medium: SSRF in Scrape
- `desktop/service.py:561-573` -- No URL validation, no private IP blocking

## Architectural Gaps
- No permission model, no audit logging, no rate limiting
- Error messages leak system info (paths, usernames)
