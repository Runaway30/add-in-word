# word-mcp — Project Context

## What this is
An MCP server that lets a **Microsoft 365 Copilot Studio custom agent** read and edit Word documents stored on **SharePoint / OneDrive**, applying changes as **Track Changes** so the user can accept/reject them in Word.

## Architecture

```
M365 Copilot Studio agent
        │  (MCP over streamable-http)
        ▼
FastMCP server  (mcp/src/word_mcp/server.py)
  ├── reads .docx via Microsoft Graph API (download only)
  ├── queues changes in memory (_pending dict)
  └── serves Word Add-in HTML at /addin/taskpane
        │  (polls GET /pending-changes every 3s)
        ▼
Word Add-in (Office.js)  (mcp/addin/taskpane.html)
  └── applies changes live via Word.run() with changeTrackingMode = trackAll
```

**Key design decision:** changes are NOT uploaded back to SharePoint via Graph API (that caused HTTP 423 Locked when the document was open). Instead they are queued server-side and applied live inside Word by the companion Office.js Add-in.

## Files
```
mcp/
  src/word_mcp/
    server.py        — FastMCP server, MCP tools, HTTP routes for Add-in
    graph.py         — Microsoft Graph API helpers (auth, download)
    track_changes.py — OOXML track-changes markup helpers (used for future OOXML work)
  addin/
    taskpane.html    — Word Add-in task pane (self-contained HTML+JS, served by server)
    manifest.xml     — Office Add-in manifest template ({{SERVER_URL}} replaced at runtime)
  pyproject.toml
  .env               — secrets (never commit)
```

## MCP Tools
| Tool | What it does |
|---|---|
| `get_current_document` | Returns the SharePoint URL registered by the open Add-in |
| `read_document(url)` | Downloads .docx via Graph API, returns paragraphs with index + text |
| `apply_tracked_changes(url, replacements)` | Queues text replacements (old_text → new_text) |
| `replace_paragraph(url, index, new_text)` | Queues a full paragraph replacement by index |
| `get_tracked_changes_summary(url)` | Reads w:ins/w:del from the saved file on SharePoint |

## HTTP Routes (for the Add-in)
| Route | Purpose |
|---|---|
| `POST /current-document` | Add-in registers the active document URL on load |
| `GET /current-document` | Returns currently registered document URL |
| `GET /pending-changes` | Add-in polls this to get queued changes |
| `POST /pending-changes/{id}/done` | Add-in confirms a change was applied |
| `GET /addin/taskpane` | Serves the task pane HTML |
| `GET /manifest.xml` | Serves manifest with {{SERVER_URL}} replaced |

## Environment variables (.env)
```
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
```
Azure AD app registration (`word-mcp-server`) with Graph API application permissions:
- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`

## Running locally
```sh
cd mcp
export $(cat .env | xargs)
uv run word-mcp --transport streamable-http --port 8000
# in another terminal:
cloudflared tunnel --url http://localhost:8000
```
The Cloudflare tunnel URL (e.g. `https://xxx.trycloudflare.com`) is used in Copilot Studio as the MCP endpoint: `https://xxx.trycloudflare.com/mcp`

## Word Add-in setup
1. Download the manifest: `https://<tunnel-url>/manifest.xml`
2. Sideload in Word (Mac): copy to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`, restart Word, then Developer → My Add-ins
3. In Word Online: Insert → Add-ins → Upload My Add-in (requires tenant admin to enable add-ins)

**Current blocker:** sideloading in Word for Mac desktop is not working. Word Online requires the M365 tenant admin to enable "User owned apps and services" in admin.microsoft.com → Settings → Org settings → Services.

## Copilot Studio setup
- Custom agent with MCP action pointing to `https://<tunnel-url>/mcp`
- Authentication: None
- The agent instructions in `mcp = FastMCP(instructions=...)` tell it to always call `get_current_document` first and never ask the user for the document URL

## Known issues / next steps
- **Add-in sideloading on Mac** is blocked by tenant policy — needs admin.microsoft.com fix
- **Cloudflare tunnel URL changes** on every restart — future fix is a permanent Azure deployment
- **In-memory queue** is lost on server restart — acceptable for now
- `track_changes.py` (OOXML manipulation) is no longer used in the main flow but kept for potential direct-file-edit mode
