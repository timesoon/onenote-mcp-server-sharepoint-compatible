# OneNote MCP Server — extended fork

A Model Context Protocol server for Microsoft OneNote integration with Claude Desktop. This is an actively-maintained fork that adds pagination, nested folder (section group) support, OneDrive for Business / SharePoint compatibility, page-link surfacing, and several bug fixes on top of the upstream base.

**Use this fork if you need any of:**

- Notebooks with more than ~100 pages in a section (upstream silently truncates)
- SharePoint-backed notebooks (OneDrive for Business accounts)
- Nested section group / folder hierarchies
- Direct URLs (`web_url`, `client_url`) to pages, sections, and notebooks in tool output
- Working `get_page_resources` (upstream throws `'FunctionTool' object is not callable`)

## Lineage

```
purpleslurple/onenote-mcp-server   <-- original base
  └─ peterstahley/onenote-mcp-server-sharepoint-compatible   <-- adds SharePoint auth
      └─ this fork   <-- adds pagination, section groups, URL surfacing, bug fixes
```

Both upstream maintainers have open PRs from this fork ([upstream pagination fix](https://github.com/purpleslurple/onenote-mcp-server/pulls), [feature PR to peter](https://github.com/peterstahley/onenote-mcp-server-sharepoint-compatible/pulls)). If those merge you can switch back to upstream; until then, use this.

## What's new vs upstream

### Bug fixes
- Pagination on every collection endpoint via new `make_graph_request_all` helper. Handles OneNote's known `@odata.nextLink` omission with a manual `$skip` fallback.
- `get_page_resources` no longer throws `FunctionTool object is not callable`.
- `list_pages` `personal_err` Python 3 scoping bug fixed.
- `python-dotenv` import is now optional.

### New tools
- `list_section_groups(notebook_id)` — top-level folders in a notebook
- `list_sections_in_group(group_id)` — sections directly inside a folder
- `list_section_group_contents(group_id)` — sections + nested folders inside a folder
- `list_all_sections(notebook_id, max_depth=10)` — flat list of every section across the entire notebook with breadcrumb `group_name` tags
- `enumerate_notebook(notebook_id, include_pages=False, max_depth=10)` — recursive tree of the full notebook
- `debug_list_pages(...)` — diagnostic tool for investigating pagination issues
- `restart_server()` — reload server code without restarting Claude Desktop

### URL surfacing
Every notebook / section / page returned by `list_*` and `enumerate_notebook` now includes:
- `web_url` — opens the item in OneNote on the web (SharePoint URL with `?wd=target(...)` deep-link)
- `client_url` — `onenote:` scheme URL that deep-links into the OneNote desktop app

## Quick start

### Prerequisites
- Python 3.10+
- [uv](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or pip
- Claude Desktop
- Microsoft Azure account (free)

### 1. Clone

```bash
git clone https://github.com/timesoon/onenote-mcp-server-sharepoint-compatible.git
cd onenote-mcp-server-sharepoint-compatible
uv sync
```

### 2. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. **Azure Active Directory** → **App registrations** → **New registration**
3. Name: anything; account type: "Accounts in any organizational directory and personal Microsoft accounts"; redirect URI: select "Public client/native" and enter `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Register, copy the **Application (client) ID**
5. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → add `Notes.Read`, `Notes.ReadWrite`, `Notes.Read.All`, `Notes.ReadWrite.All`, `User.Read`
6. Click **Grant admin consent**

### 3. Configure Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json` (Windows) or `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory", "/FULL/PATH/TO/onenote-mcp-server-sharepoint-compatible",
        "run", "python", "onenote_mcp_server.py"
      ],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id-here"
      }
    }
  }
}
```

Restart Claude Desktop. You should see OneNote tools in the 🔨 menu.

### 4. Authenticate

In Claude Desktop, say `Start OneNote authentication`. Follow the device-code flow. Then `Complete OneNote authentication`.

## SharePoint / OneDrive for Business

If your OneNote notebooks live on SharePoint (typical for OneDrive for Business / Microsoft 365 work accounts), set these env vars in addition to `AZURE_CLIENT_ID`:

```json
"env": {
    "AZURE_CLIENT_ID": "your-client-id-here",
    "AZURE_TENANT_ID": "your-tenant-id-here",
    "SHAREPOINT_HOST": "yourcompany-my.sharepoint.com",
    "SHAREPOINT_USER_PATH": "/personal/your_user_yourcompany_com"
}
```

Without these, the server falls back to personal-only OneNote — which works for `live.com` / personal Microsoft accounts but returns 404s for work accounts.

## Environment variables

| Variable | Required | Purpose |
|----------|----------|---------|
| `AZURE_CLIENT_ID` | Yes | Azure App Registration client ID |
| `AZURE_TENANT_ID` | No (default: `common`) | Set to your tenant ID for SharePoint |
| `SHAREPOINT_HOST` | No | SharePoint host, e.g. `contoso-my.sharepoint.com` |
| `SHAREPOINT_USER_PATH` | No | Path to your personal site, e.g. `/personal/alice_contoso_com` |
| `ONENOTE_CACHE_TOKENS` | No (default: `true`) | Set to `false` to disable on-disk token caching |

## Recursion depth

`enumerate_notebook` and `list_all_sections` take a `max_depth` parameter (default `10`). OneNote's section groups form a tree (no cycles), so this is a soft safety valve. If hit, a warning is logged and deeper branches are omitted — no error. Real notebooks rarely go deeper than 3–4 levels.

## Troubleshooting

**"No tools available" in Claude Desktop** — confirm you restarted Claude Desktop after editing the config, and that the path in the config is absolute (no `~`). Run `uv --version` to confirm uv is installed.

**Auth fails** — Firefox / Chrome / Edge work; Safari has known issues with Microsoft's OAuth redirect. Use `Clear OneNote token cache` if you need to reset.

**Empty results from work accounts** — set `SHAREPOINT_HOST` and `SHAREPOINT_USER_PATH` (see above).

**Pagination still truncating** — call `debug_list_pages` on the offending section with `count=true` to see actual record counts and whether `@odata.nextLink` was returned. Send the output if filing an issue.

## License

MIT (inherited from upstream). See LICENSE.

## Acknowledgments

- [purpleslurple](https://github.com/purpleslurple) for the original implementation
- [peterstahley](https://github.com/peterstahley) for the SharePoint auth work
- Built on [FastMCP](https://github.com/jlowin/fastmcp) and the Microsoft Graph API
