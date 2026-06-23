# OneNote MCP Server

A complete, robust Model Context Protocol (MCP) server for Microsoft OneNote integration with Claude Desktop. Access your entire OneNote knowledge base through natural language queries.

## 🎯 What This Does

Transform your OneNote notebooks into an AI-accessible knowledge base:
- **List all your notebooks, sections, and pages**
- **Read page content** for analysis and search
- **Natural language queries** like "Show me my DevOps notes" or "Find pages about project planning"
- **Secure OAuth authentication** with Microsoft Graph API
- **Bulletproof error handling** with detailed debugging

## ✨ Why This Implementation

Unlike other OneNote MCP servers, this one:
- ✅ **Actually works** - tested extensively with real OneNote data
- ✅ **Complete functionality** - all core OneNote operations implemented
- ✅ **Robust authentication** - two-step device flow that handles edge cases
- ✅ **Production ready** - proper error handling and logging
- ✅ **Easy setup** - detailed instructions for non-technical users

## 🚀 Quick Start

### Prerequisites
- Python 3.10+ 
- [uv package manager](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or pip
- Claude Desktop
- Microsoft Azure account (free)

### 1. Install uv (if you don't have it)
```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# or with Homebrew
brew install uv
```

### 2. Clone and Setup
```bash
git clone https://github.com/yourusername/onenote-mcp-server.git
cd onenote-mcp-server

# Create virtual environment and install dependencies
uv sync
```

### 3. Azure App Registration
You need to create an Azure app to access OneNote. **Don't worry, it's free and takes 5 minutes:**

1. Go to [Azure Portal](https://portal.azure.com) (sign in with your Microsoft account)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Fill out the form:
   - **Name**: "OneNote MCP Server" (or whatever you like)
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Select "Public client/native" and enter: `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Register**
5. Copy the **Application (client) ID** - you'll need this!

### 4. Add Permissions
Still in your Azure app:
1. Go to **API permissions** → **Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Notes.Read` - Read OneNote notebooks
   - `Notes.ReadWrite` - Create/modify OneNote content (optional but recommended)
   - `User.Read` - Read user profile
4. Click **Grant admin consent** (the button at the top)

### 5. Configure Claude Desktop
Edit your Claude Desktop config file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\\Claude\\claude_desktop_config.json`

Add this configuration (replace `/ABSOLUTE/PATH/TO/PARENT/FOLDER/weather` with your actual path):

**Basic configuration:**
```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory", "/FULL/PATH/TO/onenote-mcp-server",
        "run", "python", "onenote_mcp_server.py"
      ],
      "env": {
        "AZURE_CLIENT_ID": "your-azure-client-id-here"
      }
    }
  }
}
```

**With explicit token caching control:**
```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory", "/FULL/PATH/TO/onenote-mcp-server",
        "run", "python", "onenote_mcp_server.py"
      ],
      "env": {
        "AZURE_CLIENT_ID": "your-azure-client-id-here",
        "ONENOTE_CACHE_TOKENS": "true"
      }
    }
  }
}
```

Replace `/FULL/PATH/TO/onenote-mcp-server` with the actual path to this project.

### 6. Restart Claude Desktop
Completely quit and restart Claude Desktop. You should see OneNote tools in the 🔨 menu.

## 🔐 First Time Authentication

1. In Claude Desktop, say: **"Start OneNote authentication"**
2. Claude will give you a URL and code
3. Visit the URL in your browser, enter the code, and sign in
4. **Browser compatibility**: 
   - ✅ **Firefox** (tested with 139.0.4) - works perfectly
   - ❌ **Safari** - may have issues with Microsoft OAuth redirect
   - ✅ **Chrome/Edge** - should work (Microsoft's browsers)
5. Come back to Claude and say: **"Complete OneNote authentication"**
6. You're ready to go!

### Token Persistence

By default, authentication tokens are cached securely on your local machine so you only need to authenticate once every few weeks/months. 

**To disable token caching** (for security-sensitive environments):
```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory", "/FULL/PATH/TO/onenote-mcp-server",
        "run", "python", "onenote_mcp_server.py"
      ],
      "env": {
        "AZURE_CLIENT_ID": "YOUR_CLIENT_ID_HERE",
        "ONENOTE_CACHE_TOKENS": "false"
      }
    }
  }
}
```

**Token caching options:**
- `ONENOTE_CACHE_TOKENS=true` (default) - Tokens persist across sessions
- `ONENOTE_CACHE_TOKENS=false` - Authenticate every session (more secure)

## 📖 Usage Examples

Once authenticated, try these commands in Claude Desktop:

```
List my OneNote notebooks
Show me sections in my Work notebook  
What pages are in my Ideas section?
Read the content of my "Project Plan" page
```

## 🛠 Troubleshooting

### "No tools available" in Claude Desktop
- Make sure you restarted Claude Desktop after config changes
- Check that the path in your config is correct (use full absolute path)
- Verify uv is installed: `uv --version`

### Authentication issues
- **Safari OAuth problems**: Safari may not handle Microsoft's OAuth redirect properly - use Firefox or Chrome instead
- **"nativeclient" prompts**: Normal Microsoft OAuth behavior, but if it blocks authentication, try a different browser
- **Authentication expired**: Use "Check OneNote authentication status" to see token expiry
- **Clear cached tokens**: Use "Clear OneNote token cache" if you need to reset authentication
- **Recommended browsers**: Firefox (confirmed working), Chrome, or Edge for best compatibility

### "Command not found" errors
- Make sure uv is in your PATH
- Alternative: replace `"uv"` with `"python"` in the config and use the full path to your Python interpreter

### Permission denied errors
- Check the file permissions in your project directory
- Make sure Claude Desktop can read the files

## 🏗 Development

### Project Structure
```
onenote-mcp-server/
├── onenote_mcp_server.py      # Main server implementation
├── pyproject.toml             # Dependencies and metadata  
├── README.md                  # This file
├── LICENSE                    # MIT License
└── .gitignore                 # Git ignore rules
```

### Key Features
- **Two-step authentication**: Handles device code flow properly
- **Complete Graph API integration**: All OneNote operations supported
- **Robust error handling**: Detailed logging and graceful failures
- **FastMCP framework**: Clean, maintainable code structure
- **Environment variable configuration**: Secure credential handling

### Adding New Features
The server is built with FastMCP, making it easy to add new tools:

```python
@mcp.tool()
async def your_new_tool(param: str) -> str:
    """Description of what your tool does."""
    # Your implementation here
    return result
```

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repo
2. Create a feature branch
3. Add tests for new functionality  
4. Submit a pull request

## 📄 License

MIT License - see LICENSE file for details.

## 🙏 Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp) framework
- Uses Microsoft Graph API for OneNote access
- Inspired by the amazing potential of AI + personal knowledge bases

## ⚠️ Important Notes

- This server only reads/writes data you already have access to
- Your Azure app credentials stay on your machine
- All authentication happens directly between you and Microsoft
- No data is sent to third parties

---

**Built with ❤️ for the Claude + OneNote community**

*Turn your OneNote into an AI-accessible knowledge base!*


---

# README additions

These sections should be added/merged into peter's existing `README.md`. They do not replace it — the existing setup instructions, Azure app registration steps, and licence still apply.

---

## What this release adds

### Full pagination

Every OneNote collection endpoint is now paginated. This fixes silent truncation at ~100 records that affected `list_pages` in particular (OneNote's API does not always emit `@odata.nextLink` for the pages endpoint). Internally this is handled by `make_graph_request_all` which uses `@odata.nextLink` when available and falls back to manual `$skip` pagination when it isn't.

### Section group (folder) support

OneNote notebooks can contain nested folder hierarchies via section groups. Four new tools expose this structure:

| Tool | Description |
|------|-------------|
| `list_section_groups(notebook_id)` | Top-level folders in a notebook |
| `list_sections_in_group(group_id)` | Sections directly inside a folder |
| `list_section_group_contents(group_id)` | Sections + nested subfolders inside a folder (one level) |
| `enumerate_notebook(notebook_id, include_pages, max_depth)` | Full recursive tree of the notebook |
| `list_all_sections(notebook_id, max_depth)` | Flat list of every section with breadcrumb `group_name` tags |

`list_sections` returns root-level sections only, matching the upstream base and composing with the tools above.

### New diagnostic tool

`debug_list_pages(section_id, top, orderby, count)` surfaces pagination behavior for a section — useful for investigating suspected truncation without editing source.

### Optional python-dotenv

The server now runs whether or not `python-dotenv` is installed. If you previously set `AZURE_CLIENT_ID` etc. via Claude Desktop's config, nothing changes. If you were using a `.env` file, install `python-dotenv` as before.

---

## Environment variables

| Variable | Required | Purpose |
|----------|----------|---------|
| `AZURE_CLIENT_ID` | Yes | Azure App Registration client ID |
| `AZURE_TENANT_ID` | No (default: `common`) | Azure tenant — set to your tenant ID for SharePoint |
| `SHAREPOINT_HOST` | No | SharePoint host, e.g. `contoso-my.sharepoint.com`. Required for SharePoint-backed notebooks. |
| `SHAREPOINT_USER_PATH` | No | Path to your personal site, e.g. `/personal/alice_contoso_com`. Required with `SHAREPOINT_HOST`. |
| `ONENOTE_CACHE_TOKENS` | No (default: `true`) | Set to `false` to disable on-disk token caching |

---

## Recursion depth

`enumerate_notebook` and `list_all_sections` take a `max_depth` parameter (default `10`). OneNote's section groups form a tree (no cycles), so this is a soft safety valve rather than a correctness requirement. If it's hit, a warning is logged and deeper branches are omitted — you won't get an error. Real notebooks almost never go deeper than 3–4 levels.

