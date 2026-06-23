#!/usr/bin/env python3
"""
OneNote MCP Server

A Model Context Protocol server for Microsoft OneNote integration.
Supports personal OneDrive notebooks and SharePoint / OneDrive-for-Business
notebooks. Handles OneNote's pagination quirks (nextLink omission, silent
100-record truncation) and exposes section-group / nested-folder tools.

Lineage:
    * Base:   purpleslurple/onenote-mcp-server
    * Fork:   peterstahley/onenote-mcp-server-sharepoint-compatible
              (adds SharePoint support, get_page_resources, restart_server)
    * This release adds:
        - Dual-strategy pagination (nextLink + manual $skip fallback)
          applied to every collection endpoint
        - list_section_groups / list_section_group_contents tools
        - enumerate_notebook tool with configurable recursion depth
        - list_all_sections flat view (with group_name tags)
        - debug_list_pages diagnostic tool
        - list_pages personal_err scope-bug fix
        - Optional python-dotenv (no hard dependency)
"""

import os
import asyncio
import json
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
import time
from msal import ConfidentialClientApplication, PublicClientApplication
import httpx
from fastmcp import FastMCP

# Optional .env support. The server works fine without python-dotenv — env
# vars set by Claude Desktop (or the shell) take precedence either way.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastMCP instance
mcp = FastMCP("OneNote MCP Server")

# Microsoft Graph API constants
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "https://graph.microsoft.com/Notes.Read",
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/Notes.Read.All",
    "https://graph.microsoft.com/Notes.ReadWrite.All",
    "https://graph.microsoft.com/User.Read"
]

# Tenant / SharePoint configuration
TENANT_ID = os.getenv("AZURE_TENANT_ID", "common")
SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST", "")
SHAREPOINT_USER_PATH = os.getenv("SHAREPOINT_USER_PATH", "")

# Token cache configuration
TOKEN_CACHE_ENABLED = os.getenv("ONENOTE_CACHE_TOKENS", "true").lower() in ("true", "1", "yes")
TOKEN_CACHE_FILE = Path.home() / ".onenote_mcp_tokens.json"

# Safe-default recursion depth for enumerate_notebook. No hard cap — set
# by caller. This is a soft warning threshold.
DEFAULT_ENUM_DEPTH = 10

# Global variables for authentication
access_token: Optional[str] = None
refresh_token: Optional[str] = None
token_expires_at: Optional[float] = None
msal_app: Optional[PublicClientApplication] = None

def get_client_id() -> str:
    """Get the Azure client ID from environment variable."""
    client_id = os.getenv("AZURE_CLIENT_ID")
    if not client_id:
        raise Exception("AZURE_CLIENT_ID environment variable not set")
    return client_id

def save_tokens(access_tok: str, refresh_tok: str = None, expires_in: int = 3600) -> None:
    """Save tokens to disk for persistence across sessions."""
    global access_token, refresh_token, token_expires_at

    access_token = access_tok
    if refresh_tok:
        refresh_token = refresh_tok
    token_expires_at = time.time() + expires_in - 300  # 5 min buffer

    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - tokens will not persist across sessions")
        return

    try:
        token_data = {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": token_expires_at
        }
        with open(TOKEN_CACHE_FILE, 'w') as f:
            json.dump(token_data, f)
        TOKEN_CACHE_FILE.chmod(0o600)
        logger.info(f"Tokens saved to {TOKEN_CACHE_FILE}")
    except Exception as e:
        logger.warning(f"Failed to save tokens: {e}")

def load_tokens() -> bool:
    """Load tokens from disk. Returns True if valid tokens loaded."""
    global access_token, refresh_token, token_expires_at

    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - will not load cached tokens")
        return False

    try:
        if not TOKEN_CACHE_FILE.exists():
            logger.info(f"No token cache file found at {TOKEN_CACHE_FILE}")
            return False

        with open(TOKEN_CACHE_FILE, 'r') as f:
            token_data = json.load(f)

        access_token = token_data.get("access_token")
        refresh_token = token_data.get("refresh_token")
        token_expires_at = token_data.get("expires_at")

        if token_expires_at and time.time() < token_expires_at:
            logger.info(f"Valid tokens loaded from {TOKEN_CACHE_FILE}")
            return True
        else:
            logger.info("Cached tokens expired")
            return False
    except Exception as e:
        logger.warning(f"Failed to load tokens: {e}")
        return False

async def refresh_access_token() -> bool:
    """Try to refresh the access token using the refresh token."""
    global access_token, msal_app

    if not refresh_token or not msal_app:
        return False

    try:
        accounts = msal_app.get_accounts()
        if accounts:
            result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                save_tokens(
                    result["access_token"],
                    result.get("refresh_token", refresh_token),
                    result.get("expires_in", 3600)
                )
                logger.info("Token refreshed successfully via MSAL silent acquisition")
                return True

        logger.info("MSAL silent acquisition failed, trying manual refresh with cached token")
        return await manual_token_refresh()
    except Exception as e:
        logger.warning(f"Token refresh error: {e}")
        return False

async def manual_token_refresh() -> bool:
    """Manually refresh access token using cached refresh token."""
    global access_token, refresh_token

    if not refresh_token:
        logger.info("No refresh token available for manual refresh")
        return False

    try:
        client_id = get_client_id()
        token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
            "scope": " ".join(SCOPES + ["offline_access"])
        }

        async with httpx.AsyncClient() as client:
            response = await client.post(
                token_url,
                data=data,
                headers={"Content-Type": "application/x-www-form-urlencoded"}
            )
            if response.status_code == 200:
                token_data = response.json()
                save_tokens(
                    token_data["access_token"],
                    token_data.get("refresh_token", refresh_token),
                    token_data.get("expires_in", 3600)
                )
                logger.info("Token refreshed successfully via manual refresh")
                return True
            else:
                logger.warning(f"Manual token refresh failed: {response.status_code} - {response.text}")
                return False
    except Exception as e:
        logger.warning(f"Manual token refresh error: {e}")
        return False

def init_msal_app(client_id: str) -> PublicClientApplication:
    """Initialize MSAL application for authentication."""
    return PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )

async def ensure_valid_token() -> bool:
    """Ensure we have a valid access token, refreshing if needed."""
    global access_token, msal_app

    if not access_token:
        load_tokens()

    if access_token and token_expires_at and time.time() < token_expires_at:
        return True

    if not msal_app:
        msal_app = init_msal_app(get_client_id())

    if await refresh_access_token():
        return True

    access_token = None
    return False

# Global variable to store the current authentication flow
current_flow = None

@mcp.tool()
async def start_authentication() -> str:
    """
    Start the full authentication process.

    Returns:
        Authentication instructions with device code
    """
    global access_token, msal_app, current_flow

    try:
        client_id = get_client_id()
        logger.info(f"Starting authentication with client_id: {client_id[:8]}...")

        if not msal_app:
            msal_app = init_msal_app(client_id)

        logger.info("Initiating device flow for authentication...")
        flow = msal_app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            error_msg = flow.get('error_description', 'Unknown error in device flow')
            raise Exception(f"Failed to create device flow: {error_msg}")

        result = {
            "status": "authentication_required",
            "instructions": f"Go to {flow['verification_uri']} and enter code: {flow['user_code']}",
            "verification_uri": flow['verification_uri'],
            "user_code": flow['user_code'],
            "expires_in": flow.get('expires_in', 900),
            "message": "Please complete authentication, then call 'complete_authentication'"
        }
        current_flow = flow
        return json.dumps(result, indent=2)

    except Exception as e:
        logger.error(f"Start authentication error: {str(e)}")
        return json.dumps({"status": "error", "error": str(e)}, indent=2)

@mcp.tool()
async def complete_authentication() -> str:
    """
    Complete the authentication process after user enters device code.

    Returns:
        Authentication status and user info
    """
    global access_token, msal_app, current_flow

    try:
        if not current_flow:
            return json.dumps({
                "status": "error",
                "error": "No authentication flow in progress. Call 'start_authentication' first."
            }, indent=2)

        if not msal_app:
            return json.dumps({
                "status": "error",
                "error": "MSAL app not initialized"
            }, indent=2)

        logger.info("Completing device flow authentication...")
        result = msal_app.acquire_token_by_device_flow(current_flow)

        if "access_token" in result:
            save_tokens(
                result["access_token"],
                result.get("refresh_token"),
                result.get("expires_in", 3600)
            )
            logger.info("Authentication successful and tokens cached!")

            try:
                user_info = await make_graph_request("/me")
                return json.dumps({
                    "status": "success",
                    "message": "Authentication completed successfully and tokens cached for future use",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown")
                }, indent=2)
            except Exception as graph_error:
                return json.dumps({
                    "status": "partial_success",
                    "message": "Got access token but Graph API test failed",
                    "graph_error": str(graph_error)
                }, indent=2)
        else:
            error_desc = result.get('error_description', 'Unknown authentication error')
            return json.dumps({
                "status": "error",
                "error": f"Authentication failed: {error_desc}"
            }, indent=2)

    except Exception as e:
        logger.error(f"Complete authentication error: {str(e)}")
        return json.dumps({"status": "error", "error": str(e)}, indent=2)
    finally:
        current_flow = None

@mcp.tool()
async def check_authentication() -> str:
    """
    Check current authentication status and token validity.

    Returns:
        Authentication status information
    """
    try:
        cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
        cache_file_exists = TOKEN_CACHE_FILE.exists() if TOKEN_CACHE_ENABLED else False

        if await ensure_valid_token():
            try:
                user_info = await make_graph_request("/me")
                time_until_expiry = int(token_expires_at - time.time()) if token_expires_at else 0

                return json.dumps({
                    "status": "authenticated",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown"),
                    "token_valid_for_seconds": max(0, time_until_expiry),
                    "token_valid_for_hours": round(max(0, time_until_expiry) / 3600, 1),
                    "token_caching": cache_status,
                    "cache_file_exists": cache_file_exists,
                    "cache_file_path": str(TOKEN_CACHE_FILE) if TOKEN_CACHE_ENABLED else "N/A"
                }, indent=2)
            except Exception as graph_error:
                return json.dumps({
                    "status": "token_invalid",
                    "error": str(graph_error),
                    "message": "Token exists but API call failed - may need re-authentication",
                    "token_caching": cache_status
                }, indent=2)
        else:
            return json.dumps({
                "status": "not_authenticated",
                "message": "No valid authentication token. Please call 'start_authentication'",
                "token_caching": cache_status,
                "cache_file_exists": cache_file_exists
            }, indent=2)

    except Exception as e:
        return json.dumps({
            "status": "error",
            "error": str(e),
            "token_caching": "unknown"
        }, indent=2)

# -----------------------------------------------------------------------------
# Graph API helpers
# -----------------------------------------------------------------------------

async def make_graph_request(endpoint: str, method: str = "GET", data: Dict = None) -> Dict:
    """Make a single request to Microsoft Graph API."""
    if not await ensure_valid_token():
        raise Exception("Not authenticated. Please call 'start_authentication' and 'complete_authentication' first.")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"{GRAPH_BASE_URL}{endpoint}"

    async with httpx.AsyncClient() as client:
        if method == "GET":
            response = await client.get(url, headers=headers)
        elif method == "POST":
            response = await client.post(url, headers=headers, json=data)
        elif method == "PATCH":
            response = await client.patch(url, headers=headers, json=data)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")

    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.status_code} - {response.text}")

    return response.json()

async def make_graph_request_all(endpoint: str) -> list:
    """Make a paginated GET request to Microsoft Graph API.

    Handles two pagination strategies:
    1. Follow @odata.nextLink if the API provides one
    2. Fall back to manual $skip pagination if nextLink is missing
       (OneNote pages API silently truncates at 100 without nextLink)
    """
    if not await ensure_valid_token():
        raise Exception("Not authenticated. Please call 'start_authentication' and 'complete_authentication' first.")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    results = []
    url = f"{GRAPH_BASE_URL}{endpoint}"
    sep = "&" if "?" in endpoint else "?"
    if "$top=" not in endpoint:
        url += f"{sep}$top=100"

    async with httpx.AsyncClient() as client:
        while url:
            response = await client.get(url, headers=headers)
            if response.status_code >= 400:
                raise Exception(f"Graph API error: {response.status_code} - {response.text}")
            data = response.json()
            batch = data.get("value", [])
            results.extend(batch)

            # Strategy 1: follow nextLink if provided
            next_link = data.get("@odata.nextLink")
            if next_link:
                url = next_link
                continue

            # Strategy 2: manual $skip if we got a full page but no nextLink
            # (OneNote pages API is known to omit nextLink)
            if len(batch) > 0 and len(batch) >= 100:
                base = f"{GRAPH_BASE_URL}{endpoint}"
                skip_sep = "&" if "?" in endpoint else "?"
                if "$top=" not in endpoint:
                    base += f"{skip_sep}$top=100"
                    skip_sep = "&"
                url = f"{base}{skip_sep}$skip={len(results)}"
            else:
                url = None

    return results

async def make_graph_request_binary(endpoint: str) -> bytes:
    """Make a Graph API request that returns binary content (images, files)."""
    if not await ensure_valid_token():
        raise Exception("Not authenticated.")

    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_BASE_URL}{endpoint}"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers)

    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.status_code} - {response.text}")

    return response.content

# Cache for resolved SharePoint site ID
_sharepoint_site_id: Optional[str] = None

async def get_sharepoint_site_id() -> Optional[str]:
    """Resolve and cache the SharePoint personal site ID for OneNote queries."""
    global _sharepoint_site_id

    if _sharepoint_site_id:
        return _sharepoint_site_id

    if not SHAREPOINT_HOST or not SHAREPOINT_USER_PATH:
        return None

    try:
        site_path = SHAREPOINT_USER_PATH.lstrip("/")
        result = await make_graph_request(f"/sites/{SHAREPOINT_HOST}:/{site_path}")
        _sharepoint_site_id = result.get("id")
        logger.info(f"Resolved SharePoint site ID: {_sharepoint_site_id}")
        return _sharepoint_site_id
    except Exception as e:
        logger.warning(f"Could not resolve SharePoint site ID: {e}")
        return None

# -----------------------------------------------------------------------------
# Notebook / section / page tools
# -----------------------------------------------------------------------------

def _parse_sections(raw_list, group_name=None):
    out = []
    for s in raw_list:
        links = s.get("links") or {}
        entry = {
            "id": s.get("id"),
            "name": s.get("displayName"),
            "created": s.get("createdDateTime"),
            "modified": s.get("lastModifiedDateTime"),
            "web_url": (links.get("oneNoteWebUrl") or {}).get("href"),
            "client_url": (links.get("oneNoteClientUrl") or {}).get("href"),
        }
        if group_name:
            entry["group_name"] = group_name
        out.append(entry)
    return out

def _parse_section_groups(raw_list):
    out = []
    for g in raw_list:
        links = g.get("links") or {}
        out.append({
            "id": g.get("id"),
            "name": g.get("displayName"),
            "created": g.get("createdDateTime"),
            "modified": g.get("lastModifiedDateTime"),
            "sections_url": g.get("sectionsUrl"),
            "section_groups_url": g.get("sectionGroupsUrl"),
            "web_url": (links.get("oneNoteWebUrl") or {}).get("href"),
            "client_url": (links.get("oneNoteClientUrl") or {}).get("href"),
        })
    return out

@mcp.tool()
async def list_notebooks() -> str:
    """
    List all OneNote notebooks, including SharePoint-backed notebooks.
    Paginated — returns every notebook regardless of count.

    Returns:
        JSON string containing notebook information
    """
    try:
        seen_ids = set()
        result = []

        def add_notebooks(raw_list):
            for notebook in raw_list:
                nb_id = notebook.get("id")
                if nb_id and nb_id not in seen_ids:
                    seen_ids.add(nb_id)
                    links = notebook.get("links") or {}
                    result.append({
                        "id": nb_id,
                        "name": notebook.get("displayName"),
                        "created": notebook.get("createdDateTime"),
                        "modified": notebook.get("lastModifiedDateTime"),
                        "web_url": (links.get("oneNoteWebUrl") or {}).get("href"),
                        "client_url": (links.get("oneNoteClientUrl") or {}).get("href"),
                    })

        # Personal OneDrive notebooks
        logger.info("Fetching personal notebooks from /me/onenote/notebooks")
        try:
            personal_list = await make_graph_request_all("/me/onenote/notebooks")
            add_notebooks(personal_list)
            logger.info(f"Found {len(result)} personal notebooks")
        except Exception as e:
            logger.warning(f"Could not fetch personal notebooks: {e}")

        # SharePoint-backed notebooks
        site_id = await get_sharepoint_site_id()
        if site_id:
            logger.info(f"Fetching SharePoint notebooks from /sites/{site_id}/onenote/notebooks")
            try:
                sp_list = await make_graph_request_all(f"/sites/{site_id}/onenote/notebooks")
                before = len(result)
                add_notebooks(sp_list)
                logger.info(f"Found {len(result) - before} additional SharePoint notebooks")
            except Exception as e:
                logger.warning(f"Could not fetch SharePoint notebooks: {e}")

        logger.info(f"Returning {len(result)} total notebooks")
        return json.dumps(result, indent=2)

    except Exception as e:
        logger.error(f"Error in list_notebooks: {str(e)}")
        return f"Error listing notebooks: {str(e)}"

@mcp.tool()
async def list_sections(notebook_id: str) -> str:
    """
    List root-level sections in a specific notebook (does NOT recurse into
    section groups). For a flat list of every section including those in
    nested folders, use list_all_sections. For the full tree, use
    enumerate_notebook.

    Args:
        notebook_id: ID of the notebook to list sections from

    Returns:
        JSON string containing section information
    """
    try:
        section_list = await make_graph_request_all(
            f"/me/onenote/notebooks/{notebook_id}/sections"
        )
        return json.dumps(_parse_sections(section_list), indent=2)
    except Exception as e:
        return f"Error listing sections: {str(e)}"

@mcp.tool()
async def list_all_sections(notebook_id: str, max_depth: int = DEFAULT_ENUM_DEPTH) -> str:
    """
    List every section in a notebook as a flat list, including sections inside
    nested section groups. Each result tagged with a 'group_name' field
    (breadcrumbs like 'Engineering > Q1 > Retros') when inside a group.

    Args:
        notebook_id: ID of the notebook
        max_depth: Soft recursion limit for nested section groups (default 10).
                   Logs a warning if hit but does not error.

    Returns:
        JSON array of sections with optional group_name breadcrumb
    """
    all_sections: List[Dict] = []

    async def _walk_group(group_id: str, breadcrumb: str, depth: int) -> None:
        if depth > max_depth:
            logger.warning(
                f"list_all_sections: max_depth={max_depth} reached at '{breadcrumb}' — "
                f"deeper sections will be omitted"
            )
            return
        try:
            sections_in_group = await make_graph_request_all(
                f"/me/onenote/sectionGroups/{group_id}/sections"
            )
            all_sections.extend(_parse_sections(sections_in_group, group_name=breadcrumb))
        except Exception as e:
            logger.warning(f"Could not fetch sections in group '{breadcrumb}': {e}")

        try:
            nested_groups = await make_graph_request_all(
                f"/me/onenote/sectionGroups/{group_id}/sectionGroups"
            )
            for ng in nested_groups:
                ng_name = ng.get("displayName", "?")
                ng_crumb = f"{breadcrumb} > {ng_name}"
                await _walk_group(ng.get("id"), ng_crumb, depth + 1)
        except Exception as e:
            logger.warning(f"Could not fetch nested groups of '{breadcrumb}': {e}")

    try:
        # Root-level sections
        try:
            root_sections = await make_graph_request_all(
                f"/me/onenote/notebooks/{notebook_id}/sections"
            )
            all_sections.extend(_parse_sections(root_sections))
        except Exception as e:
            logger.warning(f"Could not fetch root sections: {e}")

        # Top-level section groups
        try:
            root_groups = await make_graph_request_all(
                f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
            )
            for group in root_groups:
                await _walk_group(group.get("id"), group.get("displayName", "?"), depth=1)
        except Exception as e:
            logger.warning(f"Could not fetch top-level section groups: {e}")

        return json.dumps(all_sections, indent=2)

    except Exception as e:
        return f"Error listing all sections: {str(e)}"

@mcp.tool()
async def list_section_groups(notebook_id: str) -> str:
    """
    List top-level section groups (folders) in a specific notebook.
    Use list_section_group_contents to descend into a group.

    Args:
        notebook_id: ID of the notebook

    Returns:
        JSON string containing section group information
    """
    try:
        group_list = await make_graph_request_all(
            f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
        )
        return json.dumps(_parse_section_groups(group_list), indent=2)
    except Exception as e:
        return f"Error listing section groups: {str(e)}"

@mcp.tool()
async def list_sections_in_group(group_id: str) -> str:
    """
    List sections directly inside a specific section group (one level,
    non-recursive).

    Args:
        group_id: ID of the section group

    Returns:
        JSON string containing section information
    """
    try:
        section_list = await make_graph_request_all(
            f"/me/onenote/sectionGroups/{group_id}/sections"
        )
        return json.dumps(_parse_sections(section_list), indent=2)
    except Exception as e:
        return f"Error listing sections in group: {str(e)}"

@mcp.tool()
async def list_section_group_contents(section_group_id: str) -> str:
    """
    List both sections and nested section groups inside a section group
    (one level, non-recursive). Useful for building a lazy file-tree UI.

    Args:
        section_group_id: ID of the section group

    Returns:
        JSON with 'sections' and 'section_groups' arrays
    """
    try:
        section_list = await make_graph_request_all(
            f"/me/onenote/sectionGroups/{section_group_id}/sections"
        )
        nested_group_list = await make_graph_request_all(
            f"/me/onenote/sectionGroups/{section_group_id}/sectionGroups"
        )

        return json.dumps({
            "sections": _parse_sections(section_list),
            "section_groups": _parse_section_groups(nested_group_list)
        }, indent=2)
    except Exception as e:
        return f"Error listing section group contents: {str(e)}"

@mcp.tool()
async def enumerate_notebook(
    notebook_id: str,
    include_pages: bool = False,
    max_depth: int = DEFAULT_ENUM_DEPTH
) -> str:
    """
    Recursively enumerate the full structure of a notebook as a tree:
    root sections plus all section groups with their (potentially nested)
    sections, and optionally all pages within each section.

    Args:
        notebook_id: ID of the notebook to enumerate
        include_pages: If True, list pages within each section (slower)
        max_depth: Soft recursion limit for nested section groups (default 10).
                   Logs a warning if hit but does not error.

    Returns:
        JSON tree with sections and section_groups (each potentially containing
        further nested contents)
    """
    async def _enumerate_group(group_id: str, depth: int) -> Dict:
        if depth > max_depth:
            logger.warning(
                f"enumerate_notebook: max_depth={max_depth} reached — "
                f"deeper content will be omitted"
            )
            return {"error": f"Max nesting depth ({max_depth}) reached"}

        sections_list = await make_graph_request_all(
            f"/me/onenote/sectionGroups/{group_id}/sections"
        )
        nested_list = await make_graph_request_all(
            f"/me/onenote/sectionGroups/{group_id}/sectionGroups"
        )

        sections = []
        for s in sections_list:
            section_info = {
                "id": s.get("id"),
                "name": s.get("displayName"),
                "modified": s.get("lastModifiedDateTime"),
                "web_url": ((s.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                "client_url": ((s.get("links") or {}).get("oneNoteClientUrl") or {}).get("href"),
            }
            if include_pages:
                pages = await make_graph_request_all(
                    f"/me/onenote/sections/{s['id']}/pages?$top=100"
                )
                section_info["pages"] = [
                    {
                        "id": p.get("id"),
                        "title": p.get("title"),
                        "modified": p.get("lastModifiedDateTime"),
                        "content_url": p.get("contentUrl"),
                        "web_url": ((p.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                        "client_url": ((p.get("links") or {}).get("oneNoteClientUrl") or {}).get("href")
                    }
                    for p in pages
                ]
            sections.append(section_info)

        groups = []
        for g in nested_list:
            group_info = {
                "id": g.get("id"),
                "name": g.get("displayName"),
                "modified": g.get("lastModifiedDateTime"),
                "web_url": ((g.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                "client_url": ((g.get("links") or {}).get("oneNoteClientUrl") or {}).get("href"),
                "contents": await _enumerate_group(g["id"], depth + 1)
            }
            groups.append(group_info)

        return {"sections": sections, "section_groups": groups}

    try:
        top_sections_list = await make_graph_request_all(
            f"/me/onenote/notebooks/{notebook_id}/sections"
        )
        top_groups_list = await make_graph_request_all(
            f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
        )

        top_sections = []
        for s in top_sections_list:
            section_info = {
                "id": s.get("id"),
                "name": s.get("displayName"),
                "modified": s.get("lastModifiedDateTime"),
                "web_url": ((s.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                "client_url": ((s.get("links") or {}).get("oneNoteClientUrl") or {}).get("href"),
            }
            if include_pages:
                pages = await make_graph_request_all(
                    f"/me/onenote/sections/{s['id']}/pages?$top=100"
                )
                section_info["pages"] = [
                    {
                        "id": p.get("id"),
                        "title": p.get("title"),
                        "modified": p.get("lastModifiedDateTime"),
                        "content_url": p.get("contentUrl"),
                        "web_url": ((p.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                        "client_url": ((p.get("links") or {}).get("oneNoteClientUrl") or {}).get("href")
                    }
                    for p in pages
                ]
            top_sections.append(section_info)

        top_groups = []
        for g in top_groups_list:
            group_info = {
                "id": g.get("id"),
                "name": g.get("displayName"),
                "modified": g.get("lastModifiedDateTime"),
                "web_url": ((g.get("links") or {}).get("oneNoteWebUrl") or {}).get("href"),
                "client_url": ((g.get("links") or {}).get("oneNoteClientUrl") or {}).get("href"),
                "contents": await _enumerate_group(g["id"], depth=1)
            }
            top_groups.append(group_info)

        return json.dumps({
            "sections": top_sections,
            "section_groups": top_groups
        }, indent=2)

    except Exception as e:
        return f"Error enumerating notebook: {str(e)}"

@mcp.tool()
async def list_pages(section_id: str) -> str:
    """
    List all pages in a specific section. Paginated — returns every page
    regardless of count (works around OneNote's silent 100-record truncation).
    Falls back to the SharePoint endpoint if the personal endpoint fails.

    Args:
        section_id: ID of the section to list pages from

    Returns:
        JSON string containing page information
    """
    def parse_pages(raw_list):
        out = []
        for p in raw_list:
            links = p.get("links") or {}
            out.append({
                "id": p.get("id"),
                "title": p.get("title"),
                "created": p.get("createdDateTime"),
                "modified": p.get("lastModifiedDateTime"),
                "content_url": p.get("contentUrl"),
                "web_url": (links.get("oneNoteWebUrl") or {}).get("href"),
                "client_url": (links.get("oneNoteClientUrl") or {}).get("href")
            })
        return out

    # Try personal endpoint first. Bind personal_err at outer scope so it
    # stays visible after the except block (Python 3 clears `as` names).
    personal_err: Optional[Exception] = None
    try:
        pages = await make_graph_request_all(
            f"/me/onenote/sections/{section_id}/pages?$top=100"
        )
        # Even if the list is empty, prefer the personal result unless an
        # actual error occurred.
        return json.dumps(parse_pages(pages), indent=2)
    except Exception as e:
        personal_err = e
        logger.info(f"Personal list_pages failed, will try SharePoint: {e}")

    # Fall back to SharePoint endpoint (handles sections in OneDrive for
    # Business notebooks where the /me/ endpoint 404s).
    site_id = await get_sharepoint_site_id()
    if site_id:
        try:
            pages = await make_graph_request_all(
                f"/sites/{site_id}/onenote/sections/{section_id}/pages?$top=100"
            )
            return json.dumps(parse_pages(pages), indent=2)
        except Exception as sp_err:
            return (
                f"Error listing pages (personal: {personal_err}; "
                f"sharepoint: {sp_err})"
            )

    return f"Error listing pages: {personal_err}"

@mcp.tool()
async def debug_list_pages(
    section_id: str,
    top: int = 20,
    orderby: str = "",
    count: bool = False
) -> str:
    """
    Diagnostic tool: list pages with custom query parameters to inspect
    pagination behavior. Use this to investigate suspected truncation.

    Args:
        section_id: ID of the section
        top: Number of results per page ($top parameter)
        orderby: Sort order e.g. "createdDateTime asc" or
                 "lastModifiedDateTime desc"
        count: If true, request $count=true to get total count from API

    Returns:
        JSON with pages, total count, request count, and whether a nextLink
        was present on any response
    """
    try:
        if not await ensure_valid_token():
            return "Not authenticated"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        params = f"$top={top}"
        if orderby:
            params += f"&$orderby={orderby}"
        if count:
            params += "&$count=true&ConsistencyLevel=eventual"

        all_pages: List[Dict] = []
        request_count = 0
        odata_count: Any = None
        had_next_link = False
        url: Optional[str] = f"{GRAPH_BASE_URL}/me/onenote/sections/{section_id}/pages?{params}"

        async with httpx.AsyncClient() as client:
            while url:
                request_count += 1
                if count:
                    headers["ConsistencyLevel"] = "eventual"
                response = await client.get(url, headers=headers)
                if response.status_code >= 400:
                    return f"Error {response.status_code}: {response.text}"
                data = response.json()
                batch = data.get("value", [])
                all_pages.extend(batch)
                next_link = data.get("@odata.nextLink")
                if next_link:
                    had_next_link = True
                odata_count = data.get("@odata.count", odata_count)
                url = next_link

        titles = [p.get("title", "untitled") for p in all_pages]
        return json.dumps({
            "total_returned": len(all_pages),
            "odata_count": odata_count if count else "not requested",
            "requests_made": request_count,
            "had_next_link": had_next_link,
            "first_title": titles[0] if titles else None,
            "last_title": titles[-1] if titles else None,
            "all_titles": titles
        }, indent=2)

    except Exception as e:
        return f"Error: {str(e)}"

async def _fetch_page_content_html(page_id: str) -> str:
    """Internal helper: fetch page HTML with SharePoint fallback.

    Used by both the get_page_content @mcp.tool and get_page_resources.
    Direct calls between @mcp.tool() functions fail with
    "FunctionTool object is not callable" in fastmcp, so the body lives here.
    """
    if not await ensure_valid_token():
        return "Error: Not authenticated."

    headers = {"Authorization": f"Bearer {access_token}"}

    async def fetch_content(url: str) -> Optional[str]:
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=headers)
        if response.status_code < 400:
            return response.text
        return None

    # Try personal endpoint first
    content = await fetch_content(f"{GRAPH_BASE_URL}/me/onenote/pages/{page_id}/content")
    if content is not None:
        return content

    # Fall back to SharePoint endpoint
    site_id = await get_sharepoint_site_id()
    if site_id:
        content = await fetch_content(
            f"{GRAPH_BASE_URL}/sites/{site_id}/onenote/pages/{page_id}/content"
        )
        if content is not None:
            return content

    return f"Error getting page content: page {page_id} not found via personal or SharePoint endpoints"

@mcp.tool()
async def get_page_content(page_id: str) -> str:
    """
    Get the HTML content of a specific page. Falls back to the SharePoint
    endpoint if the personal endpoint returns an error.

    Args:
        page_id: ID of the page to retrieve content from

    Returns:
        Page content as HTML or error message
    """
    return await _fetch_page_content_html(page_id)

@mcp.tool()
async def get_page_resources(page_id: str) -> str:
    """
    Fetch and save all image/attachment resources embedded in a OneNote page.

    Useful for pages with handwritten ink (Boox, Surface) or embedded images
    that don't render as plain text. Saves each resource as a file under
    ~/.onenote_mcp_cache/<page_id>/ and returns a manifest.

    Args:
        page_id: ID of the OneNote page

    Returns:
        JSON manifest of saved resource files with their local paths and MIME types
    """
    import re
    import mimetypes
    import hashlib

    if not await ensure_valid_token():
        return "Error: Not authenticated."

    content_html = await _fetch_page_content_html(page_id)
    if content_html.startswith("Error"):
        return content_html

    cache_dir = Path.home() / ".onenote_mcp_cache" / page_id
    cache_dir.mkdir(parents=True, exist_ok=True)

    resource_pattern = re.compile(
        r'https://graph\.microsoft\.com/v1\.0[^\s"\'<>]+/\$value',
        re.IGNORECASE
    )
    resource_urls = list(dict.fromkeys(resource_pattern.findall(content_html)))

    headers = {"Authorization": f"Bearer {access_token}"}
    manifest = []

    async with httpx.AsyncClient() as client:
        for url in resource_urls:
            try:
                response = await client.get(url, headers=headers)
                if response.status_code >= 400:
                    logger.warning(f"Failed to fetch resource {url}: {response.status_code}")
                    continue

                content_type = response.headers.get(
                    "content-type", "application/octet-stream"
                ).split(";")[0]
                ext = mimetypes.guess_extension(content_type) or ".bin"
                name = hashlib.md5(url.encode()).hexdigest()[:12] + ext
                file_path = cache_dir / name
                file_path.write_bytes(response.content)

                manifest.append({
                    "url": url,
                    "local_path": str(file_path),
                    "mime_type": content_type,
                    "size_bytes": len(response.content)
                })
                logger.info(f"Saved resource: {file_path}")

            except Exception as e:
                logger.warning(f"Error fetching resource {url}: {e}")

    return json.dumps({
        "page_id": page_id,
        "cache_dir": str(cache_dir),
        "resources_found": len(resource_urls),
        "resources_saved": len(manifest),
        "files": manifest
    }, indent=2)

@mcp.tool()
async def clear_token_cache() -> str:
    """
    Clear the stored authentication tokens.

    Returns:
        Status message
    """
    global access_token, refresh_token, token_expires_at

    try:
        access_token = None
        refresh_token = None
        token_expires_at = None

        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()

        return json.dumps({
            "status": "success",
            "message": "Token cache cleared. You will need to re-authenticate."
        }, indent=2)

    except Exception as e:
        return json.dumps({"status": "error", "error": str(e)}, indent=2)

@mcp.tool()
async def restart_server() -> str:
    """
    Restart the MCP server process. Use this to pick up code changes without
    restarting Claude. The server will automatically reconnect.

    Returns:
        Status message (sent before exit)
    """
    import threading
    threading.Timer(0.5, lambda: os._exit(0)).start()
    return json.dumps({"status": "restarting", "message": "Server is restarting..."}, indent=2)

@mcp.tool()
async def create_notebook(name: str, description: str = None) -> str:
    """
    Create a new OneNote notebook.

    Args:
        name: Name of the new notebook
        description: Optional description for the notebook

    Returns:
        JSON string with the created notebook information
    """
    try:
        data = {"displayName": name}
        if description:
            data["description"] = description

        notebook = await make_graph_request("/me/onenote/notebooks", method="POST", data=data)

        return json.dumps({
            "status": "success",
            "message": f"Notebook '{name}' created successfully",
            "notebook": {
                "id": notebook.get("id"),
                "name": notebook.get("displayName"),
                "created": notebook.get("createdDateTime")
            }
        }, indent=2)

    except Exception as e:
        return f"Error creating notebook: {str(e)}"

@mcp.tool()
async def create_section(notebook_id: str, name: str) -> str:
    """
    Create a new section in a OneNote notebook.

    Args:
        notebook_id: ID of the notebook to create the section in
        name: Name of the new section

    Returns:
        JSON string with the created section information
    """
    try:
        data = {"displayName": name}

        section = await make_graph_request(
            f"/me/onenote/notebooks/{notebook_id}/sections",
            method="POST",
            data=data
        )

        return json.dumps({
            "status": "success",
            "message": f"Section '{name}' created successfully",
            "section": {
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime")
            }
        }, indent=2)

    except Exception as e:
        return f"Error creating section: {str(e)}"

@mcp.tool()
async def create_page(section_id: str, title: str, content_html: str = None) -> str:
    """
    Create a new page in a OneNote section.

    Args:
        section_id: ID of the section to create the page in
        title: Title of the new page
        content_html: Optional HTML content for the page body

    Returns:
        JSON string with the created page information
    """
    try:
        if content_html:
            if not content_html.strip().startswith('<html>'):
                page_html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{time.strftime('%Y-%m-%dT%H:%M:%S.0000000')}" />
</head>
<body>
    <div>
        <h1>{title}</h1>
        <div>{content_html}</div>
    </div>
</body>
</html>"""
            else:
                page_html = content_html
        else:
            page_html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{time.strftime('%Y-%m-%dT%H:%M:%S.0000000')}" />
</head>
<body>
    <div>
        <h1>{title}</h1>
        <p>Page created by OneNote MCP Server</p>
    </div>
</body>
</html>"""

        async with httpx.AsyncClient() as client:
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/xhtml+xml"
            }
            response = await client.post(
                f"{GRAPH_BASE_URL}/me/onenote/sections/{section_id}/pages",
                headers=headers,
                content=page_html
            )
            if response.status_code >= 400:
                return f"Error creating page: {response.status_code} - {response.text}"
            page = response.json()

        return json.dumps({
            "status": "success",
            "message": f"Page '{title}' created successfully",
            "page": {
                "id": page.get("id"),
                "title": page.get("title"),
                "created": page.get("createdDateTime"),
                "content_url": page.get("contentUrl")
            }
        }, indent=2)

    except Exception as e:
        return f"Error creating page: {str(e)}"

@mcp.tool()
async def update_page_content(
    page_id: str,
    content_html: str,
    target_element: str = "body"
) -> str:
    """
    Update the content of an existing OneNote page.

    Args:
        page_id: ID of the page to update
        content_html: New HTML content to add/replace
        target_element: Target element to update (default: "body")

    Returns:
        Status message
    """
    try:
        patch_data = [
            {
                "target": target_element,
                "action": "append",
                "content": content_html
            }
        ]

        async with httpx.AsyncClient() as client:
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            response = await client.patch(
                f"{GRAPH_BASE_URL}/me/onenote/pages/{page_id}/content",
                headers=headers,
                json=patch_data
            )
            if response.status_code >= 400:
                return f"Error updating page: {response.status_code} - {response.text}"

        return json.dumps({
            "status": "success",
            "message": "Page content updated successfully",
            "page_id": page_id
        }, indent=2)

    except Exception as e:
        return f"Error updating page content: {str(e)}"

def main():
    """Main entry point for the server."""
    cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
    logger.info(f"OneNote MCP Server starting - Token caching: {cache_status}")

    if TOKEN_CACHE_ENABLED:
        logger.info(f"Token cache file: {TOKEN_CACHE_FILE}")
        if load_tokens():
            logger.info("Cached tokens loaded successfully")
        else:
            logger.info("No valid cached tokens found")

    if SHAREPOINT_HOST and SHAREPOINT_USER_PATH:
        logger.info(f"SharePoint support enabled: {SHAREPOINT_HOST}{SHAREPOINT_USER_PATH}")
    else:
        logger.info("SharePoint support disabled (SHAREPOINT_HOST / SHAREPOINT_USER_PATH not set)")

    mcp.run()

if __name__ == "__main__":
    main()
