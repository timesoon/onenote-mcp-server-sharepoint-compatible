# Changelog

## [Unreleased] — proposed for merge into peterstahley/onenote-mcp-server-sharepoint-compatible

### Added
- **Pagination for every collection endpoint** via new `make_graph_request_all` helper. Handles both of OneNote's pagination modes:
  1. Standard OData `@odata.nextLink` follow
  2. Manual `$skip` fallback when the API omits `nextLink` (OneNote's pages endpoint silently truncates at 100 records without a nextLink — this was the original motivation)
- `list_section_groups(notebook_id)` — top-level folders in a notebook, now including `sections_url` and `section_groups_url` fields so callers can descend without another lookup.
- `list_section_group_contents(section_group_id)` — sections and nested groups inside a group (one level, non-recursive).
- `list_sections_in_group(group_id)` — sections directly inside one group.
- `list_all_sections(notebook_id, max_depth=10)` — flat list of every section across the entire notebook, each tagged with a breadcrumb `group_name` like `"Engineering > Q1 > Retros"`. This replaces the old flatten-on-fetch behavior of `list_sections` in a separate, opt-in tool.
- `enumerate_notebook(notebook_id, include_pages=False, max_depth=10)` — recursive tree of the entire notebook including optional pages-per-section.
- `debug_list_pages(section_id, top, orderby, count)` — diagnostic tool for investigating pagination behavior.
- **Page URLs surfaced** in `list_pages` and `enumerate_notebook` (when `include_pages=True`). Each page now includes `web_url` (opens in OneNote on the web) and `client_url` (deep-links into the OneNote desktop app) alongside the existing `content_url` (Graph API content endpoint). Pulled from the `links.oneNoteWebUrl.href` / `links.oneNoteClientUrl.href` fields that Graph already returns by default — no `$select` needed.
- **Section and notebook URLs surfaced** in `list_notebooks`, `list_sections`, `list_all_sections`, `list_sections_in_group`, `list_section_groups`, `list_section_group_contents`, and the section / group entries in `enumerate_notebook`. Same `web_url` / `client_url` shape as pages. Section groups fall through to `None` for these fields if Graph doesn't supply them for groups, which is harmless.

### Fixed (post-release-1 patches)
- `get_page_resources` threw `'FunctionTool' object is not callable` because it tried to invoke `get_page_content` (an `@mcp.tool()`-decorated function) from inside another tool. Fixed by extracting `_fetch_page_content_html(page_id)` as a plain async helper; both `get_page_content` and `get_page_resources` now delegate to it.

### Changed
- **BREAKING (minor):** `list_sections(notebook_id)` now returns **root-level sections only**. The previous fork-added behavior of flattening all sections (including those in section groups) has moved to `list_all_sections`. This restores backwards compatibility with the upstream base and composes cleanly with the new section-group tools.
- `get_page_content` retains SharePoint fallback but now uses the shared `fetch_content` helper for consistency.
- `list_pages` now paginated and retains SharePoint fallback. Fixed a latent Python 3 scoping bug where the `as personal_err` name went out of scope after the `except` block, which would have triggered a `NameError` if both personal and SharePoint endpoints failed.
- `python-dotenv` is now **optional**. The server logs and continues if the package isn't installed; env vars set by Claude Desktop continue to work.
- `main()` startup log now reports whether SharePoint is configured.

### Fixed
- Silent truncation at ~100 records in `list_pages`, `list_notebooks`, `list_sections`, `list_section_groups`, and anywhere a OneNote collection is fetched. Every collection call now paginates.
- `list_pages` `personal_err` scoping bug (Python 3 clears `as` names on exit from `except`).

### Internal
- New private helpers `_parse_sections` and `_parse_section_groups` consolidate result shaping.
- `make_graph_request_binary` retained from peter's fork (used for future image/resource work).
- Recursion is governed by a `max_depth` parameter (default **10**, no hard cap). Exceeding the limit logs a warning and omits deeper content rather than erroring — this is a soft cap for pathological notebook structures.

### Migration notes for existing users
If you were relying on `list_sections` returning every section including those inside section groups, switch to `list_all_sections(notebook_id)` — same output shape, new tool name. Everything else is strictly additive.
