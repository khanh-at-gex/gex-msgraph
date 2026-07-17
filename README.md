# gex-msgraph

Async Python wrapper around Microsoft Graph API for clean access to M365 (SharePoint, OneDrive, Outlook, Teams) from any context — FastAPI, Prefect, scripts, notebooks.

**Status:** v0.3.0 | Python >=3.11 | Internal/Private

## Quickstart

```python
from gex_msgraph import GraphClient
import asyncio

async def main():
    # Credentials loaded from MS_DAS_U1_* env vars automatically
    async with GraphClient("das_u1") as client:
        # Read Excel directly to a pandas DataFrame
        df = await client.read_excel(item_path="Reports/Q1.xlsx")
        print(df.head())

if __name__ == "__main__":
    asyncio.run(main())
```

## Documentation

- **[USAGE.md](USAGE.md)** — installation, setup, recipes, API reference, and deployment.

## For Contributors

See **[CLAUDE.md](CLAUDE.md)** for architecture, invariants, and development workflow.

### Release Process
Bump `__version__` in `__init__.py`, update the Changelog below, `git tag v0.X.Y`, `git push --tags`.

## Changelog

### [0.3.0] - 2026-07-17

#### Breaking Changes
- `move_file` / `copy_file`: the positional `source_path` parameter is now the keyword-only `item_path` / `share_url` / `item_id` trio (matching every other method). Migrate `move_file("a.xlsx", ...)` → `move_file(item_path="a.xlsx", ...)`.
- `GraphClient(default_site_id=...)` removed — it was accepted but never used. Pre-resolve a site's drive and pass `default_drive_id` instead.
- `python-dotenv` is no longer a hard dependency (the library never imported it). If you rely on `load_dotenv()`, add `python-dotenv` to your own project or install `gex-msgraph[dotenv]`.

#### Added
- **Exception hierarchy**: `GraphError` (subclasses `httpx.HTTPStatusError`, so existing `except httpx.HTTPStatusError` code keeps working), with `AuthError` (401), `PermissionDeniedError` (403), `NotFoundError` (404), `RateLimitExhaustedError` (retries exhausted), plus `GraphAuthenticationError` (MSAL token failure) and `GraphSyncInLoopError`.
- **App-only auth**: omit both `username` and `password` to use the client-credentials flow. App-only supports drive operations only (requires `default_drive_id`); mail/Teams/chat methods still require delegated auth.
- `copy_file(..., wait=True, wait_timeout=60.0)` — polls Graph's copy job until completion and returns the new item's `FileItem`.
- `download(..., to_path=...)` — stream a file straight to disk (chunked, not buffered in memory); returns the written `Path`.
- `upload` now transparently uses a chunked Graph upload session for files over 4 MiB (previously failed with HTTP 413).
- `close_sync()` — synchronous companion to `close()` for scripts using the `*_sync` wrappers.

#### Fixed
- `*_sync` wrappers no longer break on repeated calls (they previously created a new event loop per call while reusing loop-bound resources). They now run on a persistent background loop, and raise a clear `GraphSyncInLoopError` when called from inside a running loop (e.g. Jupyter) instead of a confusing asyncio error.

### [0.2.1] - 2026-07-13

#### Added
- `FileItem.webUrl` — Graph's `driveItem.webUrl`, for building Office Online embed URLs (`?action=embedview`). Unlike `get_share_link`'s `/:x:/` sharing links, this is safe to use in an iframe (no X-Frame-Options/CSP block).
- `upload_many` now accepts `max_concurrent` (previously only `read_excel_many`/`read_csv_many` did).

#### Fixed
- Package version now resolves consistently (`pyproject.toml` reads it dynamically from `__init__.__version__` instead of drifting out of sync).
- `download` now retries transient 429/5xx errors when fetching the file bytes themselves (previously only the metadata/resolve step retried).
- `search_files(limit=...)` now actually caps the number of results returned; previously `limit` only hinted at page size and pagination could return far more.
- `get_folder_tree` now recurses into sibling subfolders concurrently (matching `walk`'s existing behavior) instead of sequentially, speeding up deep/wide trees.

#### Changed
- Internal: `read_excel_many`, `read_csv_many`, and `upload_many` now share a common concurrency/status-tracking helper — no behavior change for existing callers.

### [0.2.0] - 2026-05-12

#### Added
- `read_parquet` — read Parquet files directly into a DataFrame.
- `copy_file` — copy a file to a new location/name.
- `exists` — check if a file or folder exists without raising on 404.
- `upload_many` — concurrent bulk upload with error tolerance and status reporting.
- `get_share_link` — create a view/edit sharing link for any item.
- `search_files` — search files across the drive by name or content.
- `list_mail` — list messages from any mail folder (default: inbox).
- `send_mail` now supports HTML body (`body_type="html"`) and file attachments (local paths, in-memory bytes, or SharePoint identifiers).
- `read_excel_many_sync` / `read_csv_many_sync` — sync wrappers for bulk read methods.

#### Changed
- `python-calamine` is now a hard dependency and the default Excel engine (previously optional).

### [0.1.0] - 2026-05-03

#### Added
- Async Python wrapper around Microsoft Graph API for M365 access.
- `GraphClient` public class with ROPC auth flow and connection management.
- `FileItem` dataclass for SharePoint/OneDrive file representation.
- Read files single/bulk: `read_excel`, `read_csv`, `download`, `read_excel_many`, `read_csv_many`.
- Discovery endpoints: `walk`, `list_files`, `get_folder_tree`, `get_metadata`.
- File management: `upload`, `delete_file`, `move_file`, `create_folder`, `list_excel_sheets`.
- Communication endpoints: `send_mail`, `send_teams_message`, `get_teams_messages`, `list_chats`, `get_chat_messages`, `send_chat_message`.
- Sync wrappers for ease of use in scripts/notebooks: `read_excel_sync`, `read_csv_sync`, `download_sync`.
- Automatic retry logic and concurrency capping per account.
