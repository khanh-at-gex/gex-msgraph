# gex-msgraph

Async Python wrapper around Microsoft Graph API for clean access to M365 (SharePoint, OneDrive, Outlook, Teams) from any context — FastAPI, Prefect, scripts, notebooks.

**Status:** v0.2.0 | Python >=3.11 | Internal/Private

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
