# gex-msgraph

Async Python wrapper around Microsoft Graph API for clean access to M365 (SharePoint, OneDrive, Outlook, Teams) from any context — FastAPI, Prefect, scripts, notebooks.

**Status:** v0.1.0 | Python >=3.11 | Internal/Private

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

## For Contributors / AI Assistants

### Architecture
`gex-msgraph` is an async Python wrapper for Microsoft Graph API focusing on M365 file and communication access.

### Module Map
- `_core.py`: GraphClient, token provider, request retry logic.
- `_files.py`: URL resolution, sheet matching, FileItem dataclass.
- `__init__.py`: Public exports.

### Conventions
- `from __future__ import annotations`
- Strict type hints on public functions.
- Private modules/functions are underscore-prefixed.

### Restrictions
- Do not add dependencies without discussion.
- Do not log secrets.
- All Graph calls MUST go through `_request`.

### Adding a Method
Add signature to `_core.py` → implement → add test in `test_client.py` → add recipe in `USAGE.md`.

### Adding Identifier Type
Touches `_files.py` only (`validate_identifier`, `build_resolution_url`).

### Consuming the Library (AI Assistants)
- Always check the recipes in `USAGE.md` for standard patterns.
- Prefer `await client.read_excel()` and `await client.read_excel_many()` for reading files.
- `item_path` is relative to the drive root, e.g. `Reports/Q1.xlsx`.
- Error handling: use `try...except httpx.HTTPStatusError` since all calls go through `httpx`.
- Do not instantiate multiple `GraphClient`s for the same account; reuse via `async with` context manager.
- In Jupyter or scripts without an event loop, use `_sync` wrappers like `client.read_excel_sync()`.

### Release Process
Bump `__version__` in `__init__.py`, update the Changelog below, `git tag v0.X.Y`, `git push --tags`.

## Changelog

### [0.1.0] - 2026-05-03

#### Added
- Async Python wrapper around Microsoft Graph API for M365 access.
- `GraphClient` public class with ROPC auth flow and connection management.
- `FileItem` dataclass for SharePoint/OneDrive file representation.
- Read files single/bulk: `read_excel`, `read_csv`, `download`, `read_excel_many`.
- Discovery endpoints: `walk`, `list_files`.
- Write endpoint: `upload`.
- Communication endpoints: `send_mail`, `send_teams_message`.
- Sync wrappers for ease of use in scripts/notebooks: `read_excel_sync`, `read_csv_sync`, `download_sync`.
- Automatic retry logic and concurrency capping per account.
