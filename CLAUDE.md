# gex-msgraph — CLAUDE.md

Async Python wrapper for Microsoft Graph API (M365: OneDrive, SharePoint, Outlook, Teams).
Full API reference → [USAGE.md](USAGE.md).

## Commands

```bash
# Run tests
pytest

# Lint + type-check
ruff check src tests
mypy src
```

## Architecture

```
src/gex_msgraph/
├── _core.py    # GraphClient + _TokenProvider + retry logic
├── _files.py   # FileItem, TreeNode, URL resolution, sheet matching
└── __init__.py # Public exports: GraphClient, FileItem
```

## Invariants — never break these

- **All Graph HTTP calls MUST go through `_request()`** — it handles auth header injection, semaphore, and retry.
- **Never log credentials** — tokens, passwords, secrets. Logger is `logging.getLogger("gex_msgraph")`.
- **Do not add dependencies** without discussion (`httpx`, `msal`, `pandas`, `openpyxl`, `python-dotenv` only).
- **`_TokenProvider.get_token()` is synchronous** — always call via `asyncio.to_thread()` inside async context.
- Retry only on **429 and 5xx**. Never retry 4xx (auth/permission errors should surface immediately).

## Common patterns

```python
# Identifier: exactly one of item_path / share_url / item_id
await client.download(item_path="Reports/Q1.xlsx")
await client.download(share_url="https://tenant.sharepoint.com/:x:/r/...")
await client.download(item_id="01ABC123...")

# Bulk read with error tolerance
df, status = await client.read_excel_many(
    paths, sheet="Data", on_error="warn", return_status=True
)

# Drive root: /me/drive (default) or /drives/{id} when default_drive_id is set
# _drive_root() returns the correct fragment — use it in all new drive-scoped URLs
```

## Adding a new method

1. Implement in `_core.py` — use `self._request()` for all HTTP calls
2. Add test in `tests/test_client.py` using `respx` to mock Graph responses
3. Add recipe (one-liner) to **Recipes** section in `USAGE.md`
4. Add full entry (signature + Parameters + Returns + Raises + Example) to **API Reference** in `USAGE.md`

## Adding a new identifier type

Touches `_files.py` only: `validate_identifier` → `build_resolution_url`.

## Env var convention

`MS_<ACCOUNT>_<KEY>` — e.g. `MS_DAS_U1_CLIENT_ID`. `_load_account_env(name)` reads them.
Optional overrides: `DEFAULT_DRIVE_ID`, `MAX_CONCURRENT`, `REQUEST_TIMEOUT`.

## Tests

Fixtures in `tests/conftest.py`: `mock_token` patches `_TokenProvider.get_token`, `mock_graph` mounts a `respx` router at `https://graph.microsoft.com/v1.0`.

```python
# Typical test skeleton
async def test_something(mock_token, mock_graph):
    mock_graph.get("/me/drive/root:/file.xlsx").mock(return_value=httpx.Response(200, json={...}))
    async with GraphClient("das_u1") as client:
        result = await client.some_method(item_path="file.xlsx")
    assert result == expected
```
