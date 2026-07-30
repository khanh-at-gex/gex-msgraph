# gex-msgraph — CLAUDE.md

Async Python wrapper for Microsoft Graph API (M365: OneDrive, SharePoint, Outlook, Teams).
Full API reference → [USAGE.md](USAGE.md).

**Companion files**
- `AGENTS.md` — user-facing quick reference for junior devs (Gemini, Codex, etc.). Keep in sync when adding public methods.
- `USAGE.md` — full API reference (signatures, Parameters, Returns, Raises, Examples).

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
├── _core.py           # GraphClient + _TokenProvider + retry logic
├── _files.py          # FileItem, TreeNode, URL resolution, sheet matching
├── _exceptions.py     # GraphError hierarchy (subclasses httpx.HTTPStatusError)
├── _spreadsheetml.py  # SpreadsheetML 2003 (SAP ".xls") → xlsx bytes
└── __init__.py        # Public exports: GraphClient, FileItem, exceptions
```

`_spreadsheetml.py` is pure byte-in/byte-out — no HTTP, no client state. `read_excel`
and `read_excel_many` funnel the downloaded bytes through the single entry point
`prepare_excel_bytes(data, sheet)` before handing them to pandas; it returns the bytes
unchanged for anything that is not SpreadsheetML. Route any new Excel-reading method
through it too rather than calling `looks_like_spreadsheetml`/`to_xlsx_bytes` directly —
one entry point is what stops the call sites from drifting apart.

Detection is by content (BOM + namespace), never by extension. Conversion renames sheets
whose titles xlsx cannot hold, so the caller's `sheet` name is remapped the same way —
except for `sheet_match="glob"`, where `prepare_excel_bytes(..., sanitize_sheet_name=False)`
keeps `[…]` meaningful as pattern syntax.

## Invariants — never break these

- **All Graph HTTP calls MUST go through `_request()`** — it handles auth header injection, semaphore, and retry.
  - Exception: `_stream_to_bytes` (used by `download`) fetches Graph's pre-authenticated `downloadUrl`, which must NOT get a Bearer header, so it can't go through `_request()`. It keeps its own 429/5xx retry loop reusing `_compute_backoff`/`_DEFAULT_MAX_RETRIES` — keep both retry loops in sync if backoff behavior changes.
- **Never log credentials** — tokens, passwords, secrets. Logger is `logging.getLogger("gex_msgraph")`.
- **Do not add dependencies** without discussion (`httpx`, `msal`, `pandas`, `openpyxl`, `python-calamine` only; `python-dotenv` is an optional extra, never import it in `src/`).
- **`_TokenProvider.get_token()` is synchronous** — always call via `asyncio.to_thread()` inside async context.
- Retry only on **429 and 5xx**. Never retry 4xx (auth/permission errors should surface immediately).
- **Raise `GraphError` subclasses** (via `raise_graph_error` in `_exceptions.py`) for HTTP failures, never bare `raise_for_status()`. `GraphError` must keep subclassing `httpx.HTTPStatusError` — consumer code depends on it.
- The chunked-upload PUTs (`_put_chunk_with_retry`) and `_stream_to_path` share `_stream_to_bytes`'s no-Bearer/retry contract — keep all three in sync with `_request`'s backoff behavior.

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
5. Add a snippet to `AGENTS.md` under the relevant section
6. If the method returns a new type, export it from `__init__.py`

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
