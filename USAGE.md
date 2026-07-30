# gex-msgraph Usage Guide

## 1. Prerequisites
- Python 3.11+
- `pip` or `uv`
- Git + SSH key for GitHub

## 2. Credentials
Ask IT for a service account. You need: Client ID, Client Secret, Tenant ID, and — for delegated (ROPC) auth — Username and Password. Omit username/password entirely to use app-only (client credentials) auth instead; see the Authentication flows section in the API Reference for the app-only limitations (no `/me/*` endpoints, `default_drive_id` required).

## 3. Install in your project

**Workflow A: pip + venv (Windows)**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

**Workflow B: pip + venv (Ubuntu)**
```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

**Workflow B: uv**
```bash
uv add "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

## 4. Configure `.env`
Create a `.env` file (ensure it's in `.gitignore`!). Secure it with `chmod 600 .env` on Ubuntu or Windows file permissions.
```bash
MS_ACCOUNTS=das_u1

MS_DAS_U1_CLIENT_ID=your_client_id
MS_DAS_U1_CLIENT_SECRET=your_client_secret
MS_DAS_U1_TENANT_ID=your_tenant_id
MS_DAS_U1_USERNAME=service@company.com
MS_DAS_U1_PASSWORD=your_password
```

## 5. Find SharePoint IDs
Use Graph Explorer (`https://developer.microsoft.com/en-us/graph/graph-explorer`) to find your Site ID and Drive ID if needed. Add `MS_DAS_U1_DEFAULT_DRIVE_ID` to `.env`.

## 6. First call
```python
import asyncio
from gex_msgraph import GraphClient

async def main():
    # Option 1: Use credentials automatically loaded from .env
    async with GraphClient("das_u1") as client:
        df = await client.read_excel(item_path="Reports/Q1.xlsx")
        print(df)

    # Option 2: Provide credentials explicitly
    explicit_client = GraphClient(
        client_id="my_client_id",
        client_secret="my_client_secret",
        tenant_id="my_tenant_id",
        username="my_username",
        password="my_password",
        default_drive_id="my_drive_id" # Optional
    )
    async with explicit_client:
        df2 = await explicit_client.read_excel(item_path="Reports/Q2.xlsx")

asyncio.run(main())
```
Or sync:
```python
client = GraphClient("das_u1")
df = client.read_excel_sync(item_path="Reports/Q1.xlsx")
```

## 7. Recipes
- **Read Excel by path:** `await client.read_excel(item_path="Folder/File.xlsx")`
- **Read Excel by share link:** `await client.read_excel(share_url="https://tenant.sharepoint.com/.../File.xlsx")`
- **Read CSV:** `await client.read_csv(item_path="data.csv")`
- **Bulk read Excel:** `await client.read_excel_many(["A.xlsx", "B.xlsx"], on_error="warn")`
- **Bulk read CSV (with status logging):** `df, status_df = await client.read_csv_many(["A.csv", "B.csv"], on_error="skip", return_status=True)`
- **Read Parquet:** `await client.read_parquet(item_path="data.parquet")`
- **Read a SAP-style `.xls` (SpreadsheetML 2003):** `await client.read_excel(item_path="TB/jan.xls")` — detected and decoded automatically, no flag needed.
- **List files:** `await client.list_files("Folder")`
- **Walk recursively with glob:** `await client.walk("Folder", pattern="*.xlsx")`
- **Search files:** `await client.search_files("budget", limit=25)` — `limit` is a hard cap on the returned list, not just a page-size hint.
- **Download raw bytes:** `await client.download(item_path="image.png")`
- **Download large file to disk:** `await client.download(item_path="big.parquet", to_path="big.parquet")`
- **Upload (any size — >4 MiB chunks automatically):** `await client.upload("local.txt", "remote.txt")`
- **Bulk upload:** `await client.upload_many([("local/a.xlsx", "remote/a.xlsx"), ...], max_concurrent=5)`

### File & Folder Management
- **Get File Metadata (Size, Date):** `meta = await client.get_metadata(item_path="File.xlsx")`
- **Delete File:** `await client.delete_file(item_path="File.xlsx")`
- **Copy File:** `await client.copy_file(item_path="File.xlsx", dest_folder_path="Archive", new_name="File_copy.xlsx")`
- **Copy File and wait for result:** `item = await client.copy_file(item_path="File.xlsx", dest_folder_path="Archive", wait=True)`
- **Move/Rename File:** `await client.move_file(item_path="Old.xlsx", dest_folder_path="Archive", new_name="New.xlsx")`
- **Create Folder:** `await client.create_folder("NewFolder")`
- **Check Exists:** `exists = await client.exists(item_path="File.xlsx")`
- **Get Share Link:** `url = await client.get_share_link(item_path="File.xlsx")`
- **Print Folder Tree:**
  ```python
  tree = await client.get_folder_tree("Reports")
  tree.print()
  ```
- **List Excel Sheets (Without downloading):** `sheets = await client.list_excel_sheets(item_path="File.xlsx")`

### Communications
- **List inbox:** `msgs = await client.list_mail(limit=10)`
- **Send mail:** `await client.send_mail("test@test.com", "Subj", "Body")`
- **Send Teams channel:** `await client.send_teams_message("team1", "chan1", "Hello")`
- **Read Teams channel:** `msgs = await client.get_teams_messages("team1", "chan1", limit=10)`
- **List chats:** `chats = await client.list_chats(limit=20)`
- **Get chat messages:** `msgs = await client.get_chat_messages("chat_id", limit=10)`
- **Send chat message:** `await client.send_chat_message("chat_id", "Hello!")`

## 8. Multi-account
Instantiate two clients, e.g. `client1 = GraphClient("das_u1")` and `client2 = GraphClient("das_u2")`.

## 9. FastAPI integration
Use a lifespan context manager to init `client = GraphClient("das_u1")` and pass it to endpoints.

## 10. Prefect integration
Call the async methods inside `@task` functions using `asyncio.run` or use async flows.

## 11. Notebook usage
In Jupyter, `asyncio.run` fails because an event loop is already running. Just `await client.read_excel(...)` directly!

## 12. Deployment to Ubuntu
- **Bare Metal:** `source /opt/myapp/.venv/bin/activate`
- **Docker:** Use `--env-file .env` or mount.
- **systemd:** Use `EnvironmentFile=/etc/myapp/.env`.

## 13. Updating the library
Update the pin to `@v0.2.0` in `requirements.txt` or via `uv add "gex-msgraph @ git+ssh...v0.2.0"`.

## 14. Troubleshooting
| Symptom | Fix |
|---|---|
| KeyError | Env vars not set. If you use a `.env` file, install `python-dotenv` yourself (`pip install python-dotenv` or the `gex-msgraph[dotenv]` extra — it is no longer a hard dependency since v0.3.0) and call `load_dotenv()` before creating the client. |
| `PermissionDeniedError` (403) | Ensure account has SharePoint access. Under app-only auth, check the app has *application* permissions and `default_drive_id` is set. |
| `GraphSyncInLoopError` | You called a `*_sync` method inside Jupyter or an async app — use `await client.method(...)` directly instead. |
| pip git error | Ensure git is installed and SSH agent is running. |
| `CalamineError: Cannot detect file format` / `XLRDError: Expected BOF record` / `ValueError: Excel file format cannot be determined` on a `.xls` | The file is **SpreadsheetML 2003**, not a real `.xls` — plain XML (usually UTF-16) that Excel and SharePoint open happily because of its `<?mso-application progid="Excel.Sheet"?>` hint. Common with SAP exports. Confirm with the magic bytes: `open(f,"rb").read(4)` → `b"\xfe\xff\x00<"` (SpreadsheetML) vs `b"\xd0\xcf\x11\xe0"` (real `.xls`). **Handled automatically since v0.4.0** — `read_excel` / `read_excel_many` detect and decode it. If you hit this, upgrade. |
| `PermissionDeniedError` (403) `Could not obtain a WAC access token` from `list_excel_sheets` | `list_excel_sheets` calls the Graph *workbook* API, which needs a genuine xlsx workbook server-side. It does not work on SpreadsheetML files (nor on `.xls`/`.csv`). Read the file instead and inspect `pd.ExcelFile(...).sheet_names` locally. |

---

## API Reference

Public exports: `from gex_msgraph import GraphClient, FileItem, TreeNode, GraphError, AuthError, PermissionDeniedError, NotFoundError, RateLimitExhaustedError, GraphAuthenticationError, GraphSyncInLoopError`

---

### Exceptions

All Graph HTTP failures raise a **`GraphError`** subclass. `GraphError` itself subclasses `httpx.HTTPStatusError`, so pre-v0.3.0 code that catches `httpx.HTTPStatusError` keeps working unchanged, and `.request` / `.response` remain available.

- **`GraphError`** — base; also raised directly for unmapped 4xx statuses (e.g. 400).
- **`AuthError`** — 401 Unauthorized (token invalid/expired/insufficient).
- **`PermissionDeniedError`** — 403 Forbidden.
- **`NotFoundError`** — 404 Not Found.
- **`RateLimitExhaustedError`** — a 429/5xx persisted through all retry attempts.
- **`GraphAuthenticationError`** (plain `Exception`) — MSAL token acquisition failed before any HTTP call; no request/response attached.
- **`GraphSyncInLoopError`** (`RuntimeError`) — a `*_sync` method was called from inside a running event loop (Jupyter, async frameworks).

```python
from gex_msgraph import GraphClient, NotFoundError, GraphError

async with GraphClient("das_u1") as client:
    try:
        df = await client.read_excel(item_path="Reports/Q1.xlsx")
    except NotFoundError:
        print("File does not exist")
    except GraphError as e:
        print(f"Graph call failed: {e.response.status_code}")
```

---

### `FileItem`

Immutable dataclass representing a file or folder in OneDrive / SharePoint.

**Fields**

- **`name`** (`str`) — File or folder name, without path (e.g. `"Q1.xlsx"`).
- **`path`** (`str`) — Path relative to the drive root (e.g. `"Reports/Q1.xlsx"`).
- **`id`** (`str`) — Microsoft Graph item ID (opaque string).
- **`size`** (`int`) — Size in bytes. `0` for folders.
- **`modified`** (`datetime`) — Last modification time, timezone-aware UTC.
- **`is_folder`** (`bool`) — `True` when the item is a folder.
- **`webUrl`** (`str | None`) — Graph's `driveItem.webUrl` (e.g. a `Doc.aspx` page or a normal SharePoint/OneDrive URL). Use this to build Office Online embed URLs (`?action=embedview`) for an iframe — it's a real item URL, not a sharing link, so it won't be blocked by X-Frame-Options/CSP the way [`get_share_link`](#graphclientget_share_link)'s `/:x:/` links are. `None` when Graph doesn't return the field.

---

### `TreeNode`

Recursive tree node returned by [`get_folder_tree`](#graphclientget_folder_treefolderpath). Rendered with `print()`.

**Fields**

- **`item`** (`FileItem | None`) — Metadata for this node. `None` only for the drive root.
- **`children`** (`list[TreeNode]`) — Direct child nodes (files and sub-folders).

#### `TreeNode.print(indent=0)`

Print the tree structure to stdout with `📁`/`📄` icons.

**Parameters**

- **`indent`** (`int`, default `0`) — Starting indentation level; each level adds two spaces.

**Example**

```python
tree = await client.get_folder_tree("Reports")
tree.print()
# 📁 Reports (0 bytes)
#   📄 Q1.xlsx (45231 bytes)
#   📁 Archive (0 bytes)
#     📄 2025.xlsx (12000 bytes)
```

---

### `GraphClient`

Async client for Microsoft Graph API. Manages authentication (delegated ROPC or app-only client-credentials flow), connection pooling, semaphore-based concurrency, and automatic retry on 429 / 5xx responses (up to 3 attempts, exponential backoff capped at 30 s).

#### Authentication flows

- **Delegated (ROPC)** — provide `username` + `password` along with `client_id`/`client_secret`/`tenant_id`. The client acts as that signed-in user; all methods work.
- **App-only (client credentials)** — omit **both** `username` and `password`. The client acts as the application itself, with *application* permissions instead of delegated ones. **Constraint:** `/me/*` endpoints have no meaning without a signed-in user, so under app-only auth `send_mail`, `list_mail`, `list_chats`, `get_chat_messages`, `send_chat_message` do not work, and all drive operations require `default_drive_id` to be set (the `/me/drive` default cannot resolve). Use app-only for drive-scoped automation on a known SharePoint drive; use ROPC when you need mail/Teams or `/me/drive`.

#### `GraphClient.__init__`

```python
GraphClient(
    account: str | None = None,
    *,
    client_id: str | None = None,
    client_secret: str | None = None,
    tenant_id: str | None = None,
    username: str | None = None,
    password: str | None = None,
    default_drive_id: str | None = None,
    max_concurrent: int | None = None,
    request_timeout: float | None = None,
)
```

**Parameters**

- **`account`** (`str | None`, default `None`) — Env-var prefix. When set to e.g. `"das_u1"`, the following variables are read: `MS_DAS_U1_CLIENT_ID`, `MS_DAS_U1_CLIENT_SECRET`, `MS_DAS_U1_TENANT_ID`, `MS_DAS_U1_USERNAME`, `MS_DAS_U1_PASSWORD`. Keyword arguments below override any env var. When `None`, credentials must be passed explicitly; `account` is recorded as `"custom"`.
- **`client_id`** (`str | None`, default `None`) — Azure app registration client ID. Overrides `MS_<ACCOUNT>_CLIENT_ID`. Required.
- **`client_secret`** (`str | None`, default `None`) — Azure app client secret. Overrides `MS_<ACCOUNT>_CLIENT_SECRET`. Required.
- **`tenant_id`** (`str | None`, default `None`) — Azure tenant (directory) ID. Overrides `MS_<ACCOUNT>_TENANT_ID`. Required.
- **`username`** (`str | None`, default `None`) — Service account UPN (e.g. `svc@company.com`). Overrides `MS_<ACCOUNT>_USERNAME`. Provide together with `password` for delegated (ROPC) auth; omit both for app-only auth.
- **`password`** (`str | None`, default `None`) — Service account password. Overrides `MS_<ACCOUNT>_PASSWORD`. Paired with `username`.
- **`default_drive_id`** (`str | None`, default `None`) — Drive ID. When set, all drive-scoped URLs use `/drives/{id}` instead of `/me/drive`. Readable from `MS_<ACCOUNT>_DEFAULT_DRIVE_ID`. **Required under app-only auth.**
- **`max_concurrent`** (`int | None`, default `10`) — Maximum simultaneous in-flight Graph requests per client instance. Readable from `MS_<ACCOUNT>_MAX_CONCURRENT`.
- **`request_timeout`** (`float | None`, default `30.0`) — Per-request HTTP timeout in seconds. Readable from `MS_<ACCOUNT>_REQUEST_TIMEOUT`.

**Raises**

- **`KeyError`** — If `client_id`, `client_secret`, or `tenant_id` is neither provided explicitly nor found in the environment.
- **`ValueError`** — If only one of `username` / `password` is provided (they must come as a pair, or both be absent for app-only auth), or if app-only auth is used without `default_drive_id`.

**Notes**

Use as an async context manager to ensure the underlying HTTP client is closed properly. For synchronous scripts using the `*_sync` wrappers, call `close_sync()` when done. `default_site_id` was removed in v0.3.0 (it was accepted but never used).

**Example**

```python
# From env vars (delegated)
async with GraphClient("das_u1") as client:
    df = await client.read_excel(item_path="Reports/Q1.xlsx")

# Explicit credentials, app-only (no username/password)
client = GraphClient(
    client_id="...",
    client_secret="...",
    tenant_id="...",
    default_drive_id="b!abc123",  # required for app-only
    max_concurrent=5,
)
```

---

#### `GraphClient.close()` / `GraphClient.close_sync()`

`close()` (async) closes the underlying HTTP client; safe to call multiple times; called automatically when the client is used as an async context manager. `close_sync()` is the synchronous companion for scripts that only use the `*_sync` wrappers — it also stops the background event loop those wrappers run on.

---

### File Operations

---

#### `GraphClient.download`

```python
async def download(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    to_path: str | os.PathLike | None = None,
) -> bytes | Path
```

Download a file. By default returns its raw content as bytes; with `to_path` set, streams the body to disk in chunks (never fully buffered in memory) and returns the written `Path`. Exactly one of the three identifier parameters must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root, e.g. `"Documents/report.pdf"`.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link (full `https://…` URL from "Copy link").
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.
- **`to_path`** (`str | os.PathLike | None`, default `None`) — Local destination file. When given, the download is streamed to this file instead of being returned as bytes. Use for large files.

**Returns**

`bytes` — Raw file content (when `to_path` is `None`).
`Path` — The written destination (when `to_path` is given).

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`RuntimeError`** — If the Graph API response does not contain a `@microsoft.graph.downloadUrl`.
- **`GraphError`** (subclass of `httpx.HTTPStatusError`) — On non-retryable HTTP errors (e.g. `PermissionDeniedError`, `NotFoundError`) or when retries are exhausted (`RateLimitExhaustedError`).

**Example**

```python
data = await client.download(item_path="images/logo.png")
with open("logo.png", "wb") as f:
    f.write(data)

# Large file: stream straight to disk
path = await client.download(item_path="exports/huge.parquet", to_path="huge.parquet")
```

---

#### `GraphClient.read_excel`

```python
async def read_excel(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    sheet: str | int = 0,
    **kwargs,
) -> pd.DataFrame
```

Download an Excel file and parse one sheet into a pandas DataFrame. Exactly one identifier must be provided.

**SpreadsheetML 2003 is handled automatically (v0.4.0+).** Files that carry an `.xls` extension but are really XML Spreadsheet documents (typical of SAP exports — UTF-16 XML, magic bytes `fe ff` instead of `d0 cf 11 e0`) are detected by content and converted to xlsx in memory before parsing. No flag or kwarg is needed, and every kwarg below still applies. Only cell *values* survive the conversion — styles, formulas and merge geometry are dropped, and `ss:Type="DateTime"` cells arrive as strings (pass `parse_dates` or convert yourself). Detection is content-based, so genuine `.xlsx`/`.xls` files are untouched.

Sheet names that xlsx cannot represent (containing `\ / * ? : [ ]`, or longer than 31 characters) are renamed during conversion, with a warning on the `gex_msgraph` logger. Pass `sheet` using the name as it appears in the source document — it is mapped the same way — except for `sheet_match="glob"` in `read_excel_many`, where the pattern is used verbatim so `[…]` keeps its character-class meaning and must therefore target the renamed title.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.
- **`sheet`** (`str | int`, default `0`) — Sheet to read. An integer is a zero-based positional index; a string is the exact sheet name.
- **`**kwargs`** — Forwarded verbatim to `pandas.read_excel` (e.g. `header`, `usecols`, `dtype`, `skiprows`). **Default engine:** `"calamine"` (fast Rust-based parser, hard dep). Override with `engine="openpyxl"` if needed.

**Returns**

`pd.DataFrame` — Contents of the requested sheet.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`RuntimeError`** — If the Graph API response does not contain a download URL.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
# By path, second sheet (positional)
df = await client.read_excel(item_path="Finance/budget.xlsx", sheet=1)

# By share link, named sheet, subset of columns
df = await client.read_excel(
    share_url="https://tenant.sharepoint.com/:x:/r/...",
    sheet="Q1 Data",
    usecols="A:F",
    dtype={"Amount": float},
)
```

---

#### `GraphClient.read_csv`

```python
async def read_csv(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    **kwargs,
) -> pd.DataFrame
```

Download a CSV file and parse it into a pandas DataFrame. Exactly one identifier must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.
- **`**kwargs`** — Forwarded verbatim to `pandas.read_csv` (e.g. `sep`, `encoding`, `dtype`, `parse_dates`).

**Returns**

`pd.DataFrame` — Parsed CSV content.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`RuntimeError`** — If the Graph API response does not contain a download URL.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
df = await client.read_csv(item_path="data/users.csv", sep=";", encoding="utf-8-sig")
```

---

#### `GraphClient.read_parquet`

```python
async def read_parquet(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    **kwargs,
) -> pd.DataFrame
```

Download a Parquet file and parse it into a pandas DataFrame. Exactly one identifier must be provided. Extra keyword arguments are forwarded to `pandas.read_parquet` (requires `pyarrow` or `fastparquet` in your project).

**Example**

```python
df = await client.read_parquet(item_path="exports/facts.parquet", columns=["id", "amount"])
```

---

#### `GraphClient.read_excel_many`

```python
async def read_excel_many(
    paths: list[str],
    *,
    sheet: str | int = 0,
    sheet_match: Literal["exact", "ci", "glob"] = "exact",
    on_missing_sheet: Literal["raise", "skip", "warn"] = "raise",
    on_error: Literal["raise", "skip", "warn"] = "raise",
    add_source_column: bool = True,
    max_concurrent: int | None = None,
    return_status: bool = False,
    **kwargs,
) -> pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]
```

Read multiple Excel files concurrently and concatenate them into one DataFrame.

Like `read_excel`, this transparently decodes **SpreadsheetML 2003** files (v0.4.0+) — see the note under `read_excel` for the renaming/caveats. `sheet`, `sheet_match` and `on_missing_sheet` all resolve against the source document's own sheet names (remapped internally when a title had to be sanitized), so they behave the same as on a native xlsx file.

**Parameters**

- **`paths`** (`list[str]`) — Item paths relative to the drive root to read.
- **`sheet`** (`str | int`, default `0`) — Sheet selector applied to every file. An integer is a zero-based positional index; a string is matched using `sheet_match`.
- **`sheet_match`** (`Literal["exact", "ci", "glob"]`, default `"exact"`) — Matching strategy for a string `sheet`:
  - `"exact"` — case-sensitive equality.
  - `"ci"` — case-insensitive equality.
  - `"glob"` — `fnmatch` wildcard pattern (e.g. `"Sales*"`).
- **`on_missing_sheet`** (`Literal["raise", "skip", "warn"]`, default `"raise"`) — Behaviour when the requested sheet is absent from a workbook:
  - `"raise"` — raises `ValueError` immediately.
  - `"skip"` — silently omits the file from the result.
  - `"warn"` — logs a warning and omits the file.
- **`on_error`** (`Literal["raise", "skip", "warn"]`, default `"raise"`) — Behaviour when a file fails to download or parse:
  - `"raise"` — re-raises the exception immediately.
  - `"skip"` — silently omits the file.
  - `"warn"` — logs a warning and omits the file.
- **`add_source_column`** (`bool`, default `True`) — Append a `_source` column to each file's rows containing its originating path.
- **`max_concurrent`** (`int | None`, default `None`) — Additional concurrency cap for this call only. `None` applies no extra limit beyond the client-level semaphore.
- **`return_status`** (`bool`, default `False`) — When `True`, return a two-element tuple instead of a single DataFrame.
- **`**kwargs`** — Forwarded verbatim to `pandas.read_excel` (e.g. `usecols="A:M"`, `header`, `skiprows`, `dtype`, `nrows`). The `engine` kwarg is honored and applied to the underlying `pd.ExcelFile`. **Default engine:** `"calamine"` (fast Rust-based parser, hard dep). Override with `engine="openpyxl"` if needed.

**Returns**

- `pd.DataFrame` — Concatenated data from all successfully read files (`return_status=False`).
- `tuple[pd.DataFrame, pd.DataFrame]` — `(combined_df, status_df)` when `return_status=True`. `status_df` has columns:
  - `path` — file path.
  - `status` — `"success"`, `"missing_sheet"`, or `"error"`.
  - `error` — error message string, or empty string on success.

**Raises**

- **`ValueError`** — If `on_missing_sheet="raise"` and a sheet is not found.
- **`Exception`** — Re-raises the download / parse exception if `on_error="raise"`.

**Example**

```python
# Strict: any failure raises immediately
df = await client.read_excel_many(["jan.xlsx", "feb.xlsx"], sheet="Sales")

# Tolerant: warn and continue, then inspect which files failed
df, status = await client.read_excel_many(
    ["jan.xlsx", "feb.xlsx", "mar.xlsx"],
    sheet="Sales*",
    sheet_match="glob",
    on_missing_sheet="warn",
    on_error="warn",
    return_status=True,
)
print(status[status["status"] != "success"])

# Faster parsing with calamine engine + column subset
df = await client.read_excel_many(
    ["jan.xlsx", "feb.xlsx"],
    sheet="Sales",
    usecols="A:M",
    engine="calamine",
)
```

---

#### `GraphClient.read_csv_many`

```python
async def read_csv_many(
    paths: list[str],
    *,
    on_error: Literal["raise", "skip", "warn"] = "raise",
    add_source_column: bool = True,
    max_concurrent: int | None = None,
    return_status: bool = False,
    **kwargs,
) -> pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]
```

Read multiple CSV files concurrently and concatenate them into one DataFrame.

**Parameters**

- **`paths`** (`list[str]`) — Item paths relative to the drive root to read.
- **`on_error`** (`Literal["raise", "skip", "warn"]`, default `"raise"`) — Behaviour when a file fails to download or parse:
  - `"raise"` — re-raises the exception immediately.
  - `"skip"` — silently omits the file.
  - `"warn"` — logs a warning and omits the file.
- **`add_source_column`** (`bool`, default `True`) — Append a `_source` column containing the originating path.
- **`max_concurrent`** (`int | None`, default `None`) — Additional concurrency cap for this call only.
- **`return_status`** (`bool`, default `False`) — When `True`, return `(combined_df, status_df)`. `status_df` columns: `path`, `status` (`"success"` | `"error"`), `error`.
- **`**kwargs`** — Forwarded verbatim to `pandas.read_csv`.

**Returns**

- `pd.DataFrame` — Concatenated data (`return_status=False`).
- `tuple[pd.DataFrame, pd.DataFrame]` — `(combined_df, status_df)` when `return_status=True`.

**Raises**

- **`Exception`** — Re-raises the download / parse exception if `on_error="raise"`.

**Example**

```python
df, status = await client.read_csv_many(
    ["exports/jan.csv", "exports/feb.csv"],
    on_error="warn",
    return_status=True,
    sep=";",
)
failed = status[status["status"] == "error"]
```

---

#### `GraphClient.list_excel_sheets`

```python
async def list_excel_sheets(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
) -> list[str]
```

Return the worksheet names of an Excel file without downloading its full content. Uses the Graph workbook API to read sheet metadata only.

Exactly one identifier must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.

**Returns**

`list[str]` — Worksheet names in workbook order.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
sheets = await client.list_excel_sheets(item_path="Finance/budget.xlsx")
# ['Summary', 'Q1', 'Q2', 'Q3', 'Q4']
```

---

### File & Folder Discovery

---

#### `GraphClient.walk`

```python
async def walk(
    folder_path: str = "",
    *,
    pattern: str | None = None,
    recursive: bool = True,
) -> list[FileItem]
```

List files under a folder. Folders are not included in the result.

**Parameters**

- **`folder_path`** (`str`, default `""`) — Starting folder path relative to the drive root. Empty string means the drive root.
- **`pattern`** (`str | None`, default `None`) — `fnmatch` glob applied to file names only (not full paths), e.g. `"*.xlsx"`. `None` returns all files.
- **`recursive`** (`bool`, default `True`) — When `True`, descend into sub-folders concurrently. When `False`, only immediate children are scanned.

**Returns**

`list[FileItem]` — Matching files. Order is not guaranteed when `recursive=True`.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
# All Excel files under "Reports", recursively
files = await client.walk("Reports", pattern="*.xlsx")

# Non-recursive — immediate children only
files = await client.walk("Inbox", recursive=False)
```

---

#### `GraphClient.list_files`

```python
async def list_files(folder_path: str = "") -> list[FileItem]
```

List the immediate children (files *and* folders) of a single folder. Non-recursive.

**Parameters**

- **`folder_path`** (`str`, default `""`) — Folder path relative to the drive root. Empty string means the drive root.

**Returns**

`list[FileItem]` — Direct children in the order returned by the API. Use `item.is_folder` to distinguish files from folders.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
items   = await client.list_files("Projects")
files   = [i for i in items if not i.is_folder]
folders = [i for i in items if i.is_folder]
```

---

#### `GraphClient.get_folder_tree`

```python
async def get_folder_tree(folder_path: str = "") -> TreeNode
```

Build and return a recursive `TreeNode` tree for a folder and all its contents. Internally calls `list_files` depth-first.

**Parameters**

- **`folder_path`** (`str`, default `""`) — Folder path relative to the drive root. Empty string means the drive root; the root node's `item` field will be `None`.

**Returns**

`TreeNode` — Root node of the tree. Call `.print()` for a console view.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
tree = await client.get_folder_tree("Reports")
tree.print()
# 📁 Reports (0 bytes)
#   📄 Q1.xlsx (45231 bytes)
```

---

#### `GraphClient.get_metadata`

```python
async def get_metadata(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
) -> FileItem
```

Fetch metadata for a single file or folder without downloading its content. Exactly one identifier must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.

**Returns**

`FileItem` — Metadata for the item.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 404 if the path does not exist).

**Example**

```python
meta = await client.get_metadata(item_path="Reports/Q1.xlsx")
print(f"{meta.size} bytes, last modified {meta.modified:%Y-%m-%d}")
```

---

#### `GraphClient.exists`

```python
async def exists(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
) -> bool
```

Check whether an item exists without raising on 404. Exactly one identifier must be provided.

**Returns**

`bool` — `True` if the item exists, `False` if Graph returns 404.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`GraphError`** — On any HTTP error other than 404 (e.g. `PermissionDeniedError`).

**Example**

```python
if not await client.exists(item_path="Reports/Q1.xlsx"):
    print("missing!")
```

---

#### `GraphClient.search_files`

```python
async def search_files(
    query: str,
    *,
    limit: int = 25,
) -> list[FileItem]
```

Search files across the drive by name or content, using Graph's drive search. `limit` is a hard cap on the number of returned items (pagination stops once reached); `limit <= 0` returns `[]` without a request.

**Parameters**

- **`query`** (`str`) — Search text.
- **`limit`** (`int`, default `25`) — Maximum number of results returned.

**Returns**

`list[FileItem]` — Matching items (files and folders), at most `limit`.

**Example**

```python
hits = await client.search_files("budget", limit=10)
for f in hits:
    print(f.path)
```

---

### File & Folder Management

---

#### `GraphClient.upload`

```python
async def upload(
    local_path: str | os.PathLike,
    remote_path: str,
) -> dict
```

Upload a local file to OneDrive / SharePoint. If the remote file already exists it is overwritten. Files up to ~4 MiB use a single PUT; larger files are transparently uploaded via a Graph upload session in 10 MiB chunks, so there is no practical size limit.

**Parameters**

- **`local_path`** (`str | os.PathLike`) — Absolute or relative path to the local source file.
- **`remote_path`** (`str`) — Destination path relative to the drive root, including file name (e.g. `"Uploads/report.xlsx"`).

**Returns**

`dict` — Raw Graph API `driveItem` response. Useful keys: `id`, `name`, `size`, `webUrl`.

**Raises**

- **`FileNotFoundError`** — If `local_path` does not exist or is not a file.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
result = await client.upload("./output/report.xlsx", "Reports/2026/report.xlsx")
print(result["webUrl"])
```

---

#### `GraphClient.upload_many`

```python
async def upload_many(
    items: list[tuple[str | os.PathLike, str]],
    *,
    on_error: Literal["raise", "skip", "warn"] = "raise",
    max_concurrent: int | None = None,
    return_status: bool = False,
) -> list[dict] | tuple[list[dict], pd.DataFrame]
```

Upload multiple local files concurrently. Each item is a `(local_path, remote_path)` tuple.

**Parameters**

- **`items`** — List of `(local_path, remote_path)` pairs.
- **`on_error`** (default `"raise"`) — `"raise"` aborts on first failure; `"skip"` continues silently; `"warn"` continues and logs a warning.
- **`max_concurrent`** (`int | None`, default `None`) — Optional bound on concurrent uploads (added in v0.2.1; unbounded when `None`, subject to the client-wide semaphore).
- **`return_status`** (default `False`) — When `True`, also return a status DataFrame with `path` / `status` / `error` columns.

**Returns**

`list[dict]` — Graph `driveItem` dicts for successful uploads; or `(results, status_df)` when `return_status=True`.

**Example**

```python
results, status = await client.upload_many(
    [("./a.xlsx", "Reports/a.xlsx"), ("./b.xlsx", "Reports/b.xlsx")],
    on_error="warn",
    return_status=True,
)
```

---

#### `GraphClient.delete_file`

```python
async def delete_file(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
) -> None
```

Delete a file or folder. The item is moved to the OneDrive / SharePoint recycle bin. Exactly one identifier must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.

**Returns**

`None`

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 404 if the item does not exist).

**Example**

```python
await client.delete_file(item_path="temp/scratch.xlsx")
```

---

#### `GraphClient.move_file`

```python
async def move_file(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    dest_folder_path: str | None = None,
    new_name: str | None = None,
) -> FileItem
```

Move and/or rename a file in a single API call. The source accepts any of the three identifier kinds (exactly one); the destination is always a folder path. At least one of `dest_folder_path` or `new_name` must be provided.

> **Breaking change in v0.3.0**: the positional `source_path` parameter was replaced by the keyword-only `item_path` / `share_url` / `item_id` trio, matching every other method. Migrate `move_file("a.xlsx", ...)` → `move_file(item_path="a.xlsx", ...)`.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Current file path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link identifying the source item.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID of the source item.
- **`dest_folder_path`** (`str | None`, default `None`) — Target folder path relative to the drive root. Pass `""` to move to the drive root. `None` keeps the current parent folder.
- **`new_name`** (`str | None`, default `None`) — New file name including extension. `None` keeps the current name.

**Returns**

`FileItem` — Metadata of the item at its new location.

**Raises**

- **`ValueError`** — If not exactly one source identifier is provided, or if neither `dest_folder_path` nor `new_name` is provided.
- **`GraphError`** — On non-retryable HTTP errors.

**Example**

```python
# Move and rename in one call
item = await client.move_file(
    item_path="Drafts/report_v2.xlsx",
    dest_folder_path="Published",
    new_name="report_final.xlsx",
)
print(item.path)  # "Published/report_final.xlsx"

# Rename only
await client.move_file(item_path="Reports/old_name.xlsx", new_name="new_name.xlsx")

# Move by item id
await client.move_file(item_id="01ABC123", dest_folder_path="Archive")
```

---

#### `GraphClient.copy_file`

```python
async def copy_file(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    dest_folder_path: str,
    new_name: str | None = None,
    wait: bool = False,
    wait_timeout: float = 60.0,
) -> FileItem | None
```

Copy a file to a new location. Graph executes the copy as an asynchronous job. With `wait=False` (default) the method returns `None` immediately after the job is accepted; with `wait=True` it polls the job's monitor URL until completion and returns the new item's `FileItem`.

> **Breaking change in v0.3.0**: the positional `source_path` parameter was replaced by the keyword-only `item_path` / `share_url` / `item_id` trio. Migrate `copy_file("a.xlsx", "Archive")` → `copy_file(item_path="a.xlsx", dest_folder_path="Archive")`.

**Parameters**

- **`item_path`** / **`share_url`** / **`item_id`** — Source item; exactly one must be provided.
- **`dest_folder_path`** (`str`) — Target folder path relative to the drive root. Pass `""` for the drive root.
- **`new_name`** (`str | None`, default `None`) — Name for the copy. `None` keeps the source name.
- **`wait`** (`bool`, default `False`) — When `True`, poll until the copy job completes and return the new `FileItem`.
- **`wait_timeout`** (`float`, default `60.0`) — Maximum seconds to poll when `wait=True`.

**Returns**

`FileItem | None` — `None` when `wait=False`; the new item's metadata when `wait=True`.

**Raises**

- **`ValueError`** — If not exactly one source identifier is provided.
- **`GraphError`** — If the copy job reports `failed` (when `wait=True`), or on non-retryable HTTP errors.
- **`TimeoutError`** — If the job does not complete within `wait_timeout` (when `wait=True`).

**Example**

```python
# Fire-and-forget
await client.copy_file(item_path="Reports/Q1.xlsx", dest_folder_path="Archive")

# Wait for the copy and get the new item
item = await client.copy_file(
    item_path="Reports/Q1.xlsx",
    dest_folder_path="Archive",
    new_name="Q1_backup.xlsx",
    wait=True,
)
print(item.id, item.webUrl)
```

---

#### `GraphClient.create_folder`

```python
async def create_folder(folder_path: str) -> FileItem
```

Create a new folder. Intermediate parent folders must already exist. If a folder with the same name already exists at the destination, Graph will auto-rename the new one (e.g. `Q1 (1)`).

**Parameters**

- **`folder_path`** (`str`) — Path of the folder to create, relative to the drive root (e.g. `"Reports/2026/Q1"`).

**Returns**

`FileItem` — Metadata of the newly created folder.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 404 if the parent folder does not exist).

**Example**

```python
folder = await client.create_folder("Reports/2026/Q1")
print(folder.id)
```

---

#### `GraphClient.get_share_link`

```python
async def get_share_link(
    *,
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
    link_type: Literal["view", "edit"] = "view",
    scope: Literal["anonymous", "organization"] = "organization",
) -> str
```

Create a sharing link for an item and return its `webUrl`. Exactly one identifier must be provided.

> Sharing links (`/:x:/…`) are for people to open in a browser. They are **blocked inside iframes** by SharePoint — for Office Online embedding use `FileItem.webUrl` instead.

**Parameters**

- **`item_path`** / **`share_url`** / **`item_id`** — Target item; exactly one.
- **`link_type`** (default `"view"`) — `"view"` (read-only) or `"edit"`.
- **`scope`** (default `"organization"`) — `"organization"` (tenant-only) or `"anonymous"` (anyone with the link, if tenant policy allows).

**Returns**

`str` — The sharing link URL.

**Example**

```python
url = await client.get_share_link(item_path="Reports/Q1.xlsx", link_type="edit")
```

---

### Communication

---

#### `GraphClient.list_mail`

```python
async def list_mail(
    limit: int = 20,
    *,
    folder: str = "inbox",
) -> list[dict]
```

List messages from a mail folder. Requires delegated (ROPC) auth — this is a `/me/*` endpoint.

**Parameters**

- **`limit`** (`int`, default `20`) — Maximum number of messages.
- **`folder`** (`str`, default `"inbox"`) — Well-known folder name (`"inbox"`, `"sentitems"`, `"drafts"`, …) or a folder ID.

**Returns**

`list[dict]` — Message dicts with `id`, `subject`, `from`, `receivedDateTime`, `bodyPreview`, `hasAttachments`.

**Example**

```python
for msg in await client.list_mail(10):
    print(msg["receivedDateTime"], msg["subject"])
```

---

#### `GraphClient.send_mail`

```python
async def send_mail(
    to: str | list[str],
    subject: str,
    body: str,
    *,
    body_type: Literal["text", "html"] = "text",
    cc: str | list[str] | None = None,
    attachments: list[str | os.PathLike | tuple[str, bytes] | dict[str, str]] | None = None,
) -> None
```

Send an email from the authenticated account's mailbox. Supports plain-text or HTML body and file attachments. The sent message is saved to Sent Items automatically.

**Parameters**

- **`to`** (`str | list[str]`) — Recipient email address or list of addresses.
- **`subject`** (`str`) — Email subject line.
- **`body`** (`str`) — Message body (plain text or HTML depending on `body_type`).
- **`body_type`** (`Literal["text", "html"]`, default `"text"`) — Content type of `body`. Use `"html"` to send rich HTML email.
- **`cc`** (`str | list[str] | None`, default `None`) — CC recipient(s). `None` sends no CC.
- **`attachments`** (`list | None`, default `None`) — Files to attach. Each item is one of:
  - `str | os.PathLike` — path to a local file; filename inferred from the path.
  - `(filename, bytes)` tuple — attach in-memory content (e.g. a DataFrame exported to CSV).
  - `dict` with exactly one of `item_path`, `share_url`, or `item_id` — fetched from SharePoint/OneDrive via `download()`.

**Returns**

`None`

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 403 if the account lacks `Mail.Send` permission).

**Example**

```python
# Plain text
await client.send_mail(
    to=["analyst@company.com", "manager@company.com"],
    subject="Daily ETL complete",
    body="The pipeline finished successfully.",
    cc="lead@company.com",
)

# HTML body + mixed attachments
await client.send_mail(
    to="analyst@company.com",
    subject="ETL report",
    body="<h2>Done</h2><p>Rows processed: <b>1,204</b></p>",
    body_type="html",
    attachments=[
        "./output/summary.xlsx",                      # local file
        ("log.csv", log_df.to_csv().encode()),         # in-memory bytes
        {"item_path": "Reports/Q1.xlsx"},             # SharePoint file
    ],
)
```

---

#### `GraphClient.send_teams_message`

```python
async def send_teams_message(
    team_id: str,
    channel_id: str,
    text: str,
) -> None
```

Post a plain-text message to a Teams channel.

**Parameters**

- **`team_id`** (`str`) — The Teams group ID (visible in Graph Explorer or in the Teams channel URL).
- **`channel_id`** (`str`) — The channel ID within the team.
- **`text`** (`str`) — Message body (plain text).

**Returns**

`None`

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
await client.send_teams_message(
    team_id="19:abc...@thread.tacv2",
    channel_id="19:xyz...@thread.tacv2",
    text="Deployment finished successfully.",
)
```

---

#### `GraphClient.get_teams_messages`

```python
async def get_teams_messages(
    team_id: str,
    channel_id: str,
    limit: int = 10,
) -> list[dict]
```

Fetch the most recent messages from a Teams channel.

**Parameters**

- **`team_id`** (`str`) — The Teams group ID.
- **`channel_id`** (`str`) — The channel ID within the team.
- **`limit`** (`int`, default `10`) — Maximum number of messages to return.

**Returns**

`list[dict]` — Graph `chatMessage` objects in reverse-chronological order. Commonly used fields: `id`, `body.content`, `from.user.displayName`, `createdDateTime`.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
messages = await client.get_teams_messages(
    "19:abc...@thread.tacv2",
    "19:xyz...@thread.tacv2",
    limit=5,
)
for msg in messages:
    print(msg["from"]["user"]["displayName"], ":", msg["body"]["content"])
```

---

#### `GraphClient.list_chats`

```python
async def list_chats(limit: int = 50) -> list[dict]
```

List all chats (1-1, group, meeting) the authenticated account participates in, with member details expanded.

**Parameters**

- **`limit`** (`int`, default `50`) — Maximum number of chats to return.

**Returns**

`list[dict]` — Graph `chat` objects. Commonly used fields:

| Field | Type | Description |
|---|---|---|
| `id` | `str` | Chat ID — pass to `get_chat_messages` / `send_chat_message`. |
| `chatType` | `str` | `"oneOnOne"`, `"group"`, or `"meeting"`. |
| `topic` | `str \| None` | Chat display name (group chats only). |
| `members` | `list[dict]` | Members with `displayName` and `userId`. |

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
chats = await client.list_chats()
for chat in chats:
    members = [m["displayName"] for m in chat.get("members", [])]
    print(chat["chatType"], chat["id"], members)
```

---

#### `GraphClient.get_chat_messages`

```python
async def get_chat_messages(
    chat_id: str,
    limit: int = 10,
) -> list[dict]
```

Fetch the most recent messages from a Teams 1-1 or group chat.

**Parameters**

- **`chat_id`** (`str`) — Chat ID obtained from `list_chats()`.
- **`limit`** (`int`, default `10`) — Maximum number of messages to return.

**Returns**

`list[dict]` — Graph `chatMessage` objects in reverse-chronological order. Commonly used fields: `id`, `body.content`, `from.user.displayName`, `createdDateTime`.

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
chats = await client.list_chats()
chat_id = chats[0]["id"]

messages = await client.get_chat_messages(chat_id, limit=20)
for msg in messages:
    print(msg["createdDateTime"], msg["from"]["user"]["displayName"], msg["body"]["content"])
```

---

#### `GraphClient.send_chat_message`

```python
async def send_chat_message(chat_id: str, text: str) -> None
```

Post a plain-text message to a Teams 1-1 or group chat.

**Parameters**

- **`chat_id`** (`str`) — Chat ID obtained from `list_chats()`.
- **`text`** (`str`) — Message body (plain text).

**Returns**

`None`

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
await client.send_chat_message(chat_id, "ETL finished. Check the report on SharePoint.")
```

---

### Synchronous Helpers

Convenience wrappers for scripts, Prefect tasks, or any context without a running event loop. Since v0.3.0 they run on a dedicated background event loop (created lazily on the first sync call and reused for the client's lifetime), so repeated sync calls on the same client are safe. Call `close_sync()` when done.

> **Note:** Calling these from within an `async def` function or inside Jupyter raises `GraphSyncInLoopError` — use the native `await` syntax there instead. Pick one usage mode per client instance: either async (`async with` / `await`) or sync (`*_sync` + `close_sync()`), not both.

---

#### `GraphClient.read_excel_sync(**kwargs)` → `pd.DataFrame`

Synchronous wrapper for [`read_excel`](#graphclientread_excel). Accepts all the same keyword arguments.

```python
df = client.read_excel_sync(item_path="Reports/Q1.xlsx", sheet="Summary")
```

---

#### `GraphClient.read_csv_sync(**kwargs)` → `pd.DataFrame`

Synchronous wrapper for [`read_csv`](#graphclientread_csv). Accepts all the same keyword arguments.

```python
df = client.read_csv_sync(item_path="data/users.csv", sep=";")
```

---

#### `GraphClient.download_sync(**kwargs)` → `bytes | Path`

Synchronous wrapper for [`download`](#graphclientdownload). Accepts all the same keyword arguments, including `to_path`.

```python
data = client.download_sync(item_path="archive/data.bin")
```

---

#### `GraphClient.read_excel_many_sync(paths, **kwargs)` → `pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]`

Synchronous wrapper for [`read_excel_many`](#graphclientread_excel_many). Accepts all the same arguments.

```python
df, status = client.read_excel_many_sync(["a.xlsx", "b.xlsx"], on_error="warn", return_status=True)
```

---

#### `GraphClient.read_csv_many_sync(paths, **kwargs)` → `pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]`

Synchronous wrapper for [`read_csv_many`](#graphclientread_csv_many). Accepts all the same arguments.

```python
df = client.read_csv_many_sync(["a.csv", "b.csv"], on_error="skip")
```
