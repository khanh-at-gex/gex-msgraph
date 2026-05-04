# gex-msgraph Usage Guide

## 1. Prerequisites
- Python 3.11+
- `pip` or `uv`
- Git + SSH key for GitHub

## 2. Credentials
Ask IT for a service account. You need: Client ID, Client Secret, Tenant ID, Username, Password.

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
- **List files:** `await client.list_files("Folder")`
- **Walk recursively with glob:** `await client.walk("Folder", pattern="*.xlsx")`
- **Download raw bytes:** `await client.download(item_path="image.png")`
- **Upload:** `await client.upload("local.txt", "remote.txt")`

### File & Folder Management
- **Get File Metadata (Size, Date):** `meta = await client.get_metadata(item_path="File.xlsx")`
- **Delete File:** `await client.delete_file(item_path="File.xlsx")`
- **Move/Rename File:** `await client.move_file("Old.xlsx", dest_folder_path="Archive", new_name="New.xlsx")`
- **Create Folder:** `await client.create_folder("NewFolder")`
- **Print Folder Tree:**
  ```python
  tree = await client.get_folder_tree("Reports")
  tree.print()
  ```
- **List Excel Sheets (Without downloading):** `sheets = await client.list_excel_sheets(item_path="File.xlsx")`

### Communications
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
| KeyError | Check `.env` is loaded via `python-dotenv`. |
| 403 Forbidden | Ensure account has SharePoint access. |
| pip git error | Ensure git is installed and SSH agent is running. |

---

## API Reference

Public exports: `from gex_msgraph import GraphClient, FileItem, TreeNode`

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

Async client for Microsoft Graph API. Manages authentication (MSAL ROPC flow), connection pooling, semaphore-based concurrency, and automatic retry on 429 / 5xx responses (up to 3 attempts, exponential backoff capped at 30 s).

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
    default_site_id: str | None = None,
    default_drive_id: str | None = None,
    max_concurrent: int | None = None,
    request_timeout: float | None = None,
)
```

**Parameters**

- **`account`** (`str | None`, default `None`) — Env-var prefix. When set to e.g. `"das_u1"`, the following variables are read: `MS_DAS_U1_CLIENT_ID`, `MS_DAS_U1_CLIENT_SECRET`, `MS_DAS_U1_TENANT_ID`, `MS_DAS_U1_USERNAME`, `MS_DAS_U1_PASSWORD`. Keyword arguments below override any env var. When `None`, all five credentials must be passed explicitly; `account` is recorded as `"custom"`.
- **`client_id`** (`str | None`, default `None`) — Azure app registration client ID. Overrides `MS_<ACCOUNT>_CLIENT_ID`.
- **`client_secret`** (`str | None`, default `None`) — Azure app client secret. Overrides `MS_<ACCOUNT>_CLIENT_SECRET`.
- **`tenant_id`** (`str | None`, default `None`) — Azure tenant (directory) ID. Overrides `MS_<ACCOUNT>_TENANT_ID`.
- **`username`** (`str | None`, default `None`) — Service account UPN (e.g. `svc@company.com`). Overrides `MS_<ACCOUNT>_USERNAME`.
- **`password`** (`str | None`, default `None`) — Service account password. Overrides `MS_<ACCOUNT>_PASSWORD`.
- **`default_site_id`** (`str | None`, default `None`) — SharePoint site ID (informational; not currently used in URL routing).
- **`default_drive_id`** (`str | None`, default `None`) — Drive ID. When set, all drive-scoped URLs use `/drives/{id}` instead of `/me/drive`. Readable from `MS_<ACCOUNT>_DEFAULT_DRIVE_ID`.
- **`max_concurrent`** (`int | None`, default `10`) — Maximum simultaneous in-flight Graph requests per client instance. Readable from `MS_<ACCOUNT>_MAX_CONCURRENT`.
- **`request_timeout`** (`float | None`, default `30.0`) — Per-request HTTP timeout in seconds. Readable from `MS_<ACCOUNT>_REQUEST_TIMEOUT`.

**Raises**

- **`KeyError`** — If any required credential (`client_id`, `client_secret`, `tenant_id`, `username`, `password`) is neither provided explicitly nor found in the environment.

**Notes**

Use as an async context manager to ensure the underlying HTTP client is closed properly. For scripts, call `close()` manually or use the `_sync` wrappers.

**Example**

```python
# From env vars
async with GraphClient("das_u1") as client:
    df = await client.read_excel(item_path="Reports/Q1.xlsx")

# Explicit credentials
client = GraphClient(
    client_id="...",
    client_secret="...",
    tenant_id="...",
    username="svc@company.com",
    password="...",
    default_drive_id="b!abc123",
    max_concurrent=5,
)
```

---

#### `GraphClient.close()`

Close the underlying HTTP client. Safe to call multiple times. Called automatically when the client is used as an async context manager.

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
) -> bytes
```

Download a file and return its raw content as bytes. Exactly one of the three identifier parameters must be provided.

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root, e.g. `"Documents/report.pdf"`.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link (full `https://…` URL from "Copy link").
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.

**Returns**

`bytes` — Raw file content.

**Raises**

- **`ValueError`** — If not exactly one identifier is provided.
- **`RuntimeError`** — If the Graph API response does not contain a `@microsoft.graph.downloadUrl`.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 403 Forbidden, 404 Not Found).

**Example**

```python
data = await client.download(item_path="images/logo.png")
with open("logo.png", "wb") as f:
    f.write(data)
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

**Parameters**

- **`item_path`** (`str | None`, default `None`) — Path relative to the drive root.
- **`share_url`** (`str | None`, default `None`) — SharePoint share link.
- **`item_id`** (`str | None`, default `None`) — Microsoft Graph item ID.
- **`sheet`** (`str | int`, default `0`) — Sheet to read. An integer is a zero-based positional index; a string is the exact sheet name.
- **`**kwargs`** — Forwarded verbatim to `pandas.read_excel` (e.g. `header`, `usecols`, `dtype`, `skiprows`).

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
- **`**kwargs`** — Forwarded verbatim to `pandas.read_excel`.

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

### File & Folder Management

---

#### `GraphClient.upload`

```python
async def upload(
    local_path: str | os.PathLike,
    remote_path: str,
) -> dict
```

Upload a local file to OneDrive / SharePoint. If the remote file already exists it is overwritten. The entire file is read into memory before upload; for files larger than ~4 MB consider using a Graph upload session directly.

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
    source_path: str,
    dest_folder_path: str | None = None,
    new_name: str | None = None,
) -> FileItem
```

Move and/or rename a file in a single API call. At least one of `dest_folder_path` or `new_name` must be provided.

**Parameters**

- **`source_path`** (`str`) — Current file path relative to the drive root.
- **`dest_folder_path`** (`str | None`, default `None`) — Target folder path relative to the drive root. Pass `""` to move to the drive root. `None` keeps the current parent folder.
- **`new_name`** (`str | None`, default `None`) — New file name including extension. `None` keeps the current name.

**Returns**

`FileItem` — Metadata of the item at its new location.

**Raises**

- **`ValueError`** — If neither `dest_folder_path` nor `new_name` is provided.
- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors.

**Example**

```python
# Move and rename in one call
item = await client.move_file(
    "Drafts/report_v2.xlsx",
    dest_folder_path="Published",
    new_name="report_final.xlsx",
)
print(item.path)  # "Published/report_final.xlsx"

# Rename only
await client.move_file("Reports/old_name.xlsx", new_name="new_name.xlsx")
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

### Communication

---

#### `GraphClient.send_mail`

```python
async def send_mail(
    to: str | list[str],
    subject: str,
    body: str,
    *,
    cc: str | list[str] | None = None,
) -> None
```

Send a plain-text email from the authenticated account's mailbox. The sent message is saved to Sent Items automatically.

**Parameters**

- **`to`** (`str | list[str]`) — Recipient email address or list of addresses.
- **`subject`** (`str`) — Email subject line.
- **`body`** (`str`) — Plain-text message body.
- **`cc`** (`str | list[str] | None`, default `None`) — CC recipient(s). `None` sends no CC.

**Returns**

`None`

**Raises**

- **`httpx.HTTPStatusError`** — On non-retryable HTTP errors (e.g. 403 if the account lacks `Mail.Send` permission).

**Example**

```python
await client.send_mail(
    to=["analyst@company.com", "manager@company.com"],
    subject="Daily ETL complete",
    body="The pipeline finished successfully. See the report on SharePoint.",
    cc="lead@company.com",
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

Convenience wrappers that call `asyncio.run()` internally. Intended for scripts, Prefect tasks, or any context without a running event loop.

> **Note:** Do not call these from within an `async def` function or inside Jupyter — use the native `await` syntax instead.

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

#### `GraphClient.download_sync(**kwargs)` → `bytes`

Synchronous wrapper for [`download`](#graphclientdownload). Accepts all the same keyword arguments.

```python
data = client.download_sync(item_path="archive/data.bin")
```
