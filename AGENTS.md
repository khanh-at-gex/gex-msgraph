# gex-msgraph

> Async Python wrapper for Microsoft Graph API. Provides clean access to M365 services (OneDrive, SharePoint, Outlook, Teams) from scripts, FastAPI, Prefect, or Jupyter notebooks.

Public exports: `from gex_msgraph import GraphClient, FileItem, TreeNode, GraphError, AuthError, PermissionDeniedError, NotFoundError, RateLimitExhaustedError, GraphAuthenticationError, GraphSyncInLoopError`

## Setup

Install:
```bash
pip install "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

`.env` file (never commit):
```
MS_DAS_U1_CLIENT_ID=...
MS_DAS_U1_CLIENT_SECRET=...
MS_DAS_U1_TENANT_ID=...
MS_DAS_U1_USERNAME=service@company.com   # omit BOTH username+password for app-only auth
MS_DAS_U1_PASSWORD=...
MS_DAS_U1_DEFAULT_DRIVE_ID=...   # optional (required for app-only auth)
```

Load env in code (`python-dotenv` is no longer bundled since v0.3.0 — add it to your own project, or install `gex-msgraph[dotenv]`):
```python
from dotenv import load_dotenv
load_dotenv()
```

## Key rules

- All methods are `async` — use `await` or the `_sync` wrappers below.
- Use as async context manager: `async with GraphClient("das_u1") as client:`
- Exactly one of `item_path`, `share_url`, or `item_id` must be provided per call.
- `item_path` is always relative to the drive root, e.g. `"Reports/Q1.xlsx"` (no leading slash).
- All Graph HTTP errors raise a `GraphError` subclass (`AuthError` 401, `PermissionDeniedError` 403, `NotFoundError` 404, `RateLimitExhaustedError` when retries run out). `GraphError` subclasses `httpx.HTTPStatusError`, so old `except httpx.HTTPStatusError` code still works.
- App-only auth (no username/password): only drive operations work, and `default_drive_id` is required. Mail/Teams/chat methods need delegated (username+password) auth.
- Since v0.4.0, `read_excel`/`read_excel_many` auto-detect **SpreadsheetML 2003** files (SAP exports named `.xls` that are really UTF-16 XML) and decode them. Values only — styles/formulas dropped, `DateTime` cells arrive as strings. Sheet names illegal in xlsx (`\ / * ? : [ ]`, or >31 chars) are renamed on conversion; pass `sheet` as it appears in the source and it is mapped for you, except with `sheet_match="glob"` where the pattern is used verbatim. `list_excel_sheets` does **not** work on them (Graph workbook API needs a real xlsx; returns 403).

## GraphClient — instantiation

```python
from gex_msgraph import GraphClient

# From env vars (MS_DAS_U1_*)
async with GraphClient("das_u1") as client:
    ...

# Explicit credentials
client = GraphClient(
    client_id="...", client_secret="...", tenant_id="...",
    username="svc@company.com", password="...",
    default_drive_id="b!abc123",  # optional
)
```

## File operations — read

```python
# Read Excel → DataFrame
df = await client.read_excel(item_path="Reports/Q1.xlsx")
df = await client.read_excel(item_path="Reports/Q1.xlsx", sheet="Summary")
df = await client.read_excel(share_url="https://tenant.sharepoint.com/:x:/r/...", sheet=1)

# SAP-style ".xls" that is really SpreadsheetML 2003 XML — decoded automatically
df = await client.read_excel(item_path="TB/jan.xls", header=None)
df = await client.read_excel_many(["TB/jan.xls", "TB/feb.xls"], header=None)

# Read CSV → DataFrame
df = await client.read_csv(item_path="data/users.csv", sep=";")

# Read Parquet → DataFrame
df = await client.read_parquet(item_path="data/users.parquet")

# Download raw bytes
data = await client.download(item_path="images/logo.png")

# Download a large file straight to disk (streamed, not buffered)
path = await client.download(item_path="exports/huge.parquet", to_path="huge.parquet")

# Bulk read Excel (concurrent)
df = await client.read_excel_many(["jan.xlsx", "feb.xlsx"], sheet="Sales", on_error="warn")
df, status = await client.read_excel_many(
    ["jan.xlsx", "feb.xlsx"],
    sheet="Sales*", sheet_match="glob",      # "exact" | "ci" | "glob"
    on_missing_sheet="warn",                  # "raise" | "skip" | "warn"
    on_error="warn",
    return_status=True,                       # returns (combined_df, status_df)
)

# Engine defaults to "calamine" (fast). Override explicitly if needed:
df = await client.read_excel_many(
    ["jan.xlsx", "feb.xlsx"], sheet="Sales",
    usecols="A:M", engine="openpyxl",
)

# Bulk read CSV (concurrent)
df, status = await client.read_csv_many(
    ["exports/jan.csv", "exports/feb.csv"],
    on_error="warn", return_status=True, sep=";"
)

# List worksheet names (no download)
sheets = await client.list_excel_sheets(item_path="Finance/budget.xlsx")
# → ['Summary', 'Q1', 'Q2', ...]

# Search files across the drive
files = await client.search_files("budget")          # → list[FileItem]
```

## Bulk read sync wrappers (scripts / no event loop)

```python
df = client.read_excel_many_sync(["jan.xlsx", "feb.xlsx"], sheet="Sales")
df = client.read_csv_many_sync(["jan.csv", "feb.csv"])
```

## File & folder discovery

```python
# List immediate children (files + folders)
items = await client.list_files("Reports")

# Walk recursively, filter by glob
files = await client.walk("Reports", pattern="*.xlsx")
files = await client.walk("Reports", pattern="*.xlsx", recursive=False)

# Get single item metadata
meta = await client.get_metadata(item_path="Reports/Q1.xlsx")
# meta.name, meta.path, meta.size, meta.modified, meta.is_folder, meta.id

# Visualise folder tree
tree = await client.get_folder_tree("Reports")
tree.print()
```

## File & folder management

```python
# Upload local file (any size — files over 4 MiB are chunked automatically)
result = await client.upload("./local/report.xlsx", "Reports/2026/report.xlsx")

# Bulk upload (concurrent, optionally bounded)
results = await client.upload_many([
    ("./local/jan.xlsx", "Reports/jan.xlsx"),
    ("./local/feb.xlsx", "Reports/feb.xlsx"),
], max_concurrent=5)

# Delete (moves to recycle bin)
await client.delete_file(item_path="temp/scratch.xlsx")

# Copy file (keyword-only since v0.3.0; source accepts item_path | share_url | item_id)
await client.copy_file(item_path="Reports/Q1.xlsx", dest_folder_path="Archive", new_name="Q1_backup.xlsx")
item = await client.copy_file(item_path="Reports/Q1.xlsx", dest_folder_path="Archive", wait=True)  # blocks until done, returns FileItem

# Move and/or rename (keyword-only since v0.3.0)
await client.move_file(item_path="Drafts/v2.xlsx", dest_folder_path="Published", new_name="final.xlsx")
await client.move_file(item_path="Reports/old.xlsx", new_name="new.xlsx")  # rename only

# Create folder
folder = await client.create_folder("Reports/2026/Q1")

# Check existence
if await client.exists(item_path="Reports/Q1.xlsx"):
    ...

# Get a sharing link
url = await client.get_share_link(item_path="Reports/Q1.xlsx")
url = await client.get_share_link(item_path="Reports/Q1.xlsx", link_type="edit", scope="anonymous")
```

## Communication

```python
# List inbox messages
msgs = await client.list_mail(limit=10)
# msgs[0] keys: id, subject, from, receivedDateTime, bodyPreview, hasAttachments
msgs = await client.list_mail(limit=20, folder="sentitems")
```

```python
# Send plain-text email
await client.send_mail(
    to=["analyst@company.com", "manager@company.com"],
    subject="Pipeline complete",
    body="Job finished.",
    cc="lead@company.com",                            # optional
)

# Send HTML email with attachments
await client.send_mail(
    to="analyst@company.com",
    subject="Report",
    body="<h2>Done</h2><p>See attached.</p>",
    body_type="html",                                 # "text" (default) | "html"
    attachments=[
        "./output/report.xlsx",                       # local file path
        ("data.csv", df.to_csv().encode()),            # (filename, bytes) tuple
        {"item_path": "Reports/Q1.xlsx"},             # SharePoint file by path
        {"share_url": "https://tenant.sharepoint.com/:x:/r/..."},  # by share link
    ],
)

# Teams channel
await client.send_teams_message(team_id="19:...", channel_id="19:...", text="Done.")
messages = await client.get_teams_messages(team_id="19:...", channel_id="19:...", limit=10)

# Teams chat (1-1 or group)
chats = await client.list_chats()
chat_id = chats[0]["id"]
messages = await client.get_chat_messages(chat_id, limit=10)
await client.send_chat_message(chat_id, "ETL finished.")
```

## Sync wrappers (scripts / notebooks without event loop)

```python
client = GraphClient("das_u1")
df   = client.read_excel_sync(item_path="Reports/Q1.xlsx")
df   = client.read_csv_sync(item_path="data.csv")
data = client.download_sync(item_path="image.png")
client.close_sync()  # when done
```

> In Jupyter, use `await` directly — `_sync` wrappers raise `GraphSyncInLoopError` inside a running event loop.

## FileItem fields

`name`, `path`, `id`, `size` (bytes), `modified` (datetime UTC), `is_folder` (bool), `webUrl` (`str | None` — real SharePoint/OneDrive item URL, safe for Office Online embed iframes; `get_share_link`'s `/:x:/` links are NOT safe for embedding)

## Error handling

```python
from gex_msgraph import GraphError, NotFoundError, PermissionDeniedError

try:
    df = await client.read_excel(item_path="Reports/Q1.xlsx")
except NotFoundError:
    print("path not found")
except PermissionDeniedError:
    print("no permission")
except GraphError as e:  # anything else from Graph
    print(e.response.status_code, e.response.text)
```

`GraphError` subclasses `httpx.HTTPStatusError`, so legacy `except httpx.HTTPStatusError` still works. `429`/5xx are auto-retried; `RateLimitExhaustedError` fires only if all retries fail.

## Full API reference

See [USAGE.md](USAGE.md) for complete signature, parameter descriptions, and examples for every method.
