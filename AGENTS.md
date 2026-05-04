# gex-msgraph

> Async Python wrapper for Microsoft Graph API. Provides clean access to M365 services (OneDrive, SharePoint, Outlook, Teams) from scripts, FastAPI, Prefect, or Jupyter notebooks.

Public exports: `from gex_msgraph import GraphClient, FileItem, TreeNode`

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
MS_DAS_U1_USERNAME=service@company.com
MS_DAS_U1_PASSWORD=...
MS_DAS_U1_DEFAULT_DRIVE_ID=...   # optional
```

Load env in code:
```python
from dotenv import load_dotenv
load_dotenv()
```

## Key rules

- All methods are `async` — use `await` or the `_sync` wrappers below.
- Use as async context manager: `async with GraphClient("das_u1") as client:`
- Exactly one of `item_path`, `share_url`, or `item_id` must be provided per call.
- `item_path` is always relative to the drive root, e.g. `"Reports/Q1.xlsx"` (no leading slash).
- All HTTP errors raise `httpx.HTTPStatusError`. Wrap calls with `try/except httpx.HTTPStatusError`.

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

# Read CSV → DataFrame
df = await client.read_csv(item_path="data/users.csv", sep=";")

# Download raw bytes
data = await client.download(item_path="images/logo.png")

# Bulk read Excel (concurrent)
df = await client.read_excel_many(["jan.xlsx", "feb.xlsx"], sheet="Sales", on_error="warn")
df, status = await client.read_excel_many(
    ["jan.xlsx", "feb.xlsx"],
    sheet="Sales*", sheet_match="glob",      # "exact" | "ci" | "glob"
    on_missing_sheet="warn",                  # "raise" | "skip" | "warn"
    on_error="warn",
    return_status=True,                       # returns (combined_df, status_df)
)

# Bulk read CSV (concurrent)
df, status = await client.read_csv_many(
    ["exports/jan.csv", "exports/feb.csv"],
    on_error="warn", return_status=True, sep=";"
)

# List worksheet names (no download)
sheets = await client.list_excel_sheets(item_path="Finance/budget.xlsx")
# → ['Summary', 'Q1', 'Q2', ...]
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
# Upload local file
result = await client.upload("./local/report.xlsx", "Reports/2026/report.xlsx")

# Delete (moves to recycle bin)
await client.delete_file(item_path="temp/scratch.xlsx")

# Move and/or rename
await client.move_file("Drafts/v2.xlsx", dest_folder_path="Published", new_name="final.xlsx")
await client.move_file("Reports/old.xlsx", new_name="new.xlsx")  # rename only

# Create folder
folder = await client.create_folder("Reports/2026/Q1")
```

## Communication

```python
# Send email
await client.send_mail(
    to=["analyst@company.com", "manager@company.com"],
    subject="Pipeline complete",
    body="Job finished.",
    cc="lead@company.com",   # optional
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
```

> In Jupyter, use `await` directly — do NOT use `_sync` wrappers inside a notebook.

## FileItem fields

`name`, `path`, `id`, `size` (bytes), `modified` (datetime UTC), `is_folder` (bool)

## Error handling

```python
import httpx

try:
    df = await client.read_excel(item_path="Reports/Q1.xlsx")
except httpx.HTTPStatusError as e:
    print(e.response.status_code, e.response.text)
```

Common status codes: `403` = no permission, `404` = path not found, `429` = rate limit (auto-retried).

## Full API reference

See [USAGE.md](USAGE.md) for complete signature, parameter descriptions, and examples for every method.
