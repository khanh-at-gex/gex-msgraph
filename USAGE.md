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

**Workflow A: pip + venv (Ubuntu)**
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
- **Send Teams:** `await client.send_teams_message("team1", "chan1", "Hello")`
- **Read Teams:** `msgs = await client.get_teams_messages("team1", "chan1", limit=10)`

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

### Instantiation

#### `GraphClient(account=None, **kwargs)`
Creates a new Graph API client.

**Params:**
- `account` (`str | None`): The prefix for environment variables (e.g. `ACCOUNT` will read `MS_ACCOUNT_CLIENT_ID`). Defaults to `"custom"`.
- `client_id`, `client_secret`, `tenant_id`, `username`, `password` (`str | None`): Explicit credentials, overriding environment variables.
- `default_site_id`, `default_drive_id` (`str | None`): The SharePoint site or Drive ID to target by default.
- `max_concurrent` (`int | None`): Maximum concurrent async requests.
- `request_timeout` (`float | None`): Timeout in seconds.

**Example:**
```python
async with GraphClient("my_app") as client:
    pass  # initialized with MS_MY_APP_CLIENT_ID, etc.
```

---

### File Operations (Download/Read)

#### `download(*, item_path=None, share_url=None, item_id=None)` → `bytes`
Return raw file bytes. Exactly one identifier must be provided.

```python
data = await client.download(item_path="documents/report.pdf")
with open("local_report.pdf", "wb") as f:
    f.write(data)
```

#### `read_excel(*, item_path=None, share_url=None, item_id=None, sheet=0, **kwargs)` → `pd.DataFrame`
Read an Excel file directly into a pandas DataFrame.

- `sheet` (`str | int`): Sheet name or positional index.
- `**kwargs`: Passed to `pandas.read_excel`.

```python
df = await client.read_excel(share_url="https://tenant.sharepoint.com/:x:/r/...", sheet="Data")
```

#### `read_csv(*, item_path=None, share_url=None, item_id=None, **kwargs)` → `pd.DataFrame`
Read a CSV file directly into a pandas DataFrame. `**kwargs` passed to `pandas.read_csv`.

```python
df = await client.read_csv(item_path="data/users.csv", sep=";")
```

#### `read_excel_many(paths, *, sheet=0, sheet_match="exact", on_missing_sheet="raise", on_error="raise", add_source_column=True, max_concurrent=None, return_status=False, **kwargs)`
Read many Excel files concurrently and concatenate into one DataFrame.

- `paths` (`list[str]`): List of paths to read.
- `sheet_match` (`Literal["exact", "ci", "glob"]`): How to match sheet name.
- `on_missing_sheet` (`Literal["raise", "skip", "warn"]`): Behavior when sheet not found.
- `on_error` (`Literal["raise", "skip", "warn"]`): Behavior when file fails.
- `add_source_column` (`bool`): Append `_source` column with originating path.
- `return_status` (`bool`): If true, returns `(combined_df, status_df)`.

```python
df = await client.read_excel_many(["jan.xlsx", "feb.xlsx"], sheet="Sales", on_error="skip")
```

#### `list_excel_sheets(*, item_path=None, share_url=None, item_id=None)` → `list[str]`
List all worksheet names in an Excel file.

```python
sheets = await client.list_excel_sheets(item_path="finance.xlsx")
# ['Summary', 'Q1', 'Q2', ...]
```

---

### File/Folder Management

#### `walk(folder_path="", *, pattern=None, recursive=True)` → `list[FileItem]`
List files under a folder. Folders are traversed but not returned.

```python
files = await client.walk("archive", pattern="*.csv")
```

#### `list_files(folder_path="")` → `list[FileItem]`
List immediate children (files and folders) of a single folder. Non-recursive.

```python
items = await client.list_files("projects")
```

#### `get_folder_tree(folder_path="")` → `TreeNode`
Returns a recursive tree representation of a folder and all its contents.

```python
tree = await client.get_folder_tree()
tree.print()
```

#### `get_metadata(*, item_path=None, share_url=None, item_id=None)` → `FileItem`
Fetch metadata for a single item without downloading its content.

```python
info = await client.get_metadata(item_path="shared/rules.txt")
print(info.modified, info.size)
```

#### `delete_file(*, item_path=None, share_url=None, item_id=None)` → `None`

```python
await client.delete_file(item_path="temp_dump.txt")
```

#### `move_file(source_path, dest_folder_path=None, new_name=None)` → `FileItem`

```python
await client.move_file("draft.docx", dest_folder_path="published", new_name="final.docx")
```

#### `create_folder(folder_path)` → `FileItem`

```python
await client.create_folder("2026/Q1")
```

#### `upload(local_path, remote_path)` → `dict`
Upload a local file. Returns Graph dictionary response.

```python
await client.upload("./local_cache/data.bin", "cloud_backup/data.bin")
```

---

### Communication

#### `send_mail(to, subject, body, *, cc=None)` → `None`
Send a plain-text email from the authenticated account's mailbox.

- `to` (`str | list[str]`): Recipient(s).
- `cc` (`str | list[str] | None`): Optional CC recipient(s).

```python
await client.send_mail("manager@company.com", "Pipeline Status", "Job completed.")
```

#### `send_teams_message(team_id, channel_id, text)` → `None`

```python
await client.send_teams_message("team-xyz", "channel-abc", "Pipeline failed!")
```

#### `get_teams_messages(team_id, channel_id, limit=10)` → `list[dict]`

```python
messages = await client.get_teams_messages("team-xyz", "channel-abc", limit=5)
```

---

### Synchronous Helpers

```python
df   = client.read_excel_sync(item_path="Reports/Q1.xlsx")
df   = client.read_csv_sync(item_path="data.csv")
data = client.download_sync(item_path="image.png")
```
