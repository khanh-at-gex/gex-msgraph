# Microsoft Graph Client API Reference

This document provides a comprehensive map of all methods available in the `GraphClient` class.

## Overview Map

### Instantiation & Context
- `__init__` - Create a new client instance
- `close` - Close the underlying HTTP client
- `__aenter__` / `__aexit__` - Async context manager

### File Operations (Download/Read)
- `download` - Get raw bytes of a file
- `read_excel` - Read an Excel file directly into a pandas DataFrame
- `read_csv` - Read a CSV file directly into a pandas DataFrame
- `read_excel_many` - Read multiple Excel files concurrently and concatenate them
- `list_excel_sheets` - List all worksheets in an Excel file

### File/Folder Management
- `walk` - Recursively list all files under a folder
- `list_files` - List immediate children of a folder
- `get_folder_tree` - Get a recursive tree representation of a folder
- `get_metadata` - Fetch metadata (FileItem) for a single item
- `delete_file` - Delete an item
- `move_file` - Move or rename an item
- `create_folder` - Create a new folder
- `upload` - Upload a local file to a remote path

### Communication
- `send_mail` - Send an email
- `send_teams_message` - Post a message to a Teams channel
- `get_teams_messages` - Fetch messages from a Teams channel

### Synchronous Helpers
- `read_excel_sync` - Synchronous wrapper for `read_excel`
- `read_csv_sync` - Synchronous wrapper for `read_csv`
- `download_sync` - Synchronous wrapper for `download`

---

## Detailed Method Reference

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
from gex_msgraph import GraphClient

async def main():
    async with GraphClient("my_app") as client:
        # Client initialized with MS_MY_APP_CLIENT_ID, etc.
        pass
```

---

### File Operations (Download/Read)

#### `download(*, item_path=None, share_url=None, item_id=None)` -> `bytes`
Return raw file bytes. Exactly one identifier must be provided.

**Params:**
- `item_path` (`str | None`): The path in the default drive.
- `share_url` (`str | None`): A SharePoint share link.
- `item_id` (`str | None`): The direct Graph item ID.

**Example:**
```python
data = await client.download(item_path="documents/report.pdf")
with open("local_report.pdf", "wb") as f:
    f.write(data)
```

#### `read_excel(*, item_path=None, share_url=None, item_id=None, sheet=0, **kwargs)` -> `pd.DataFrame`
Read an Excel file directly into a pandas DataFrame.

**Params:**
- Identifiers: `item_path`, `share_url`, `item_id` (Provide exactly one)
- `sheet` (`str | int`): The sheet name or positional index to read.
- `**kwargs`: Any additional keyword arguments supported by `pandas.read_excel`.

**Example:**
```python
df = await client.read_excel(share_url="https://tenant.sharepoint.com/:x:/r/...", sheet="Data")
print(df.head())
```

#### `read_csv(*, item_path=None, share_url=None, item_id=None, **kwargs)` -> `pd.DataFrame`
Read a CSV file directly into a pandas DataFrame.

**Params:**
- Identifiers: `item_path`, `share_url`, `item_id` (Provide exactly one)
- `**kwargs`: Additional arguments passed to `pandas.read_csv`.

**Example:**
```python
df = await client.read_csv(item_path="data/users.csv", sep=";")
```

#### `read_excel_many(paths, *, sheet=0, sheet_match="exact", on_missing_sheet="raise", on_error="raise", add_source_column=True, max_concurrent=None, return_status=False, **kwargs)`
Read many Excel files concurrently and concatenate them into one DataFrame.

**Params:**
- `paths` (`list[str]`): List of paths to read.
- `sheet` (`str | int`): Sheet to read.
- `sheet_match` (`Literal["exact", "ci", "glob"]`): How to match the sheet name.
- `on_missing_sheet` (`Literal["raise", "skip", "warn"]`): Behavior when sheet isn't found.
- `on_error` (`Literal["raise", "skip", "warn"]`): Behavior when file fails to download or parse.
- `add_source_column` (`bool`): Whether to append a `_source` column indicating the originating path.
- `return_status` (`bool`): If true, returns a tuple of `(combined_df, status_df)`.

**Example:**
```python
paths = ["january.xlsx", "february.xlsx", "march.xlsx"]
df = await client.read_excel_many(paths, sheet="Sales", on_error="skip")
```

#### `list_excel_sheets(*, item_path=None, share_url=None, item_id=None)` -> `list[str]`
List all worksheet names in an Excel file.

**Example:**
```python
sheets = await client.list_excel_sheets(item_path="finance.xlsx")
print(sheets) # ['Summary', 'Q1', 'Q2', ...]
```

---

### File/Folder Management

#### `walk(folder_path="", *, pattern=None, recursive=True)` -> `list[FileItem]`
List files under a folder. Folders are traversed but not returned in the result.

**Params:**
- `folder_path` (`str`): Path to the folder. Defaults to root.
- `pattern` (`str | None`): Glob pattern to filter names (e.g. `*.xlsx`).
- `recursive` (`bool`): If true, navigates down subfolders.

**Example:**
```python
files = await client.walk("archive", pattern="*.csv")
for f in files:
    print(f.path)
```

#### `list_files(folder_path="")` -> `list[FileItem]`
List immediate children (files and folders) of a single folder. Non-recursive.

**Example:**
```python
items = await client.list_files("projects")
for i in items:
    print(f"{i.name} (Folder? {i.is_folder})")
```

#### `get_folder_tree(folder_path="")` -> `TreeNode`
Returns a recursive tree representation of a folder and all its contents.

**Example:**
```python
tree = await client.get_folder_tree()
tree.print() # Prints a visual hierarchy to stdout
```

#### `get_metadata(*, item_path=None, share_url=None, item_id=None)` -> `FileItem`
Fetch metadata for a single item without downloading its content.

**Example:**
```python
info = await client.get_metadata(item_path="shared/rules.txt")
print(info.modified, info.size)
```

#### `delete_file(*, item_path=None, share_url=None, item_id=None)` -> `None`
Delete an item.

**Example:**
```python
await client.delete_file(item_path="temp_dump.txt")
```

#### `move_file(source_path, dest_folder_path=None, new_name=None)` -> `FileItem`
Move or rename a file.

**Params:**
- `source_path` (`str`): Original path.
- `dest_folder_path` (`str | None`): The new parent folder.
- `new_name` (`str | None`): The new file name.

**Example:**
```python
await client.move_file("draft.docx", dest_folder_path="published", new_name="final.docx")
```

#### `create_folder(folder_path)` -> `FileItem`
Create a new folder.

**Example:**
```python
await client.create_folder("2026/Q1")
```

#### `upload(local_path, remote_path)` -> `dict`
Upload a local file to the specified remote path. Returns Graph dictionary response.

**Example:**
```python
await client.upload("./local_cache/data.bin", "cloud_backup/data.bin")
```

---

### Communication

#### `send_mail(to, subject, body, *, cc=None)` -> `None`
Send a plain-text email from the authenticated account's mailbox.

**Params:**
- `to` (`str | list[str]`): Recipient email address(es).
- `subject` (`str`): Email subject line.
- `body` (`str`): Plain-text email body.
- `cc` (`str | list[str] | None`): Optional CC recipient(s).

**Example:**
```python
await client.send_mail("manager@company.com", "Pipeline Status", "Job completed successfully.")
```

#### `send_teams_message(team_id, channel_id, text)` -> `None`
Post a plain-text message to a specific Teams channel.

**Example:**
```python
await client.send_teams_message("team-xyz", "channel-abc", "Pipeline failed!")
```

#### `get_teams_messages(team_id, channel_id, limit=10)` -> `list[dict]`
Fetch incoming messages from a Teams channel.

**Example:**
```python
messages = await client.get_teams_messages("team-xyz", "channel-abc", limit=5)
```
