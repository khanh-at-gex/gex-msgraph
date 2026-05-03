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
- **Bulk read:** `await client.read_excel_many(["A.xlsx", "B.xlsx"], on_error="warn")`
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
