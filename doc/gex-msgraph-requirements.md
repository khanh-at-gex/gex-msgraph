# `gex-msgraph` — Build Requirements

> **For Claude Code:** complete spec. Build exactly as described. Defaults are listed where decisions are needed — only ask if something is genuinely ambiguous.

---

## 1. Identity

| | |
|---|---|
| Repo / dist name | `gex-msgraph` |
| Package / import name | `gex_msgraph` |
| Public exports | `GraphClient`, `FileItem` |
| Env var prefix | `MS_` |
| Python | `>=3.11` |
| Build backend | `hatchling` |
| Visibility | Private repo, not published to PyPI |

---

## 2. Purpose

Async Python wrapper around Microsoft Graph API for clean access to M365 (SharePoint, OneDrive, Outlook, Teams) from any context — FastAPI, Prefect, scripts, notebooks.

Make the 90% case trivial. Leave bytes-level escape hatches for the 10%.

---

## 3. Repository Layout

```
gex-msgraph/
├── pyproject.toml
├── README.md                 # what the library is + quick reference
├── USAGE.md                  # end-user guide: install → setup → daily use
├── AGENTS.md                 # for AI coding assistants
├── CHANGELOG.md
├── .env.example
├── .gitignore
├── .github/
│   └── workflows/
│       └── ci.yml            # lint + type check + tests on push
├── src/
│   └── gex_msgraph/
│       ├── __init__.py
│       ├── _core.py
│       ├── _files.py
│       └── py.typed
└── tests/
    ├── conftest.py
    ├── test_files.py
    └── test_client.py
```

Standard regular package — no namespace package complications.

---

## 4. Module Map — what lives where

### `src/gex_msgraph/__init__.py`

Re-exports the public API. Nothing else.

```python
from gex_msgraph._core import GraphClient
from gex_msgraph._files import FileItem

__all__ = ["GraphClient", "FileItem"]
__version__ = "0.1.0"
```

### `src/gex_msgraph/_core.py`

The `GraphClient` class plus its private auth helper.

**Classes:**

| Class | Purpose |
|---|---|
| `_TokenProvider` (private) | Wraps MSAL ROPC flow. Caches tokens. Refreshes transparently. |
| `GraphClient` (public) | The whole library — auth + http + all methods |

**Module-level constants:**

| Name | Value | Purpose |
|---|---|---|
| `_GRAPH_BASE` | `"https://graph.microsoft.com/v1.0"` | Base URL for all API calls |
| `_AUTHORITY_TEMPLATE` | `"https://login.microsoftonline.com/{tenant_id}"` | MSAL authority URL |
| `_DEFAULT_SCOPE` | `["https://graph.microsoft.com/.default"]` | OAuth scope |
| `_DEFAULT_TIMEOUT` | `30.0` | HTTP timeout in seconds |
| `_DEFAULT_MAX_CONCURRENT` | `10` | Default semaphore cap |
| `_DEFAULT_MAX_RETRIES` | `3` | Retry attempts on 429/5xx |
| `_BACKOFF_CAP` | `30.0` | Max backoff seconds |

**Module-level functions:**

| Function | Signature | Purpose |
|---|---|---|
| `_load_account_env` | `(name: str) -> dict[str, str]` | Read all `MS_<NAME>_*` env vars, validate required, return as dict. Raises `KeyError` on missing required. |
| `_compute_backoff` | `(attempt: int, retry_after: str \| None) -> float` | Compute backoff seconds: `Retry-After` header if present, else `min(2**attempt, _BACKOFF_CAP)` |

### `src/gex_msgraph/_files.py`

File identifier resolution, share URL encoding, sheet matching, dataclass.

**Classes:**

| Class | Purpose |
|---|---|
| `FileItem` (public dataclass) | Represents one file/folder in SharePoint or OneDrive |

**Module-level functions:**

| Function | Signature | Purpose |
|---|---|---|
| `validate_identifier` | `(item_path, share_url, item_id) -> Literal["path","share","id"]` | Ensure exactly one is provided. Raises `ValueError` otherwise. |
| `encode_share_url` | `(url: str) -> str` | Returns `"u!" + base64.urlsafe_b64encode(url.encode()).rstrip("=")` |
| `build_resolution_url` | `(kind, value, drive_id) -> str` | Build the right Graph endpoint for the identifier kind |
| `match_sheet_name` | `(sheets: list[str], requested: str \| int, mode: Literal["exact","ci","glob"]) -> str \| int \| None` | Find matching sheet name; return `None` if no match |
| `parse_drive_item` | `(item: dict) -> FileItem` | Convert Graph API response dict into `FileItem` |

---

## 5. Dependencies

```toml
[project]
dependencies = [
    "httpx>=0.27",
    "msal>=1.28",
    "pandas>=2.0",
    "openpyxl>=3.1",
    "python-dotenv>=1.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-asyncio>=0.23",
    "respx>=0.21",
    "ruff>=0.6",
    "mypy>=1.10",
]
```

**Use `[project.optional-dependencies]` for dev, NOT `[dependency-groups]`.** Reason: `[project.optional-dependencies]` is standard PEP 621 — works with both `pip install -e .[dev]` and `uv sync --extra dev`. `[dependency-groups]` is uv-specific and breaks pip workflows.

Library installs cleanly with either tool:
- `pip install gex-msgraph` (or `pip install -e .` for development)
- `uv add gex-msgraph` (or `uv sync` for development)

Both must work. Test both before release.

---

## 6. Auth

ROPC flow via `msal.PublicClientApplication.acquire_token_by_username_password()`. MSAL handles caching and refresh automatically.

**Critical:** never log password or token. Never include them in exception messages.

`_TokenProvider` lives inside `_core.py` (no separate `_auth.py` — MSAL wrapper is ~30 lines).

### `_TokenProvider` — exact methods

```python
class _TokenProvider:
    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        username: str,
        password: str,
    ) -> None: ...

    def get_token(self) -> str:
        """Returns access token. Synchronous (MSAL is sync). Cached automatically."""
```

That's it. One method. Sync because MSAL is sync — wrap in `asyncio.to_thread()` when called from async context.

---

## 7. `.env` Schema

```bash
MS_ACCOUNTS=das_u1,das_u2

# Per account — REQUIRED
MS_DAS_U1_CLIENT_ID=
MS_DAS_U1_CLIENT_SECRET=
MS_DAS_U1_TENANT_ID=
MS_DAS_U1_USERNAME=
MS_DAS_U1_PASSWORD=

# Per account — OPTIONAL
MS_DAS_U1_DEFAULT_SITE_ID=        # SharePoint default site
MS_DAS_U1_DEFAULT_DRIVE_ID=       # SharePoint default drive
MS_DAS_U1_MAX_CONCURRENT=10       # default 10
MS_DAS_U1_REQUEST_TIMEOUT=30      # default 30 seconds
```

Account name passed as `"das_u1"` (lowercase). Library uppercases for env lookup.

If `DEFAULT_DRIVE_ID` unset → falls back to `/me/drive` (account's OneDrive).

Missing required env var → `KeyError` at `GraphClient(name)` construction with the variable name.

---

## 8. Public API — full method list

### `GraphClient` — complete signatures

```python
class GraphClient:
    # ─────────────────────────────────────────────────────────────
    # Construction & lifecycle
    # ─────────────────────────────────────────────────────────────
    def __init__(self, account: str) -> None: ...
    
    async def close(self) -> None:
        """Close the underlying httpx client. Safe to call multiple times."""
    
    async def __aenter__(self) -> "GraphClient": ...
    async def __aexit__(self, *exc: Any) -> None: ...

    # ─────────────────────────────────────────────────────────────
    # Read single file — exactly one of item_path/share_url/item_id
    # ─────────────────────────────────────────────────────────────
    async def read_excel(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        sheet: str | int = 0,
        **read_excel_kwargs: Any,
    ) -> pd.DataFrame:
        """Read an Excel file directly into a DataFrame. Streams via BytesIO."""

    async def read_csv(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        **read_csv_kwargs: Any,
    ) -> pd.DataFrame:
        """Read a CSV file directly into a DataFrame. Streams via BytesIO."""

    async def download(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> bytes:
        """Return raw file bytes. Use for non-pandas parsing."""

    # ─────────────────────────────────────────────────────────────
    # Bulk read
    # ─────────────────────────────────────────────────────────────
    async def read_excel_many(
        self,
        paths: list[str],
        *,
        sheet: str | int = 0,
        sheet_match: Literal["exact", "ci", "glob"] = "exact",
        on_missing_sheet: Literal["raise", "skip", "warn"] = "raise",
        on_error: Literal["raise", "skip", "warn"] = "raise",
        add_source_column: bool = True,
        max_concurrent: int | None = None,
        **read_excel_kwargs: Any,
    ) -> pd.DataFrame:
        """Read many Excel files concurrently and concat into one DataFrame."""
        
    async def read_csv_many(
        self,
        paths: list[str],
        *,
        on_error: Literal["raise", "skip", "warn"] = "raise",
        add_source_column: bool = True,
        max_concurrent: int | None = None,
        return_status: bool = False,
        **read_csv_kwargs: Any,
    ) -> pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]:
        """Read many CSV files concurrently and concat into one DataFrame."""

    # ─────────────────────────────────────────────────────────────
    # Discovery
    # ─────────────────────────────────────────────────────────────
    async def walk(
        self,
        folder_path: str = "",
        *,
        pattern: str | None = None,
        recursive: bool = True,
    ) -> list[FileItem]:
        """List files (not folders) under a folder. Recursive by default."""

    async def list_files(self, folder_path: str = "") -> list[FileItem]:
        """List immediate children (files and folders) of one folder. Non-recursive."""

    # ─────────────────────────────────────────────────────────────
    # Write
    # ─────────────────────────────────────────────────────────────
    async def upload(self, local_path: Path, remote_path: str) -> dict:
        """Upload a local file to remote_path. Returns the Graph driveItem dict."""

    # ─────────────────────────────────────────────────────────────
    # Communication
    # ─────────────────────────────────────────────────────────────
    async def send_mail(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        *,
        cc: str | list[str] | None = None,
    ) -> None:
        """Send a plain-text email from the account's mailbox."""

    async def send_teams_message(
        self,
        team_id: str,
        channel_id: str,
        text: str,
    ) -> None:
        """Post a plain-text message to a Teams channel."""

    # ─────────────────────────────────────────────────────────────
    # Sync wrappers — for scripts/notebooks only
    # ─────────────────────────────────────────────────────────────
    def read_excel_sync(self, **kw: Any) -> pd.DataFrame: ...
    def read_csv_sync(self, **kw: Any) -> pd.DataFrame: ...
    def download_sync(self, **kw: Any) -> bytes: ...
```

### `GraphClient` — private methods (not exposed)

| Method | Purpose |
|---|---|
| `_request(method, url, **kw) -> httpx.Response` | Unified entry: acquires token, adds Bearer header, applies retry, applies semaphore. ALL Graph calls go through this. |
| `_get_download_url(*, item_path, share_url, item_id) -> str` | Resolve any identifier to `@microsoft.graph.downloadUrl`. Uses `_files.validate_identifier` and `_files.build_resolution_url`. |
| `_stream_to_bytes(url: str) -> bytes` | Stream a download URL into bytes. No auth header (URL is pre-signed). |
| `_iter_paginated(url: str) -> AsyncIterator[dict]` | Yield items across `@odata.nextLink` pages. |
| `_resolve_drive_id() -> str` | Returns configured drive ID or `"me"` for `/me/drive`. |

### `FileItem` — exact shape

```python
@dataclass(frozen=True)
class FileItem:
    name: str            # "Q1.xlsx"
    path: str            # "Reports/2026-Q1/Q1.xlsx" — relative to drive root
    id: str              # Graph item ID
    size: int            # bytes
    modified: datetime   # last modified
    is_folder: bool
```

---

## 9. File Identifier Resolution

Each read method requires exactly ONE of:
- `item_path` — relative to drive root, e.g. `"Reports/Q1.xlsx"`
- `share_url` — full SharePoint share link (from "Copy link")
- `item_id` — Graph item ID

Zero or multiple → `ValueError`.

**Resolution endpoints:**

| Identifier | Endpoint |
|---|---|
| `item_path` | `GET /drives/{drive_id}/root:/{path}` |
| `share_url` | `GET /shares/{base64_url}/driveItem` |
| `item_id` | `GET /drives/{drive_id}/items/{item_id}` |

**Drive selection:** uses `MS_<ACCOUNT>_DEFAULT_DRIVE_ID` if set, else `/me/drive`.

After resolving, extract `@microsoft.graph.downloadUrl` from response, then stream that URL **without** auth header (it's pre-authenticated, short-lived).

---

## 10. Performance

| Concern | Implementation |
|---|---|
| Connection reuse | One `httpx.AsyncClient` per `GraphClient`, lifetime = client lifetime |
| Concurrency cap | `asyncio.Semaphore(max_concurrent)`, default 10 |
| Token caching | MSAL in-memory cache (built-in) |
| Retries | Hand-rolled. 429/503/5xx → exponential backoff `2^attempt` (capped 30s), max 3 retries. Respect `Retry-After` header. 4xx other than 429 → no retry, raise. |
| Pagination | `_iter_paginated` follows `@odata.nextLink` for `walk` / `list_files` |
| File streaming | `client.stream("GET", url)` → `BytesIO` |
| Bulk reads | `asyncio.gather` bounded by semaphore |

No `tenacity` dependency. Inline retry helper using `_compute_backoff`.

---

## 11. Bulk Read — `read_excel_many` algorithm

For each path:
1. Resolve identifier → download URL
2. Stream to `BytesIO`
3. Open with `pd.ExcelFile(buf)`, get sheet names
4. Match `sheet` against actual sheets per `sheet_match` mode (use `_files.match_sheet_name`)
5. If no match → apply `on_missing_sheet` mode
6. Apply `**read_excel_kwargs` (skiprows, header, dtype, etc.)
7. If `add_source_column` → assign `_source` column with the path
8. Append DataFrame to results list

After all paths processed:
- Concat with `pd.concat(dfs, ignore_index=True, sort=False)`
- If results list is empty → return empty DataFrame (don't raise)

**Sheet matching modes:**

| Mode | Behavior |
|---|---|
| `"exact"` | Exact string match |
| `"ci"` | Case-insensitive |
| `"glob"` | `fnmatch.fnmatch` — first match wins per file |

If `sheet` is `int`, `sheet_match` is ignored (positional index).

**Error modes (`on_missing_sheet` and `on_error`):**

| Mode | Behavior |
|---|---|
| `"raise"` | First failure halts everything |
| `"warn"` | Log warning, skip file, continue |
| `"skip"` | Silent skip, continue |

---

## 12. Sync Wrappers

Three only: `read_excel_sync`, `read_csv_sync`, `download_sync`.

```python
def read_excel_sync(self, **kw: Any) -> pd.DataFrame:
    return asyncio.run(self.read_excel(**kw))
```

Document in README: not callable from inside a running event loop. Scripts/notebooks only.

---

## 13. Errors

No custom hierarchy. Stdlib + httpx exceptions only.

| Exception | When |
|---|---|
| `ValueError` | Bad args (missing/multiple identifiers, invalid mode) |
| `KeyError` | Missing account or required env var |
| `httpx.HTTPStatusError` | 4xx/5xx after retries; caller checks `e.response.status_code` |
| `RuntimeError` | Token acquisition failed (include MSAL error code, never password) |
| `FileNotFoundError` | `upload` local_path missing |

Test that secrets never leak into `repr`, `str`, exception messages, logs.

---

## 14. Logging

stdlib `logging`, logger name `gex_msgraph`.

| Level | Content |
|---|---|
| DEBUG | Each Graph API URL + status code |
| INFO | Token acquisitions, retry attempts |
| WARNING | Failed retries, files skipped in `"warn"` mode |
| ERROR | Terminal failures before raise |

Library never calls `logging.basicConfig()` — caller configures.

---

## 15. Code Conventions

- `from __future__ import annotations` at top of every module
- Strict type hints on every public function
- `Literal[...]` for enum-like string params
- One-line docstring on public methods only
- Comments only for non-obvious logic
- Constants UPPER_SNAKE; classes PascalCase; functions snake_case
- Private modules underscore-prefixed (`_core.py`, `_files.py`)
- Private functions/classes underscore-prefixed (`_TokenProvider`, `_compute_backoff`)

---

## 16. Tests

Three files, all using `respx` to mock httpx. No live network calls.

### `tests/conftest.py` — fixtures

| Fixture | Purpose |
|---|---|
| `env_vars` | Monkeypatch `MS_*` env vars for a fake `das_u1` account |
| `mock_token` | Patch `_TokenProvider.get_token` to return `"fake-token"` |
| `mock_graph` | `respx.MockRouter` with common Graph responses pre-stubbed |

### `tests/test_files.py` — pure unit tests

- `validate_identifier`: zero, one, multiple identifiers
- `encode_share_url`: known URL → known base64 output
- `build_resolution_url`: each kind builds correct path
- `match_sheet_name`: exact, ci, glob modes
- Glob with multiple matches picks first
- Missing match returns None
- `parse_drive_item`: valid Graph dict → correct `FileItem`

### `tests/test_client.py` — integration with respx

- Construction reads env correctly; missing env var raises `KeyError`
- `read_excel` works with each of `item_path`, `share_url`, `item_id`
- `read_excel` with both raises `ValueError`
- `download` returns expected bytes
- Retry: 429 with `Retry-After: 2` → waits, retries, succeeds
- Retry: 503 → exponential backoff, 3 attempts, eventually raises
- No retry on 401, 403, 404
- `walk` follows `@odata.nextLink` across 2 pages
- Semaphore caps concurrent `_request` calls
- Token acquired only once across multiple calls
- `read_excel_many` `on_error="warn"`: one bad file → 4 succeed, 1 warning logged
- `read_excel_many` `on_missing_sheet="skip"`: silent skip
- `read_excel_many` empty input → empty DataFrame
- `send_mail` constructs correct payload (recipients, subject, body)
- `send_teams_message` constructs correct payload
- Sync wrappers work from sync test
- `repr(client)` doesn't contain password
- Exception from token failure doesn't contain password
- Captured log records don't contain password or token

Target: ~30 tests. Runs in <5 seconds.

---

## 17. Documentation

Five files. Each has a defined audience and purpose — don't merge them.

### `README.md` — landing page (~1 page)

For someone who just opened the repo on GitHub.

1. **What** — 2 sentences: what this library does
2. **Status** — version, Python required, license
3. **Quick example** — 5 lines showing the minimal usage
4. **Links** — pointers to `USAGE.md` (setup), `CHANGELOG.md` (versions), `AGENTS.md` (contributing)

Short and scannable. Don't repeat what's in `USAGE.md`.

### `USAGE.md` — end-user guide (~4 pages)

The primary doc people read when they want to USE the library. Must cover the full Windows-laptop + Ubuntu-server + GitHub workflow per §18.

Structure:

1. **Prerequisites** — Python 3.11+, uv installed, Git + SSH key set up for the GitHub org
2. **Get credentials from IT** — what to ask for, expected fields back
3. **Install in your project** — `uv add` from Git, both Windows (PowerShell) and Ubuntu (bash) commands
4. **Configure `.env`** — full schema with comments, security notes (`chmod 600`, `.gitignore`)
5. **Find SharePoint IDs (one-time)** — Graph Explorer walkthrough with screenshots described in text
6. **First call** — minimal working example, both async and `_sync` styles
7. **Recipes** — one example per public method (10 short snippets):
   - Read single Excel by path
   - Read single Excel by share link
   - Read CSV
   - Download raw bytes (custom parsing)
   - List files in a folder
   - Walk recursively with glob filter
   - Bulk read many Excel files into one DataFrame
   - Upload a file
   - Send mail
   - Send Teams message
8. **Multi-account** — instantiate two `GraphClient`s, parallel reads
9. **FastAPI integration** — `lifespan` context manager pattern, full working endpoint
10. **Prefect integration** — async flow example with `walk` + `read_excel_many` + `send_teams_message`
11. **Notebook usage** — when to use `_sync` wrappers vs `await`
12. **Deployment to Ubuntu server** — `.env` location, `chmod 600`, Docker `env_file:` mount, systemd service env vars
13. **Updating the library** — bumping versions in `pyproject.toml`, `uv sync`
14. **Troubleshooting** — table of common errors and fixes (see §18)

`USAGE.md` is the most important file — most readers go here first, not README.

### `AGENTS.md` — for AI coding assistants (~2 pages)

For Claude Code / Cline / Cursor working on the library itself.

1. **Architecture** — 1 paragraph
2. **Module map** — `_core.py` / `_files.py` / `__init__.py`
3. **Conventions** — refer to §15
4. **Don't:**
   - Add deps without discussion
   - Add custom exceptions
   - Put auth logic outside `_TokenProvider`
   - Put file resolution outside `_files.py`
   - Log secrets
   - Add a public method without test + USAGE.md recipe
5. **Adding a method:** signature in `_core.py` → impl → test in `test_client.py` → recipe in `USAGE.md`
6. **Adding an identifier type:** touches `_files.py` only
7. **Release process:** bump `__version__`, update `CHANGELOG.md`, tag `v0.X.Y`, push tag

### `CHANGELOG.md`

Keep-a-Changelog format. Initial entry for `0.1.0` listing all features in §8.

### `.env.example`

Annotated template — every var marked `# REQUIRED` or `# OPTIONAL — default X`. Same content used in `USAGE.md` §4.

---

## 18. End-User Workflow — Windows laptop + Ubuntu server + GitHub

This section describes what `USAGE.md` must teach. Claude Code uses this as the source of truth when writing `USAGE.md`.

### 18.1 Prerequisites — what the user has installed

The library supports both `pip` (stdlib, comes with Python) and `uv` (faster, modern alternative). Use whichever your team prefers — `USAGE.md` shows both.

| Tool | Windows | Ubuntu | Required? |
|---|---|---|---|
| Python 3.11+ | `winget install Python.Python.3.11` | `apt install python3.11 python3.11-venv` | yes |
| pip | included with Python | included with Python | yes (always available) |
| uv (optional) | `powershell -c "irm https://astral.sh/uv/install.ps1 \| iex"` | `curl -LsSf https://astral.sh/uv/install.sh \| sh` | no — only if your project uses it |
| Git | `winget install Git.Git` | `apt install git` | yes |
| SSH key for GitHub | `ssh-keygen -t ed25519`, paste pubkey to GitHub | same | yes |
| VS Code (recommended) | `winget install Microsoft.VisualStudioCode` | snap or `.deb` | no |

`USAGE.md` must give exact PowerShell and bash commands for both pip and uv flows.

### 18.2 Install in a consuming project

**Two supported workflows** — `USAGE.md` shows both. Pick whichever your team uses.

#### Workflow A: pip + venv (works everywhere, no extra tool)

**Windows (PowerShell):**
```powershell
cd C:\projects\my-analytics-project
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

**Ubuntu (bash):**
```bash
cd ~/projects/my-analytics-project
python3.11 -m venv .venv
source .venv/bin/activate
pip install "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

To pin in `requirements.txt`:
```
gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0
```

Then `pip install -r requirements.txt`. For reproducibility, also commit `pip freeze > requirements.lock`.

#### Workflow B: uv (faster, single tool)

**Windows (PowerShell) and Ubuntu (bash) — same command:**
```bash
cd path/to/my-analytics-project
uv add "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.1.0"
```

`uv` creates the venv automatically. The `uv.lock` file is auto-generated and committed.

#### Comparison

| Aspect | pip + venv | uv |
|---|---|---|
| Venv creation | manual (`python -m venv`) | automatic |
| Speed of install | ~30s | ~3s |
| Lock file | manual (`pip freeze`) | automatic (`uv.lock`) |
| Available everywhere | yes | needs install |
| Familiar to most Python devs | yes | newer |

The pin `@v0.1.0` is mandatory in both workflows for reproducibility.

### 18.3 `.env` placement and security

| Environment | Location | Security |
|---|---|---|
| Windows dev | `C:\projects\<project>\.env` | NTFS permissions: only your user has access. Confirm with `icacls .env` |
| Ubuntu server | `/opt/<project>/.env` or `/etc/<project>/.env` | `chmod 600 .env && chown <service-user>:<service-user> .env` |
| Docker on Ubuntu | Mount via `env_file:` in `docker-compose.yml`, source file outside container | Same Ubuntu file permissions |

**Always:** `.env` is in `.gitignore`. Never committed. `USAGE.md` shows how to verify (`git check-ignore -v .env` should print the gitignore rule).

### 18.4 GitHub workflow for the consuming project

The user's project (the one using `gex-msgraph`) lives in its own repo. `USAGE.md` shows the recommended pattern for both workflows:

**Workflow A — pip + venv:**
```
my-analytics-project/
├── pyproject.toml         # has gex-msgraph as dependency
├── requirements.txt       # COMMITTED — has the pinned version
├── requirements.lock      # COMMITTED — `pip freeze` output, exact versions
├── .env.example           # COMMITTED — template, no secrets
├── .env                   # GITIGNORED — has real secrets
├── .gitignore             # includes .env, .venv/
├── README.md
└── src/
    └── ...
```

**Workflow B — uv:**
```
my-analytics-project/
├── pyproject.toml         # has gex-msgraph as dependency
├── uv.lock                # COMMITTED — pins exact versions
├── .env.example           # COMMITTED — template, no secrets
├── .env                   # GITIGNORED — has real secrets
├── .gitignore             # includes .env, .venv/
├── README.md
└── src/
    └── ...
```

**Always commit the lock file** (whichever tool). When teammate clones and installs, they get exact same versions.

### 18.5 Updating gex-msgraph in a consuming project

**Workflow A — pip:**
```bash
# Edit requirements.txt to bump @v0.1.0 → @v0.2.0
pip install -r requirements.txt --upgrade
pip freeze > requirements.lock
git commit requirements.txt requirements.lock -m "bump gex-msgraph to v0.2.0"
```

**Workflow B — uv:**
```bash
uv add "gex-msgraph @ git+ssh://git@github.com/companyg/gex-msgraph.git@v0.2.0"
git commit pyproject.toml uv.lock -m "bump gex-msgraph to v0.2.0"
```

### 18.6 Deploy to Ubuntu server

Three patterns covered in `USAGE.md`. Each shows both pip and uv commands:

**A. Bare metal / venv** — install in a venv on the server, `.env` in project dir.

pip:
```bash
python3.11 -m venv /opt/myapp/.venv
/opt/myapp/.venv/bin/pip install -r requirements.txt
/opt/myapp/.venv/bin/python main.py
```

uv:
```bash
cd /opt/myapp && uv sync
uv run python main.py
```

**B. Docker** — Dockerfile builds the image, `docker-compose.yml` mounts `.env` via `env_file:`. Image doesn't contain secrets.

Dockerfile snippet (pip — works without uv installed in the image):
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
CMD ["python", "main.py"]
```

Or uv-based image (faster builds, smaller layers — install uv in the Dockerfile):
```dockerfile
FROM python:3.11-slim
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv
WORKDIR /app
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev
COPY . .
CMD ["uv", "run", "python", "main.py"]
```

**C. systemd service** — `EnvironmentFile=/etc/myapp/.env` directive in unit file. Permissions enforced by systemd. ExecStart points at the venv's python:

```ini
[Service]
EnvironmentFile=/etc/myapp/.env
ExecStart=/opt/myapp/.venv/bin/python /opt/myapp/main.py
User=myapp
Group=myapp
```

`USAGE.md` shows a working snippet for each.

### 18.7 GitHub workflow for the gex-msgraph repo itself

This is the dev workflow for maintaining the library. Goes in `AGENTS.md`, not `USAGE.md`. Both pip and uv supported:

| Action | pip + venv | uv |
|---|---|---|
| Clone | `git clone git@github.com:companyg/gex-msgraph.git` | same |
| Setup dev env | `python -m venv .venv && source .venv/bin/activate && pip install -e ".[dev]"` | `uv sync --extra dev` |
| Test | `pytest` (after activating venv) | `uv run pytest` |
| Lint | `ruff check && mypy src` | `uv run ruff check && uv run mypy src` |
| Release | bump `__version__` in `_core.py`, update `CHANGELOG.md`, `git tag v0.X.Y && git push --tags` | same |

The library itself doesn't commit a lock file — it's a library, not an application. Consumers' lock files pin versions.

### 18.8 CI (`.github/workflows/ci.yml`)

Runs on every push and PR. Uses `uv` for speed — uv installs in 1 second and works without modification on GitHub-hosted runners. The library itself is pip-installable (see §5), this is just the CI tool choice:

```yaml
name: ci
on: [push, pull_request]
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v3
      - run: uv venv
      - run: uv pip install -e ".[dev]"
      - run: uv run ruff check
      - run: uv run mypy src
      - run: uv run pytest
```

Note: `uv pip install -e ".[dev]"` is used (not `uv sync`) to keep the install commands identical to what a pip user would run. Library remains 100% pip-compatible.

Must pass before merging. Required by `AGENTS.md` rules.

### 18.9 Troubleshooting (goes in `USAGE.md`)

`USAGE.md` includes this troubleshooting table:

| Symptom | Cause | Fix |
|---|---|---|
| `KeyError: MS_DAS_U1_CLIENT_ID` | `.env` not loaded or var missing | Check `load_dotenv()` is called; verify `.env` has the var |
| `RuntimeError: Token acquisition failed` | Wrong creds, MFA, account locked | Test login at portal.office.com with same creds |
| `httpx.HTTPStatusError: 403` on read | Account lacks permission to that SharePoint site | Ask IT to grant the service account access to the site |
| `httpx.HTTPStatusError: 404` on item_path | Wrong path or wrong drive | Use `walk()` first to list real paths |
| `read_excel_sync` raises "asyncio.run cannot be called" | Inside Jupyter or async context | Use `await g.read_excel(...)` instead |
| `ssh: Permission denied (publickey)` on install | SSH key not set up for GitHub | Run `ssh -T git@github.com`, fix per GitHub SSH docs |
| `pip install` fails with "could not find a version" | `git+ssh://` URLs need git installed | `apt install git` or `winget install Git.Git` |
| `pip install` hangs on git | SSH agent not running | Windows: `Start-Service ssh-agent`. Ubuntu: `eval $(ssh-agent) && ssh-add` |
| Slow reads on many files | Sequential, not concurrent | Use `read_excel_many()` or `asyncio.gather()` |
| Windows: `pip install` fails with path errors | Windows path length limit | Enable long paths: `git config --system core.longpaths true` |
| `ModuleNotFoundError: gex_msgraph` after install | Wrong venv active | Check with `pip show gex-msgraph`; activate the right venv |

---

## 19. Out of Scope (do NOT build)

- Webhooks / change notifications
- Teams meeting creation
- SharePoint search API
- Calendar / contacts
- Batch endpoint
- App-only auth (client credentials)
- Web URL parsing (only share URLs supported)
- Email attachments — v0.2 candidate
- HTML mail body — v0.2 candidate
- `download_to_file()` — caller writes bytes if needed
- `find_site()` / `list_drives()` helpers — use Graph Explorer in browser to find IDs once
- Per-call `drive_id` / `site_id` overrides — set defaults in `.env`
- File deletion / move
- Permission management
- Logging configuration

---

## 20. Build Order

1. **Skeleton** — `pyproject.toml`, `src/gex_msgraph/`, install works, empty modules import
2. **Auth + core init** — `_TokenProvider`, `GraphClient.__init__`, env loading, httpx client, `close()`, context manager
3. **`_files.py`** — `FileItem`, `validate_identifier`, `encode_share_url`, `build_resolution_url`, `match_sheet_name`, `parse_drive_item`
4. **`_request` helper** — retry + semaphore + token injection in one place
5. **`download()` + `read_excel()` + `read_csv()`** — single-file reads end-to-end
6. **`walk()` + `list_files()`** — discovery with pagination
7. **`read_excel_many()`** — bulk + error modes
8. **`upload()`** — write path
9. **`send_mail()` + `send_teams_message()`**
10. **Sync wrappers**
11. **Tests written alongside steps 2–10**
12. **CI** — `.github/workflows/ci.yml` running ruff + mypy + pytest
13. **Docs in this order:** `.env.example` → `CHANGELOG.md` → `README.md` → `USAGE.md` → `AGENTS.md`

After step 5, library is already useful for daily work. Docs are last because they reference the actual built API.

---

## 21. Definition of Done

- [ ] All 11 async methods + 3 sync wrappers implemented per §8
- [ ] `_TokenProvider`, `FileItem`, all `_files.py` helpers per §4
- [ ] All 3 file identifiers resolve correctly per §9
- [ ] Performance mechanisms in §10 implemented and tested
- [ ] `read_excel_many` honors all sheet match + error modes per §11
- [ ] No secrets in logs, exceptions, or repr (test enforces)
- [ ] `pytest` green, ~30 tests, <5s runtime
- [ ] `mypy --strict` green
- [ ] `ruff check` green
- [ ] CI workflow at `.github/workflows/ci.yml` runs all three checks
- [ ] All five doc files written per §17: `README.md`, `USAGE.md`, `AGENTS.md`, `CHANGELOG.md`, `.env.example`
- [ ] `USAGE.md` covers Windows + Ubuntu install commands per §18.1–18.2
- [ ] `USAGE.md` covers `.env` placement and security per §18.3
- [ ] `USAGE.md` covers Ubuntu deployment patterns per §18.6
- [ ] `USAGE.md` includes the troubleshooting table per §18.9
- [ ] Installs cleanly via `pip install` AND `uv add` on both Windows and Ubuntu
- [ ] `from gex_msgraph import GraphClient` works after install
- [ ] `USAGE.md` covers both pip and uv commands per §18.2 / §18.5 / §18.7
- [ ] Quickstart example from `USAGE.md` runs against a real M365 tenant

---

## 22. Defaults — don't ask, just pick

| Question | Default |
|---|---|
| Sync wrapper implementation | `asyncio.run()` — document the no-running-loop limitation |
| HTML vs text mail in v1 | Text only |
| Attachments in v1 | Not supported |
| Empty bulk-read result | Return empty DataFrame, not raise |
| Default request timeout | 30 seconds (overridable via env) |
| Default semaphore | 10 (overridable via env or `max_concurrent` arg) |
| Initial version | `0.1.0` |
| Sync vs async MSAL | MSAL is sync; wrap in `asyncio.to_thread()` when called from async context |

If something else is genuinely ambiguous, ask. Otherwise pick the simpler option.

---

**End of spec.**
