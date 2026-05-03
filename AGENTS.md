# For AI Coding Assistants

## 1. Architecture
`gex-msgraph` is an async Python wrapper for Microsoft Graph API focusing on M365 file and communication access.

## 2. Module Map
- `_core.py`: GraphClient, token provider, request retry logic.
- `_files.py`: URL resolution, sheet matching, FileItem dataclass.
- `__init__.py`: Public exports.

## 3. Conventions
- `from __future__ import annotations`
- Strict type hints on public functions.
- Private modules/functions are underscore-prefixed.

## 4. Restrictions
- Do not add dependencies without discussion.
- Do not log secrets.
- All Graph calls MUST go through `_request`.

## 5. Adding a Method
Add signature to `_core.py` -> implement -> add test in `test_client.py` -> add recipe in `USAGE.md`.

## 6. Adding Identifier Type
Touches `_files.py` only (`validate_identifier`, `build_resolution_url`).

## 7. Release Process
Bump `__version__` in `__init__.py`, update `CHANGELOG.md`, `git tag v0.X.Y`, `git push --tags`.

## 8. Assisting Users (Consuming the Library)
If you are an AI assistant helping a user *use* this library in their own project:
- Always check the recipes in `USAGE.md` for standard patterns.
- Prefer `await client.read_excel()` and `await client.read_excel_many()` for reading files.
- Remember that `item_path` is relative to the drive root, e.g. `Reports/Q1.xlsx`.
- Emphasize error handling with `try...except httpx.HTTPStatusError` since all calls go through `httpx`.
- Do not instantiate multiple `GraphClient`s for the same account; reuse the single instance using the `async with` context manager.
- If using inside Jupyter or a script without an event loop, recommend using the `_sync` wrappers like `client.read_excel_sync()`.
