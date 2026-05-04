from __future__ import annotations

import asyncio
import logging
import os
from collections.abc import AsyncIterator
from typing import TYPE_CHECKING, Any, Literal, cast

if TYPE_CHECKING:
    import pandas as pd
    from gex_msgraph._files import TreeNode

import httpx
import msal

from gex_msgraph._files import FileItem, build_resolution_url, validate_identifier

logger = logging.getLogger("gex_msgraph")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"
_DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"]
_DEFAULT_TIMEOUT = 30.0
_DEFAULT_MAX_CONCURRENT = 10
_DEFAULT_MAX_RETRIES = 3
_BACKOFF_CAP = 30.0


def _load_account_env(name: str) -> dict[str, str]:
    """Read MS_<NAME>_* env vars, return as dict."""
    prefix = f"MS_{name.upper()}_"

    res: dict[str, str] = {}
    for req in ["CLIENT_ID", "CLIENT_SECRET", "TENANT_ID", "USERNAME", "PASSWORD"]:
        val = os.environ.get(prefix + req)
        if val:
            res[req.lower()] = val

    res["default_site_id"] = os.environ.get(prefix + "DEFAULT_SITE_ID", "")
    res["default_drive_id"] = os.environ.get(prefix + "DEFAULT_DRIVE_ID", "")
    res["max_concurrent"] = os.environ.get(prefix + "MAX_CONCURRENT", str(_DEFAULT_MAX_CONCURRENT))
    res["request_timeout"] = os.environ.get(prefix + "REQUEST_TIMEOUT", str(_DEFAULT_TIMEOUT))

    return res


def _compute_backoff(attempt: int, retry_after: str | None) -> float:
    """Compute backoff seconds."""
    if retry_after:
        try:
            return float(retry_after)
        except ValueError:
            pass
    return min(float(2**attempt), _BACKOFF_CAP)


class _TokenProvider:
    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        username: str,
        password: str,
    ) -> None:
        self._username = username
        self._password = password
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=_AUTHORITY_TEMPLATE.format(tenant_id=tenant_id),
        )

    def get_token(self) -> str:
        """Returns access token. Synchronous (MSAL is sync). Cached automatically."""
        accounts = self._app.get_accounts(username=self._username)
        if accounts:
            result = self._app.acquire_token_silent(_DEFAULT_SCOPE, account=accounts[0])
            if result and "access_token" in result:
                return str(result["access_token"])

        result = self._app.acquire_token_by_username_password(
            username=self._username,
            password=self._password,
            scopes=_DEFAULT_SCOPE,
        )
        if "access_token" in result:
            return str(result["access_token"])

        err = result.get("error_description") or result.get("error") or "Unknown error"
        raise RuntimeError(f"Token acquisition failed: {err}")


class GraphClient:
    def __init__(
        self,
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
    ) -> None:
        self.account = account or "custom"
        cfg = _load_account_env(account) if account else {}

        resolved_client_id = client_id or cfg.get("client_id")
        resolved_client_secret = client_secret or cfg.get("client_secret")
        resolved_tenant_id = tenant_id or cfg.get("tenant_id")
        resolved_username = username or cfg.get("username")
        resolved_password = password or cfg.get("password")

        if not resolved_client_id:
            raise KeyError(f"MS_{self.account.upper()}_CLIENT_ID" if account else "client_id")
        if not resolved_client_secret:
            raise KeyError(f"MS_{self.account.upper()}_CLIENT_SECRET" if account else "client_secret")
        if not resolved_tenant_id:
            raise KeyError(f"MS_{self.account.upper()}_TENANT_ID" if account else "tenant_id")
        if not resolved_username:
            raise KeyError(f"MS_{self.account.upper()}_USERNAME" if account else "username")
        if not resolved_password:
            raise KeyError(f"MS_{self.account.upper()}_PASSWORD" if account else "password")

        self._provider = _TokenProvider(
            client_id=resolved_client_id,
            client_secret=resolved_client_secret,
            tenant_id=resolved_tenant_id,
            username=resolved_username,
            password=resolved_password,
        )
        self._default_drive_id = default_drive_id if default_drive_id is not None else cfg.get("default_drive_id", "")

        timeout = request_timeout if request_timeout is not None else float(cfg.get("request_timeout", _DEFAULT_TIMEOUT))
        concurrent = max_concurrent if max_concurrent is not None else int(cfg.get("max_concurrent", _DEFAULT_MAX_CONCURRENT))

        self._client = httpx.AsyncClient(timeout=timeout)
        self._semaphore = asyncio.Semaphore(concurrent)

    async def close(self) -> None:
        """Close the underlying httpx client. Safe to call multiple times."""
        await self._client.aclose()

    async def __aenter__(self) -> "GraphClient":
        return self

    async def __aexit__(self, *exc: Any) -> None:
        await self.close()

    def _drive_root(self) -> str:
        """URL fragment for the drive scope, e.g. ``/me/drive`` or ``/drives/{id}``."""
        return f"/drives/{self._default_drive_id}" if self._default_drive_id else "/me/drive"

    async def _request(self, method: str, url: str, **kw: Any) -> httpx.Response:
        if not url.startswith("http"):
            url = _GRAPH_BASE + url

        attempt = 0
        while True:
            token = await asyncio.to_thread(self._provider.get_token)

            headers = kw.pop("headers", {})
            headers["Authorization"] = f"Bearer {token}"
            kw["headers"] = headers

            async with self._semaphore:
                logger.debug("Graph request %s %s", method, url)
                response = await self._client.request(method, url, **kw)

            status = response.status_code
            logger.debug("Graph response %s %s -> %s", method, url, status)

            if status < 400:
                return response

            if status == 429 or status >= 500:
                if attempt >= _DEFAULT_MAX_RETRIES:
                    logger.error("Graph request failed after %s attempts", attempt)
                    response.raise_for_status()
                    raise AssertionError("unreachable")

                backoff = _compute_backoff(attempt, response.headers.get("Retry-After"))
                logger.info("Retrying %s %s in %.1fs (attempt %s)", method, url, backoff, attempt + 1)
                await asyncio.sleep(backoff)
                attempt += 1
                continue

            logger.error("Graph request failed %s: %s", status, response.text)
            response.raise_for_status()
            raise AssertionError("unreachable")

    async def _get_download_url(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> str:
        kind = validate_identifier(item_path, share_url, item_id)
        value = item_path if kind == "path" else share_url if kind == "share" else item_id
        assert value is not None

        url = build_resolution_url(kind, value, self._drive_root())
        resp = await self._request("GET", url)
        data = resp.json()
        download_url = data.get("@microsoft.graph.downloadUrl")
        if not download_url:
            raise RuntimeError(f"Could not resolve download URL for item: {value}")
        return str(download_url)

    async def _stream_to_bytes(self, url: str) -> bytes:
        async with self._semaphore:
            logger.debug("Streaming %s", url)
            resp = await self._client.get(url)
            resp.raise_for_status()
            return await resp.aread()

    async def _iter_paginated(self, url: str) -> AsyncIterator[dict[str, Any]]:
        while url:
            resp = await self._request("GET", url)
            data = resp.json()
            for item in data.get("value", []):
                yield item
            url = data.get("@odata.nextLink", "")

    async def read_excel(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        sheet: str | int = 0,
        **read_excel_kwargs: Any,
    ) -> "pd.DataFrame":
        """Read an Excel file directly into a DataFrame. Streams via BytesIO."""
        import io
        import pandas as pd

        data = await self.download(item_path=item_path, share_url=share_url, item_id=item_id)
        buf = io.BytesIO(data)
        return pd.read_excel(buf, sheet_name=sheet, **read_excel_kwargs)

    async def read_csv(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        **read_csv_kwargs: Any,
    ) -> "pd.DataFrame":
        """Read a CSV file directly into a DataFrame. Streams via BytesIO."""
        import io
        import pandas as pd

        data = await self.download(item_path=item_path, share_url=share_url, item_id=item_id)
        buf = io.BytesIO(data)
        return pd.read_csv(buf, **read_csv_kwargs)

    async def download(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> bytes:
        """Return raw file bytes. Use for non-pandas parsing."""
        url = await self._get_download_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        return await self._stream_to_bytes(url)

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
    ) -> "pd.DataFrame":
        """Read many Excel files concurrently and concat into one DataFrame."""
        import io
        import pandas as pd
        from gex_msgraph._files import match_sheet_name

        async def _process_one(path: str) -> pd.DataFrame | None:
            # Step 1: download + open. Failures here are governed by `on_error`.
            try:
                data = await self.download(item_path=path)
                buf = io.BytesIO(data)
                xls = pd.ExcelFile(buf)
                matched = match_sheet_name(xls.sheet_names, sheet, sheet_match)
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to read %s: %s", path, e)
                return None

            # Step 2: missing-sheet decision is independent of `on_error`.
            if matched is None:
                if on_missing_sheet == "raise":
                    raise ValueError(f"Sheet '{sheet}' not found in {path}")
                if on_missing_sheet == "warn":
                    logger.warning("Sheet '%s' not found in %s", sheet, path)
                return None

            # Step 3: parse the matched sheet. Parse failures fall under `on_error`.
            try:
                df = pd.read_excel(xls, sheet_name=matched, **read_excel_kwargs)
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to read %s: %s", path, e)
                return None

            if add_source_column:
                df["_source"] = path
            return df

        sem = asyncio.Semaphore(max_concurrent) if max_concurrent else None
        
        async def _bounded_process(path: str) -> pd.DataFrame | None:
            if sem:
                async with sem:
                    return await _process_one(path)
            return await _process_one(path)

        tasks = [_bounded_process(p) for p in paths]
        results = await asyncio.gather(*tasks)
        
        valid_dfs = [df for df in results if df is not None]
        if not valid_dfs:
            return pd.DataFrame()
        return pd.concat(valid_dfs, ignore_index=True, sort=False)

    async def walk(
        self,
        folder_path: str = "",
        *,
        pattern: str | None = None,
        recursive: bool = True,
    ) -> list[FileItem]:
        """List files (not folders) under a folder. Recursive by default."""
        import fnmatch
        from gex_msgraph._files import parse_drive_item

        drive_root = self._drive_root()
        clean_path = folder_path.lstrip("/")

        if clean_path:
            base_url = f"{drive_root}/root:/{clean_path}:/children"
        else:
            base_url = f"{drive_root}/root/children"

        results: list[FileItem] = []

        async def _explore(url: str, current_path: str) -> None:
            subfolders: list[tuple[str, str]] = []

            async for item in self._iter_paginated(url):
                file_item = parse_drive_item(item, parent_path=current_path)

                if file_item.is_folder:
                    if recursive:
                        folder_url = f"{drive_root}/items/{file_item.id}/children"
                        subfolders.append((folder_url, file_item.path))
                else:
                    if pattern is None or fnmatch.fnmatch(file_item.name, pattern):
                        results.append(file_item)

            tasks = [_explore(sub_url, sub_path) for sub_url, sub_path in subfolders]
            if tasks:
                await asyncio.gather(*tasks)

        await _explore(base_url, clean_path)
        return results

    async def list_files(self, folder_path: str = "") -> list[FileItem]:
        """List immediate children (files and folders) of one folder. Non-recursive."""
        from gex_msgraph._files import parse_drive_item

        drive_root = self._drive_root()
        clean_path = folder_path.lstrip("/")

        if clean_path:
            url = f"{drive_root}/root:/{clean_path}:/children"
        else:
            url = f"{drive_root}/root/children"

        results: list[FileItem] = []
        async for item in self._iter_paginated(url):
            results.append(parse_drive_item(item, parent_path=clean_path))

        return results

    async def get_metadata(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> FileItem:
        """Fetch metadata for a single item without downloading its content."""
        from gex_msgraph._files import parse_drive_item
        kind = validate_identifier(item_path, share_url, item_id)
        value = item_path if kind == "path" else share_url if kind == "share" else item_id
        assert value is not None

        url = build_resolution_url(kind, value, self._drive_root())
        resp = await self._request("GET", url)
        return parse_drive_item(resp.json())

    async def delete_file(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> None:
        """Delete an item."""
        kind = validate_identifier(item_path, share_url, item_id)
        value = item_path if kind == "path" else share_url if kind == "share" else item_id
        assert value is not None

        url = build_resolution_url(kind, value, self._drive_root())
        await self._request("DELETE", url)

    async def move_file(
        self,
        source_path: str,
        dest_folder_path: str | None = None,
        new_name: str | None = None,
    ) -> FileItem:
        """Move or rename a file using paths."""
        from gex_msgraph._files import parse_drive_item
        import urllib.parse

        drive_root = self._drive_root()
        source_clean = source_path.lstrip("/")

        encoded_source = urllib.parse.quote(source_clean)
        url = f"{drive_root}/root:/{encoded_source}"

        payload: dict[str, Any] = {}
        if dest_folder_path is not None:
            dest_clean = dest_folder_path.lstrip("/")
            payload["parentReference"] = {
                "path": f"{drive_root}/root:/{dest_clean}" if dest_clean else f"{drive_root}/root"
            }
        if new_name is not None:
            payload["name"] = new_name

        if not payload:
            raise ValueError("Must provide dest_folder_path or new_name")

        resp = await self._request("PATCH", url, json=payload)
        return parse_drive_item(resp.json())

    async def create_folder(self, folder_path: str) -> FileItem:
        """Create a new folder."""
        from gex_msgraph._files import parse_drive_item
        import urllib.parse

        drive_root = self._drive_root()
        clean_path = folder_path.lstrip("/")

        if "/" in clean_path:
            parent, name = clean_path.rsplit("/", 1)
            encoded_parent = urllib.parse.quote(parent)
            url = f"{drive_root}/root:/{encoded_parent}:/children"
        else:
            name = clean_path
            url = f"{drive_root}/root/children"
            
        payload = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        resp = await self._request("POST", url, json=payload)
        return parse_drive_item(resp.json())

    async def list_excel_sheets(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> list[str]:
        """List all worksheet names in an Excel file."""
        kind = validate_identifier(item_path, share_url, item_id)
        value = item_path if kind == "path" else share_url if kind == "share" else item_id
        assert value is not None

        base = build_resolution_url(kind, value, self._drive_root())
        # Path-addressed items use the ":/workbook/..." colon syntax;
        # id- and share-addressed items concatenate directly.
        suffix = ":/workbook/worksheets" if kind == "path" else "/workbook/worksheets"
        resp = await self._request("GET", base + suffix)
        return [str(sheet["name"]) for sheet in resp.json().get("value", [])]

    async def get_folder_tree(self, folder_path: str = "") -> "TreeNode":
        """Return a recursive tree representation of a folder."""
        from gex_msgraph._files import TreeNode

        # Drive root has no addressable metadata in the path-resolution form,
        # so leave `item` as None when listing from the root.
        root_item = await self.get_metadata(item_path=folder_path) if folder_path else None
        root_node = TreeNode(item=root_item, children=[])

        async def _build_tree(path: str, node: TreeNode) -> None:
            children_items = await self.list_files(path)
            for child in children_items:
                child_node = TreeNode(item=child, children=[])
                node.children.append(child_node)
                if child.is_folder:
                    await _build_tree(child.path, child_node)

        await _build_tree(folder_path, root_node)
        return root_node

    async def upload(self, local_path: str | os.PathLike[str], remote_path: str) -> dict[str, Any]:
        """Upload a local file to remote_path. Returns the Graph driveItem dict."""
        import urllib.parse
        from pathlib import Path
        
        local_p = Path(local_path)
        if not local_p.is_file():
            raise FileNotFoundError(f"File not found: {local_path}")
            
        drive_root = self._drive_root()
        clean_remote = remote_path.lstrip("/")
        encoded_path = urllib.parse.quote(clean_remote)

        url = f"{drive_root}/root:/{encoded_path}:/content"
        
        with open(local_p, "rb") as f:
            data = f.read()
            
        resp = await self._request("PUT", url, content=data, headers={"Content-Type": "application/octet-stream"})
        return cast(dict[str, Any], resp.json())

    async def send_mail(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        *,
        cc: str | list[str] | None = None,
    ) -> None:
        """Send a plain-text email from the account's mailbox."""
        def _make_recipients(addrs: str | list[str]) -> list[dict[str, Any]]:
            if isinstance(addrs, str):
                addrs = [addrs]
            return [{"emailAddress": {"address": a}} for a in addrs]
            
        payload: dict[str, Any] = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body,
                },
                "toRecipients": _make_recipients(to),
            },
            "saveToSentItems": "true"
        }
        
        if cc:
            payload["message"]["ccRecipients"] = _make_recipients(cc)
            
        await self._request("POST", "/me/sendMail", json=payload)

    async def send_teams_message(
        self,
        team_id: str,
        channel_id: str,
        text: str,
    ) -> None:
        """Post a plain-text message to a Teams channel."""
        payload = {
            "body": {
                "contentType": "Text",
                "content": text
            }
        }
        url = f"/teams/{team_id}/channels/{channel_id}/messages"
        await self._request("POST", url, json=payload)

    async def get_teams_messages(
        self,
        team_id: str,
        channel_id: str,
        limit: int = 10,
    ) -> list[dict[str, Any]]:
        """Fetch incoming messages from a Teams channel."""
        url = f"/teams/{team_id}/channels/{channel_id}/messages?$top={limit}"
        resp = await self._request("GET", url)
        return cast(list[dict[str, Any]], resp.json().get("value", []))

    def read_excel_sync(self, **kw: Any) -> "pd.DataFrame":
        return asyncio.run(self.read_excel(**kw))

    def read_csv_sync(self, **kw: Any) -> "pd.DataFrame":
        return asyncio.run(self.read_csv(**kw))

    def download_sync(self, **kw: Any) -> bytes:
        return asyncio.run(self.download(**kw))
