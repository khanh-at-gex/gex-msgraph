from __future__ import annotations

import asyncio
import base64
import logging
import mimetypes
import os
import threading
from collections.abc import AsyncIterator, Awaitable, Callable, Coroutine
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal, TypeVar, cast, overload

if TYPE_CHECKING:
    import pandas as pd
    from gex_msgraph._files import TreeNode

import httpx
import msal

from gex_msgraph._exceptions import (
    GraphAuthenticationError,
    GraphSyncInLoopError,
    NotFoundError,
    raise_graph_error,
)
from gex_msgraph._files import (
    FileItem,
    build_resolution_url,
    encode_drive_path,
    validate_identifier,
)
from gex_msgraph._spreadsheetml import prepare_excel_bytes

logger = logging.getLogger("gex_msgraph")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"
_DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"]
_DEFAULT_TIMEOUT = 30.0
_DEFAULT_MAX_CONCURRENT = 10
_DEFAULT_MAX_RETRIES = 3
_DEFAULT_EXCEL_ENGINE = "calamine"
_BACKOFF_CAP = 30.0
# Graph's simple PUT :/content caps out around 4 MiB; larger files must go
# through an upload session. Chunks must be multiples of 320 KiB.
_UPLOAD_SESSION_THRESHOLD = 4 * 1024 * 1024
_UPLOAD_CHUNK_SIZE = 10 * 1024 * 1024  # 32 x 320 KiB

_T = TypeVar("_T")
_R = TypeVar("_R")


async def _gather_with_status(
    items: list[_T],
    worker: Callable[[_T], Awaitable[tuple[_R | None, str, str]]],
    *,
    max_concurrent: int | None = None,
) -> list[tuple[_T, _R | None, str, str]]:
    """Run `worker` over `items` concurrently (optionally semaphore-bounded),
    collecting each item's `(result, status, message)` outcome."""
    sem = asyncio.Semaphore(max_concurrent) if max_concurrent else None

    async def _bounded(item: _T) -> tuple[_T, _R | None, str, str]:
        if sem:
            async with sem:
                result, status, msg = await worker(item)
        else:
            result, status, msg = await worker(item)
        return item, result, status, msg

    return await asyncio.gather(*[_bounded(item) for item in items])


def _load_account_env(name: str) -> dict[str, str]:
    """Read MS_<NAME>_* env vars, return as dict."""
    prefix = f"MS_{name.upper()}_"

    keys = ["CLIENT_ID", "CLIENT_SECRET", "TENANT_ID", "USERNAME", "PASSWORD"]
    res = {k.lower(): v for k in keys if (v := os.environ.get(prefix + k))}

    defaults = {
        "DEFAULT_DRIVE_ID": "",
        "MAX_CONCURRENT": str(_DEFAULT_MAX_CONCURRENT),
        "REQUEST_TIMEOUT": str(_DEFAULT_TIMEOUT),
    }
    for k, default_val in defaults.items():
        res[k.lower()] = os.environ.get(prefix + k, default_val)

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
        username: str | None = None,
        password: str | None = None,
    ) -> None:
        self._username = username
        self._password = password
        # Both absent -> app-only (client credentials); both present -> ROPC.
        self._app_only = username is None and password is None
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=_AUTHORITY_TEMPLATE.format(tenant_id=tenant_id),
        )

    def get_token(self) -> str:
        """Returns access token. Synchronous (MSAL is sync). Cached automatically."""
        if self._app_only:
            result = self._app.acquire_token_for_client(scopes=_DEFAULT_SCOPE)
            if result and "access_token" in result:
                return str(result["access_token"])
        else:
            accounts = self._app.get_accounts(username=self._username)
            if accounts:
                result = self._app.acquire_token_silent(
                    _DEFAULT_SCOPE, account=accounts[0]
                )
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
        raise GraphAuthenticationError(f"Token acquisition failed: {err}")


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
        default_drive_id: str | None = None,
        max_concurrent: int | None = None,
        request_timeout: float | None = None,
    ) -> None:
        self.account = account or "custom"
        cfg = _load_account_env(account) if account else {}

        creds: dict[str, str] = {}
        for key, val in [
            ("client_id", client_id),
            ("client_secret", client_secret),
            ("tenant_id", tenant_id),
        ]:
            resolved = val or cfg.get(key)
            if not resolved:
                raise KeyError(
                    f"MS_{self.account.upper()}_{key.upper()}" if account else key
                )
            creds[key] = resolved

        # username/password select the auth flow: both present -> delegated
        # (ROPC), both absent -> app-only (client credentials). One without
        # the other is a configuration error.
        resolved_username = username or cfg.get("username")
        resolved_password = password or cfg.get("password")
        if (resolved_username is None) != (resolved_password is None):
            raise ValueError(
                "Provide both username and password for delegated (ROPC) auth, "
                "or neither for app-only (client credentials) auth."
            )

        self._default_drive_id = (
            default_drive_id
            if default_drive_id is not None
            else cfg.get("default_drive_id", "")
        )

        # App-only tokens have no signed-in user, so the /me/drive default
        # cannot resolve — fail at construction, not on the first request.
        if resolved_username is None and not self._default_drive_id:
            raise ValueError(
                "App-only auth (no username/password) requires default_drive_id "
                f"(or MS_{self.account.upper()}_DEFAULT_DRIVE_ID), because /me/* "
                "endpoints have no meaning without a signed-in user."
            )

        self._provider = _TokenProvider(
            **creds, username=resolved_username, password=resolved_password
        )

        timeout = (
            request_timeout
            if request_timeout is not None
            else float(cfg.get("request_timeout", _DEFAULT_TIMEOUT))
        )
        concurrent = (
            max_concurrent
            if max_concurrent is not None
            else int(cfg.get("max_concurrent", _DEFAULT_MAX_CONCURRENT))
        )

        self._client = httpx.AsyncClient(timeout=timeout)
        self._semaphore = asyncio.Semaphore(concurrent)
        # Lazily-started background loop that hosts all *_sync work, so the
        # httpx client/semaphore stay bound to one loop across sync calls.
        self._sync_loop: asyncio.AbstractEventLoop | None = None
        self._sync_thread: threading.Thread | None = None

    def _run_sync(self, coro: Coroutine[Any, Any, _T]) -> _T:
        try:
            asyncio.get_running_loop()
        except RuntimeError:
            pass
        else:
            coro.close()
            raise GraphSyncInLoopError(
                "GraphClient *_sync methods cannot be called from within a "
                "running event loop (e.g. Jupyter, an async web framework). "
                "Use the async method directly instead, e.g. "
                "`await client.read_excel(...)`."
            )

        if self._sync_loop is None:
            self._sync_loop = asyncio.new_event_loop()
            self._sync_thread = threading.Thread(
                target=self._sync_loop.run_forever, daemon=True
            )
            self._sync_thread.start()

        future = asyncio.run_coroutine_threadsafe(coro, self._sync_loop)
        return future.result()

    def _shutdown_sync_loop(self) -> None:
        if self._sync_loop is not None:
            self._sync_loop.call_soon_threadsafe(self._sync_loop.stop)
            if self._sync_thread is not None:
                self._sync_thread.join(timeout=5)
            self._sync_loop.close()
            self._sync_loop = None
            self._sync_thread = None

    async def close(self) -> None:
        """Close the underlying httpx client. Safe to call multiple times."""
        if self._sync_loop is not None:
            # The httpx client is bound to the background sync loop — closing
            # it from this loop would be a cross-loop call. Delegate to
            # close_sync on a worker thread instead.
            await asyncio.to_thread(self.close_sync)
            return
        await self._client.aclose()

    def close_sync(self) -> None:
        """Close from synchronous code — companion to the *_sync methods."""
        if self._sync_loop is not None:
            future = asyncio.run_coroutine_threadsafe(
                self._client.aclose(), self._sync_loop
            )
            future.result(timeout=5)
            self._shutdown_sync_loop()
        else:
            # Client was never used on the sync loop; a throwaway loop is
            # fine for a plain aclose.
            asyncio.run(self._client.aclose())

    async def __aenter__(self) -> "GraphClient":
        return self

    async def __aexit__(self, *exc: Any) -> None:
        await self.close()

    def _drive_root(self) -> str:
        """URL fragment for the drive scope, e.g. ``/me/drive`` or ``/drives/{id}``."""
        return (
            f"/drives/{self._default_drive_id}"
            if self._default_drive_id
            else "/me/drive"
        )

    def _resolve_item_url(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> tuple[Literal["path", "share", "id"], str]:
        kind = validate_identifier(item_path, share_url, item_id)
        value = (
            item_path if kind == "path" else share_url if kind == "share" else item_id
        )
        assert value is not None

        url = build_resolution_url(kind, value, self._drive_root())
        return kind, url

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
                    raise_graph_error(response, retries_exhausted=True)

                backoff = _compute_backoff(attempt, response.headers.get("Retry-After"))
                logger.info(
                    "Retrying %s %s in %.1fs (attempt %s)",
                    method,
                    url,
                    backoff,
                    attempt + 1,
                )
                await asyncio.sleep(backoff)
                attempt += 1
                continue

            logger.error("Graph request failed %s: %s", status, response.text)
            raise_graph_error(response)

    async def _get_download_url(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> str:
        _, url = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        resp = await self._request("GET", url)
        data = resp.json()
        download_url = data.get("@microsoft.graph.downloadUrl")
        if not download_url:
            raise RuntimeError(f"Could not resolve download URL for item: {item_path or share_url or item_id}")
        return str(download_url)

    async def _stream_to_bytes(self, url: str) -> bytes:
        # No Bearer header here (pre-authenticated Graph downloadUrl), so this
        # bypasses _request() and keeps its own 429/5xx retry loop in sync
        # with _compute_backoff's behavior instead.
        attempt = 0
        while True:
            async with self._semaphore:
                logger.debug("Streaming %s", url)
                resp = await self._client.get(url)

            status = resp.status_code
            if status < 400:
                return await resp.aread()

            if status == 429 or status >= 500:
                if attempt >= _DEFAULT_MAX_RETRIES:
                    raise_graph_error(resp, retries_exhausted=True)

                backoff = _compute_backoff(attempt, resp.headers.get("Retry-After"))
                logger.info(
                    "Retrying download in %.1fs (attempt %s)", backoff, attempt + 1
                )
                await asyncio.sleep(backoff)
                attempt += 1
                continue

            raise_graph_error(resp)

    async def _stream_to_path(self, url: str, dest: Path) -> Path:
        # Same no-Bearer/retry contract as _stream_to_bytes, but writes the
        # body to disk in chunks instead of buffering it in memory. Retries
        # only whole requests (status known before the body is consumed);
        # mid-stream resume is out of scope.
        attempt = 0
        while True:
            async with self._semaphore:
                logger.debug("Streaming %s to %s", url, dest)
                async with self._client.stream("GET", url) as resp:
                    status = resp.status_code
                    if status < 400:
                        with open(dest, "wb") as f:
                            async for chunk in resp.aiter_bytes():
                                f.write(chunk)
                        return dest
                    await resp.aread()

            if status == 429 or status >= 500:
                if attempt >= _DEFAULT_MAX_RETRIES:
                    raise_graph_error(resp, retries_exhausted=True)

                backoff = _compute_backoff(attempt, resp.headers.get("Retry-After"))
                logger.info(
                    "Retrying download in %.1fs (attempt %s)", backoff, attempt + 1
                )
                await asyncio.sleep(backoff)
                attempt += 1
                continue

            raise_graph_error(resp)

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

        data = await self.download(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        data, sheet = prepare_excel_bytes(data, sheet)
        buf = io.BytesIO(data)
        read_excel_kwargs.setdefault("engine", _DEFAULT_EXCEL_ENGINE)
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

        data = await self.download(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        buf = io.BytesIO(data)
        return pd.read_csv(buf, **read_csv_kwargs)

    async def read_parquet(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        **kwargs: Any,
    ) -> "pd.DataFrame":
        """Read a Parquet file directly into a DataFrame."""
        import io
        import pandas as pd

        data = await self.download(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        return pd.read_parquet(io.BytesIO(data), **kwargs)

    @overload
    async def download(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        to_path: None = None,
    ) -> bytes: ...

    @overload
    async def download(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        to_path: str | os.PathLike[str],
    ) -> Path: ...

    async def download(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        to_path: str | os.PathLike[str] | None = None,
    ) -> bytes | Path:
        """Return raw file bytes, or stream to disk when `to_path` is given.

        With `to_path` set, the body is written in chunks (never fully
        buffered in memory) and the written Path is returned.
        """
        url = await self._get_download_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        if to_path is not None:
            return await self._stream_to_path(url, Path(to_path))
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
        return_status: bool = False,
        **read_excel_kwargs: Any,
    ) -> "pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]":
        """Read many Excel files concurrently and concat into one DataFrame.

        If `return_status=True`, returns `(combined_df, status_df)`.
        """
        import io
        import pandas as pd
        from gex_msgraph._files import match_sheet_name

        engine = read_excel_kwargs.pop("engine", _DEFAULT_EXCEL_ENGINE)

        async def _process_one(path: str) -> tuple[pd.DataFrame | None, str, str]:
            try:
                data = await self.download(item_path=path)
                data, requested = prepare_excel_bytes(
                    data,
                    sheet,
                    sanitize_sheet_name=(sheet_match != "glob"),
                )
                buf = io.BytesIO(data)
                xls = pd.ExcelFile(buf, engine=engine)
                matched = match_sheet_name(xls.sheet_names, requested, sheet_match)
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to read %s: %s", path, e)
                return None, "error", str(e)

            if matched is None:
                if on_missing_sheet == "raise":
                    raise ValueError(f"Sheet '{sheet}' not found in {path}")
                if on_missing_sheet == "warn":
                    logger.warning("Sheet '%s' not found in %s", sheet, path)
                return None, "missing_sheet", f"Sheet '{sheet}' not found"

            try:
                df = pd.read_excel(xls, sheet_name=matched, **read_excel_kwargs)
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to read %s: %s", path, e)
                return None, "error", str(e)

            if add_source_column:
                df["_source"] = path
            return df, "success", ""

        results = await _gather_with_status(
            paths, _process_one, max_concurrent=max_concurrent
        )

        valid_dfs = []
        status_records = []
        for path, df, status, msg in results:
            if df is not None:
                valid_dfs.append(df)
            status_records.append({"path": path, "status": status, "error": msg})

        combined_df = (
            pd.concat(valid_dfs, ignore_index=True, sort=False)
            if valid_dfs
            else pd.DataFrame()
        )

        if return_status:
            status_df = pd.DataFrame(status_records)
            return combined_df, status_df
        return combined_df

    async def read_csv_many(
        self,
        paths: list[str],
        *,
        on_error: Literal["raise", "skip", "warn"] = "raise",
        add_source_column: bool = True,
        max_concurrent: int | None = None,
        return_status: bool = False,
        **read_csv_kwargs: Any,
    ) -> "pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]":
        """Read many CSV files concurrently and concat into one DataFrame.

        If `return_status=True`, returns `(combined_df, status_df)`.
        """
        import io
        import pandas as pd

        async def _process_one(path: str) -> tuple[pd.DataFrame | None, str, str]:
            try:
                data = await self.download(item_path=path)
                buf = io.BytesIO(data)
                df = pd.read_csv(buf, **read_csv_kwargs)
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to read %s: %s", path, e)
                return None, "error", str(e)

            if add_source_column:
                df["_source"] = path
            return df, "success", ""

        results = await _gather_with_status(
            paths, _process_one, max_concurrent=max_concurrent
        )

        valid_dfs = []
        status_records = []
        for path, df, status, msg in results:
            if df is not None:
                valid_dfs.append(df)
            status_records.append({"path": path, "status": status, "error": msg})

        combined_df = (
            pd.concat(valid_dfs, ignore_index=True, sort=False)
            if valid_dfs
            else pd.DataFrame()
        )

        if return_status:
            status_df = pd.DataFrame(status_records)
            return combined_df, status_df
        return combined_df

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
            base_url = f"{drive_root}/root:/{encode_drive_path(clean_path)}:/children"
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
            url = f"{drive_root}/root:/{encode_drive_path(clean_path)}:/children"
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

        _, url = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
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
        _, url = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        await self._request("DELETE", url)

    def _dest_folder_ref(self, dest_folder_path: str) -> dict[str, str]:
        """Build the parentReference payload for a destination folder path."""
        drive_root = self._drive_root()
        dest_clean = dest_folder_path.lstrip("/")
        return {
            "path": (
                f"{drive_root}/root:/{encode_drive_path(dest_clean)}"
                if dest_clean
                else f"{drive_root}/root"
            )
        }

    async def move_file(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        dest_folder_path: str | None = None,
        new_name: str | None = None,
    ) -> FileItem:
        """Move or rename a file. Source accepts any identifier kind;
        the destination is always a folder path in the current drive."""
        from gex_msgraph._files import parse_drive_item

        _, url = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )

        payload: dict[str, Any] = {}
        if dest_folder_path is not None:
            payload["parentReference"] = self._dest_folder_ref(dest_folder_path)
        if new_name is not None:
            payload["name"] = new_name

        if not payload:
            raise ValueError("Must provide dest_folder_path or new_name")

        resp = await self._request("PATCH", url, json=payload)
        return parse_drive_item(resp.json())

    async def create_folder(self, folder_path: str) -> FileItem:
        """Create a new folder."""
        from gex_msgraph._files import parse_drive_item

        drive_root = self._drive_root()
        clean_path = folder_path.lstrip("/")

        if "/" in clean_path:
            parent, name = clean_path.rsplit("/", 1)
            encoded_parent = encode_drive_path(parent)
            url = f"{drive_root}/root:/{encoded_parent}:/children"
        else:
            name = clean_path
            url = f"{drive_root}/root/children"

        payload = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        }

        resp = await self._request("POST", url, json=payload)
        return parse_drive_item(resp.json())

    async def copy_file(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        dest_folder_path: str,
        new_name: str | None = None,
        wait: bool = False,
        wait_timeout: float = 60.0,
    ) -> FileItem | None:
        """Copy a file to a new location. Graph executes the copy asynchronously.

        With `wait=False` (default) returns None immediately. With `wait=True`,
        polls Graph's copy-job monitor URL until completion (bounded by
        `wait_timeout` seconds) and returns the new item's FileItem.
        """
        from gex_msgraph._exceptions import GraphError

        kind, base = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        suffix = ":/copy" if kind == "path" else "/copy"

        payload: dict[str, Any] = {
            "parentReference": self._dest_folder_ref(dest_folder_path)
        }
        if new_name is not None:
            payload["name"] = new_name

        resp = await self._request("POST", base + suffix, json=payload)
        if not wait:
            return None

        monitor_url = resp.headers.get("Location")
        if not monitor_url:
            raise RuntimeError("Graph did not return a monitor URL for the copy job")

        # The monitor URL is pre-authenticated (like downloadUrl): poll it
        # directly, without a Bearer header. Transient 429/5xx on the monitor
        # count against wait_timeout rather than aborting the wait; 4xx aborts.
        poll_interval = 1.0
        for _ in range(max(1, int(wait_timeout / poll_interval))):
            poll = await self._client.get(monitor_url)
            status_code = poll.status_code
            if status_code >= 400:
                if status_code == 429 or status_code >= 500:
                    logger.info(
                        "Copy monitor returned %s; retrying poll", status_code
                    )
                    await asyncio.sleep(poll_interval)
                    continue
                raise_graph_error(poll)
            data = poll.json()
            job_status = data.get("status")
            if job_status == "completed":
                return await self.get_metadata(item_id=str(data["resourceId"]))
            if job_status == "failed":
                reason = data.get("error", {}).get("message", "unknown error")
                raise GraphError(
                    f"Copy job failed: {reason}",
                    request=poll.request,
                    response=poll,
                )
            await asyncio.sleep(poll_interval)

        raise TimeoutError(f"Copy job did not complete within {wait_timeout}s")

    async def exists(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> bool:
        """Return True if the item exists, False on 404. Re-raises other HTTP errors."""
        try:
            await self.get_metadata(
                item_path=item_path, share_url=share_url, item_id=item_id
            )
            return True
        except NotFoundError:
            return False

    async def list_excel_sheets(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
    ) -> list[str]:
        """List all worksheet names in an Excel file."""
        kind, base = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        # Path-addressed items use the ":/workbook/..." colon syntax;
        # id- and share-addressed items concatenate directly.
        suffix = ":/workbook/worksheets" if kind == "path" else "/workbook/worksheets"
        resp = await self._request("GET", base + suffix)
        return [str(sheet["name"]) for sheet in resp.json().get("value", [])]

    async def get_share_link(
        self,
        *,
        item_path: str | None = None,
        share_url: str | None = None,
        item_id: str | None = None,
        link_type: Literal["view", "edit"] = "view",
        scope: Literal["anonymous", "organization"] = "organization",
    ) -> str:
        """Create a sharing link for an item. Returns the webUrl string."""
        kind, base = self._resolve_item_url(
            item_path=item_path, share_url=share_url, item_id=item_id
        )
        suffix = ":/createLink" if kind == "path" else "/createLink"
        payload: dict[str, Any] = {"type": link_type, "scope": scope}
        resp = await self._request("POST", base + suffix, json=payload)
        return str(resp.json()["link"]["webUrl"])

    async def get_folder_tree(self, folder_path: str = "") -> "TreeNode":
        """Return a recursive tree representation of a folder."""
        from gex_msgraph._files import TreeNode

        # Drive root has no addressable metadata in the path-resolution form,
        # so leave `item` as None when listing from the root.
        root_item = (
            await self.get_metadata(item_path=folder_path) if folder_path else None
        )
        root_node = TreeNode(item=root_item, children=[])

        async def _build_tree(path: str, node: TreeNode) -> None:
            children_items = await self.list_files(path)
            subfolders: list[tuple[str, TreeNode]] = []
            for child in children_items:
                child_node = TreeNode(item=child, children=[])
                node.children.append(child_node)
                if child.is_folder:
                    subfolders.append((child.path, child_node))

            tasks = [_build_tree(sub_path, sub_node) for sub_path, sub_node in subfolders]
            if tasks:
                await asyncio.gather(*tasks)

        await _build_tree(folder_path, root_node)
        return root_node

    async def search_files(
        self,
        query: str,
        *,
        limit: int = 25,
    ) -> list[FileItem]:
        """Search for files across the drive by name or content."""
        import urllib.parse
        from gex_msgraph._files import parse_drive_item

        if limit <= 0:
            return []

        # OData v4 escapes a literal "'" inside a string literal by doubling
        # it; percent-encoding alone only protects the value in transport,
        # not the OData expression Graph parses after decoding the path
        # segment — so an un-doubled "'" would close the literal early.
        odata_query = query.replace("'", "''")
        encoded_query = urllib.parse.quote(odata_query)
        url = f"{self._drive_root()}/root/search(q='{encoded_query}')?$top={limit}"
        results: list[FileItem] = []
        async for item in self._iter_paginated(url):
            results.append(parse_drive_item(item))
            if len(results) >= limit:
                break
        return results[:limit]

    async def upload(
        self, local_path: str | os.PathLike[str], remote_path: str
    ) -> dict[str, Any]:
        """Upload a local file to remote_path. Returns the Graph driveItem dict.

        Files over ~4 MiB are transparently uploaded via a Graph upload
        session in chunks (simple PUT :/content is capped by Graph).
        """
        local_p = Path(local_path)
        if not local_p.is_file():
            raise FileNotFoundError(f"File not found: {local_path}")

        drive_root = self._drive_root()
        clean_remote = remote_path.lstrip("/")
        encoded_path = encode_drive_path(clean_remote)

        if local_p.stat().st_size > _UPLOAD_SESSION_THRESHOLD:
            return await self._upload_via_session(local_p, encoded_path)

        url = f"{drive_root}/root:/{encoded_path}:/content"

        with open(local_p, "rb") as f:
            data = f.read()

        resp = await self._request(
            "PUT",
            url,
            content=data,
            headers={"Content-Type": "application/octet-stream"},
        )
        return cast(dict[str, Any], resp.json())

    async def _upload_via_session(
        self, local_p: Path, encoded_remote: str
    ) -> dict[str, Any]:
        """Chunked upload for files over the simple-PUT size cap."""
        url = f"{self._drive_root()}/root:/{encoded_remote}:/createUploadSession"
        payload = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
        resp = await self._request("POST", url, json=payload)
        upload_url = str(resp.json()["uploadUrl"])

        total = local_p.stat().st_size
        result: dict[str, Any] | None = None
        with open(local_p, "rb") as f:
            start = 0
            while start < total:
                chunk = f.read(_UPLOAD_CHUNK_SIZE)
                end = start + len(chunk) - 1
                headers = {
                    "Content-Range": f"bytes {start}-{end}/{total}",
                    "Content-Length": str(len(chunk)),
                }
                chunk_resp = await self._put_chunk_with_retry(
                    upload_url, chunk, headers
                )
                if chunk_resp.status_code in (200, 201):
                    result = cast(dict[str, Any], chunk_resp.json())
                start = end + 1

        if result is None:
            raise RuntimeError(
                f"Upload session for {local_p.name} ended without a driveItem"
            )
        return result

    async def _put_chunk_with_retry(
        self, url: str, chunk: bytes, headers: dict[str, str]
    ) -> httpx.Response:
        # Upload-session URLs are pre-authenticated (no Bearer header), so
        # like _stream_to_bytes this can't go through _request(); it keeps
        # the same 429/5xx retry contract.
        attempt = 0
        while True:
            async with self._semaphore:
                resp = await self._client.put(url, content=chunk, headers=headers)

            status = resp.status_code
            if status < 400:
                return resp

            if status == 429 or status >= 500:
                if attempt >= _DEFAULT_MAX_RETRIES:
                    raise_graph_error(resp, retries_exhausted=True)

                backoff = _compute_backoff(attempt, resp.headers.get("Retry-After"))
                logger.info(
                    "Retrying upload chunk in %.1fs (attempt %s)", backoff, attempt + 1
                )
                await asyncio.sleep(backoff)
                attempt += 1
                continue

            raise_graph_error(resp)

    async def upload_many(
        self,
        items: list[tuple[str | os.PathLike[str], str]],
        *,
        on_error: Literal["raise", "skip", "warn"] = "raise",
        max_concurrent: int | None = None,
        return_status: bool = False,
    ) -> "list[dict[str, Any]] | tuple[list[dict[str, Any]], pd.DataFrame]":
        """Upload multiple local files concurrently."""
        import pandas as pd

        async def _upload_one(
            pair: tuple[str | os.PathLike[str], str]
        ) -> tuple[dict[str, Any] | None, str, str]:
            local, remote = pair
            try:
                result = await self.upload(local, remote)
                return result, "success", ""
            except Exception as e:
                if on_error == "raise":
                    raise
                if on_error == "warn":
                    logger.warning("Failed to upload %s: %s", remote, e)
                return None, "error", str(e)

        raw_results = await _gather_with_status(
            items, _upload_one, max_concurrent=max_concurrent
        )

        results: list[dict[str, Any]] = []
        status_records = []
        for (_, remote), result, status, msg in raw_results:
            if result is not None:
                results.append(result)
            status_records.append({"path": str(remote), "status": status, "error": msg})

        if return_status:
            return results, pd.DataFrame(status_records)
        return results

    async def send_mail(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        *,
        body_type: Literal["text", "html"] = "text",
        cc: str | list[str] | None = None,
        attachments: list[str | os.PathLike[str] | tuple[str, bytes] | dict[str, str]] | None = None,
    ) -> None:
        """Send an email from the account's mailbox."""

        def _make_recipients(addrs: str | list[str]) -> list[dict[str, Any]]:
            if isinstance(addrs, str):
                addrs = [addrs]
            return [{"emailAddress": {"address": a}} for a in addrs]

        async def _make_attachment(item: str | os.PathLike[str] | tuple[str, bytes] | dict[str, str]) -> dict[str, Any]:
            if isinstance(item, tuple):
                name, data = item
            elif isinstance(item, dict):
                data = cast(bytes, await self.download(**item))
                if "item_path" in item:
                    name = item["item_path"].split("/")[-1]
                else:
                    meta = await self.get_metadata(**item)
                    name = meta.name
            else:
                p = Path(item)
                name = p.name
                data = p.read_bytes()
            mime, _ = mimetypes.guess_type(name)
            return {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": mime or "application/octet-stream",
                "contentBytes": base64.b64encode(data).decode(),
            }

        payload: dict[str, Any] = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML" if body_type == "html" else "Text",
                    "content": body,
                },
                "toRecipients": _make_recipients(to),
            },
            "saveToSentItems": "true",
        }

        if cc:
            payload["message"]["ccRecipients"] = _make_recipients(cc)

        if attachments:
            payload["message"]["attachments"] = [
                await _make_attachment(a) for a in attachments
            ]

        await self._request("POST", "/me/sendMail", json=payload)

    async def list_mail(
        self,
        limit: int = 20,
        *,
        folder: str = "inbox",
    ) -> list[dict[str, Any]]:
        """List messages from a mail folder (default: inbox)."""
        select = "id,subject,from,receivedDateTime,bodyPreview,hasAttachments"
        url = f"/me/mailFolders/{folder}/messages?$top={limit}&$select={select}"
        resp = await self._request("GET", url)
        return cast(list[dict[str, Any]], resp.json().get("value", []))

    async def send_teams_message(
        self,
        team_id: str,
        channel_id: str,
        text: str,
    ) -> None:
        """Post a plain-text message to a Teams channel."""
        payload = {"body": {"contentType": "Text", "content": text}}
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

    async def list_chats(self, limit: int = 50) -> list[dict[str, Any]]:
        """List all chats (1-1, group, meeting) the account participates in, including members."""
        url = f"/me/chats?$expand=members&$top={limit}"
        resp = await self._request("GET", url)
        return cast(list[dict[str, Any]], resp.json().get("value", []))

    async def get_chat_messages(
        self,
        chat_id: str,
        limit: int = 10,
    ) -> list[dict[str, Any]]:
        """Fetch recent messages from a Teams chat."""
        url = f"/chats/{chat_id}/messages?$top={limit}"
        resp = await self._request("GET", url)
        return cast(list[dict[str, Any]], resp.json().get("value", []))

    async def send_chat_message(self, chat_id: str, text: str) -> None:
        """Post a plain-text message to a Teams chat."""
        payload = {"body": {"contentType": "Text", "content": text}}
        await self._request("POST", f"/chats/{chat_id}/messages", json=payload)

    def read_excel_sync(self, **kw: Any) -> "pd.DataFrame":
        return self._run_sync(self.read_excel(**kw))

    def read_csv_sync(self, **kw: Any) -> "pd.DataFrame":
        return self._run_sync(self.read_csv(**kw))

    def download_sync(self, **kw: Any) -> "bytes | Path":
        return cast("bytes | Path", self._run_sync(self.download(**kw)))

    def read_excel_many_sync(
        self, paths: list[str], **kw: Any
    ) -> "pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]":
        return self._run_sync(self.read_excel_many(paths, **kw))

    def read_csv_many_sync(
        self, paths: list[str], **kw: Any
    ) -> "pd.DataFrame | tuple[pd.DataFrame, pd.DataFrame]":
        return self._run_sync(self.read_csv_many(paths, **kw))
