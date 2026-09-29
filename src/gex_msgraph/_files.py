from __future__ import annotations

import base64
import fnmatch
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Literal
from urllib.parse import parse_qs, quote, unquote, urlsplit


@dataclass(frozen=True)
class FileItem:
    name: str
    path: str
    id: str
    size: int
    modified: datetime
    is_folder: bool
    webUrl: str | None = None


@dataclass(frozen=True)
class DriveLocation:
    """A SharePoint/OneDrive folder resolved from a browser link.

    ``folder_path`` is relative to the drive root, without leading/trailing
    "/" ("" = the drive root). It is ``None`` only for a sharing link whose
    target Graph returns without a parent path (e.g. the caller can reach
    the item but not its drive by path) — use ``item_id`` then.
    """

    drive_id: str
    folder_path: str | None
    item_id: str
    drive_name: str
    drive_web_url: str
    site_id: str | None


@dataclass(frozen=True)
class ParsedFolderUrl:
    """Output of `parse_folder_url`. For ``kind="share"`` only ``url`` is
    meaningful; for ``kind="path"`` ``host`` + ``segments`` hold the decoded
    (NFC) server-relative folder path."""

    kind: Literal["path", "share"]
    url: str
    host: str = ""
    segments: tuple[str, ...] = ()


@dataclass
class TreeNode:
    item: FileItem | None
    children: list["TreeNode"]

    def print(self, indent: int = 0) -> None:
        """Print the tree structure visually."""
        prefix = "  " * indent
        if self.item:
            icon = "📁" if self.item.is_folder else "📄"
            print(f"{prefix}{icon} {self.item.name} ({self.item.size} bytes)")
        else:
            print(f"{prefix}📁 (Root)")

        for child in self.children:
            child.print(indent + 1)


def validate_identifier(
    item_path: str | None = None,
    share_url: str | None = None,
    item_id: str | None = None,
) -> Literal["path", "share", "id"]:
    """Ensure exactly one identifier is provided."""
    provided = sum(x is not None for x in (item_path, share_url, item_id))
    if provided != 1:
        raise ValueError("Must provide exactly one of: item_path, share_url, item_id")

    if item_path is not None:
        return "path"
    if share_url is not None:
        return "share"
    return "id"


def encode_share_url(url: str) -> str:
    """Encode a SharePoint share URL for Microsoft Graph."""
    # Graph API requirement: 'u!' + base64url(url) without padding
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return f"u!{b64.rstrip('=')}"


def encode_drive_path(path: str) -> str:
    """Percent-encode a drive-relative path segment for Graph's `root:/{path}`
    colon syntax — used both in request URLs and in `parentReference.path`
    JSON values, which Graph parses with the same syntax. Preserves "/" as
    the path separator."""
    return quote(path, safe="/")


def build_resolution_url(
    kind: Literal["path", "share", "id"],
    value: str,
    drive_root: str,
) -> str:
    """Build the correct Graph endpoint for the identifier kind.

    `drive_root` is the URL fragment for the drive scope, e.g.
    `/me/drive` (default OneDrive) or `/drives/{driveId}` (explicit).
    """
    if kind == "path":
        clean_path = value.lstrip("/")
        return f"{drive_root}/root:/{encode_drive_path(clean_path)}"
    elif kind == "share":
        encoded_url = encode_share_url(value)
        return f"/shares/{encoded_url}/driveItem"
    elif kind == "id":
        return f"{drive_root}/items/{value}"
    raise ValueError(f"Unknown kind: {kind}")


def match_sheet_name(
    sheets: list[str],
    requested: str | int,
    mode: Literal["exact", "ci", "glob"],
) -> str | int | None:
    """Find matching sheet name; return None if no match."""
    if isinstance(requested, int):
        if 0 <= requested < len(sheets) or -len(sheets) <= requested < 0:
            return requested
        return None

    if mode == "exact":
        if requested in sheets:
            return requested
    elif mode == "ci":
        req_lower = requested.lower()
        for sheet in sheets:
            if sheet.lower() == req_lower:
                return sheet
    elif mode == "glob":
        for sheet in sheets:
            if fnmatch.fnmatch(sheet, requested):
                return sheet
    else:
        raise ValueError(f"Invalid mode: {mode}")

    return None


def parse_drive_item(item: dict[str, Any], parent_path: str = "") -> FileItem:
    """Convert Graph API response dict into FileItem."""
    name = item.get("name", "")
    is_folder = "folder" in item

    path_val = item.get("parentReference", {}).get("path", "")
    # Graph returns the parent path as either "/drive/root:" (default OneDrive)
    # or "/drives/{driveId}/root:" (explicit drive). Strip whichever applies.
    rel_dir: str | None = None
    if path_val:
        m = re.match(r"^/drives/[^/]+/root:", path_val)
        if m:
            rel_dir = path_val[m.end():].lstrip("/")
        elif path_val.startswith("/drive/root:"):
            rel_dir = path_val[len("/drive/root:"):].lstrip("/")

    if rel_dir is not None:
        path = f"{rel_dir}/{name}" if rel_dir else name
    else:
        path = f"{parent_path}/{name}".lstrip("/") if parent_path else name

    mod_str = item.get("lastModifiedDateTime", "")
    if mod_str:
        if mod_str.endswith("Z"):
            mod_str = mod_str[:-1] + "+00:00"
        modified = datetime.fromisoformat(mod_str)
    else:
        modified = datetime.min

    return FileItem(
        name=name,
        path=path,
        id=item.get("id", ""),
        size=item.get("size", 0),
        modified=modified,
        is_folder=is_folder,
        webUrl=item.get("webUrl") or None,
    )


_SITE_ROOTS = ("sites", "teams", "personal")
_SHARING_LINK_RE = re.compile(r"^/:[a-z]:/", re.IGNORECASE)
_SUPPORTED_FOLDER_URLS = (
    "Supported SharePoint folder links: "
    "library view (…/Forms/AllItems.aspx?id=/sites/…), "
    "direct folder path (https://{host}/sites|teams/{site}/{library}/…), "
    "OneDrive (https://{tenant}-my.sharepoint.com/personal/{user}/Documents/…), "
    "sharing link (https://{host}/:f:/s|r|g/…)."
)


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


def parse_folder_url(url: str) -> ParsedFolderUrl:
    """Classify a browser-copied SharePoint folder link without any HTTP call.

    Never echoes the URL in errors — sharing links carry an access token.
    """
    parts = urlsplit(url.strip())
    host = (parts.hostname or "").lower()
    if parts.scheme not in ("http", "https") or not host.endswith(".sharepoint.com"):
        raise ValueError(f"Not a SharePoint URL. {_SUPPORTED_FOLDER_URLS}")

    if _SHARING_LINK_RE.match(parts.path):
        return ParsedFolderUrl(kind="share", url=url.strip())

    # Library views carry the folder as a server-relative path in `id`
    # (older views: `RootFolder`); every other query parameter is noise.
    query = {k.lower(): v for k, v in parse_qs(parts.query).items()}
    folder_param = query.get("id") or query.get("rootfolder")
    server_path = folder_param[0] if folder_param else unquote(parts.path)

    segments = [s for s in _nfc(server_path).split("/") if s]
    # A bare library view (no `id`) points at the library root.
    if len(segments) >= 2 and segments[-2].lower() == "forms" and segments[
        -1
    ].lower().endswith(".aspx"):
        segments = segments[:-2]

    if len(segments) < 3 or segments[0].lower() not in _SITE_ROOTS:
        raise ValueError(
            f"Could not find a site and library in the URL. {_SUPPORTED_FOLDER_URLS}"
        )
    return ParsedFolderUrl(
        kind="path",
        url=f"https://{host}/" + "/".join(segments),
        host=host,
        segments=tuple(segments),
    )


def candidate_site_paths(segments: tuple[str, ...], max_tries: int = 3) -> list[str]:
    """Server-relative site paths to try, shortest first: `/sites/x`, then
    subsites `/sites/x/y`, … Each candidate leaves at least one segment for
    the library."""
    last = min(2 + max_tries, len(segments))
    return ["/" + "/".join(segments[:n]) for n in range(2, last)]


def _url_segments(url: str) -> tuple[str, tuple[str, ...]]:
    parts = urlsplit(url)
    path = _nfc(unquote(parts.path))
    return (parts.hostname or "").lower(), tuple(
        s.casefold() for s in path.split("/") if s
    )


def match_drive(
    drives: list[dict[str, Any]], host: str, segments: tuple[str, ...]
) -> tuple[dict[str, Any], str] | None:
    """Pick the drive whose `webUrl` is the longest segment-wise prefix of
    the folder, returning it with the remaining folder path.

    Matches on webUrl, never on the drive name: the URL says "Shared
    Documents" while the name is "Documents" (or a localized name).
    Segment-wise comparison keeps ".../Documents" from matching
    ".../Documents Archive".
    """
    folder_cf = tuple(s.casefold() for s in segments)
    best: tuple[dict[str, Any], int] | None = None
    for drive in drives:
        drive_host, drive_segs = _url_segments(str(drive.get("webUrl") or ""))
        n = len(drive_segs)
        if drive_host != host or not n or folder_cf[:n] != drive_segs:
            continue
        if best is None or n > best[1]:
            best = (drive, n)
    if best is None:
        return None
    drive, n = best
    return drive, "/".join(segments[n:])


def shared_item_folder_path(item: dict[str, Any]) -> str | None:
    """Drive-relative path of a driveItem fetched via /shares, or None when
    Graph omits the parent path."""
    if "root" in item:
        return ""
    path_val = item.get("parentReference", {}).get("path")
    if not path_val or ":" not in path_val:
        return None
    rel_dir = _nfc(unquote(path_val.split(":", 1)[1])).strip("/")
    name = _nfc(item.get("name", ""))
    return f"{rel_dir}/{name}" if rel_dir else name
