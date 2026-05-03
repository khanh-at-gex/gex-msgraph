from __future__ import annotations

import base64
import fnmatch
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Literal

@dataclass(frozen=True)
class FileItem:
    name: str
    path: str
    id: str
    size: int
    modified: datetime
    is_folder: bool

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

def build_resolution_url(
    kind: Literal["path", "share", "id"],
    value: str,
    drive_id: str,
) -> str:
    """Build the correct Graph endpoint for the identifier kind."""
    if kind == "path":
        # Ensure path doesn't start with / for the root:/{path} format
        clean_path = value.lstrip("/")
        return f"/drives/{drive_id}/root:/{clean_path}"
    elif kind == "share":
        encoded_url = encode_share_url(value)
        return f"/shares/{encoded_url}/driveItem"
    elif kind == "id":
        return f"/drives/{drive_id}/items/{value}"
    raise ValueError(f"Unknown kind: {kind}")

def match_sheet_name(
    sheets: list[str],
    requested: str | int,
    mode: Literal["exact", "ci", "glob"]
) -> str | int | None:
    """Find matching sheet name; return None if no match."""
    if isinstance(requested, int):
        # Positional index: just return it if it's within bounds
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
    if path_val:
        prefix = "/drive/root:"
        if path_val.startswith(prefix):
            rel_dir = path_val[len(prefix):].lstrip("/")
            path = f"{rel_dir}/{name}" if rel_dir else name
        else:
            path = f"{parent_path}/{name}".lstrip("/") if parent_path else name
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
    )
