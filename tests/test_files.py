import pytest
from datetime import datetime, timezone
from gex_msgraph._files import (
    validate_identifier,
    encode_share_url,
    encode_drive_path,
    build_resolution_url,
    match_sheet_name,
    parse_drive_item,
)

def test_validate_identifier():
    assert validate_identifier(item_path="a") == "path"
    assert validate_identifier(share_url="b") == "share"
    assert validate_identifier(item_id="c") == "id"
    
    with pytest.raises(ValueError):
        validate_identifier()
        
    with pytest.raises(ValueError):
        validate_identifier(item_path="a", share_url="b")

def test_encode_share_url():
    url = "https://tenant.sharepoint.com/sites/site/Shared%20Documents/test.xlsx"
    res = encode_share_url(url)
    assert res.startswith("u!")

def test_build_resolution_url():
    # Default OneDrive: drive_root="/me/drive"
    assert build_resolution_url("path", "file.xlsx", "/me/drive") == "/me/drive/root:/file.xlsx"
    # Explicit drive id: drive_root="/drives/{id}"
    assert build_resolution_url("id", "xyz", "/drives/drive1") == "/drives/drive1/items/xyz"
    # Share URLs ignore drive_root entirely.
    url = "https://tenant.sharepoint.com/a"
    encoded = encode_share_url(url)
    assert build_resolution_url("share", url, "/me/drive") == f"/shares/{encoded}/driveItem"

def test_build_resolution_url_encodes_special_characters():
    # "?" and "#" are URL metacharacters — unencoded, they truncate the path
    # (query string / fragment) instead of addressing the file.
    assert build_resolution_url("path", "Reports/a?b.xlsx", "/me/drive") == "/me/drive/root:/Reports/a%3Fb.xlsx"
    assert build_resolution_url("path", "a#b.xlsx", "/me/drive") == "/me/drive/root:/a%23b.xlsx"
    assert build_resolution_url("path", "50% done.xlsx", "/me/drive") == "/me/drive/root:/50%25%20done.xlsx"
    assert build_resolution_url("path", "Báo cáo Q1.xlsx", "/me/drive") == "/me/drive/root:/B%C3%A1o%20c%C3%A1o%20Q1.xlsx"

def test_encode_drive_path():
    assert encode_drive_path("file.xlsx") == "file.xlsx"
    assert encode_drive_path("a/b c") == "a/b%20c"  # "/" preserved as separator
    assert encode_drive_path("Reports/a?b.xlsx") == "Reports/a%3Fb.xlsx"
    assert encode_drive_path("a#b.xlsx") == "a%23b.xlsx"
    assert encode_drive_path("50% done.xlsx") == "50%25%20done.xlsx"

def test_match_sheet_name():
    sheets = ["Sheet1", "data", "REPORT_2026"]
    
    # Exact
    assert match_sheet_name(sheets, "data", "exact") == "data"
    assert match_sheet_name(sheets, "Data", "exact") is None
    
    # CI
    assert match_sheet_name(sheets, "DaTa", "ci") == "data"
    
    # Glob
    assert match_sheet_name(sheets, "REPORT*", "glob") == "REPORT_2026"
    assert match_sheet_name(sheets, "*1", "glob") == "Sheet1"
    
    # Int positional
    assert match_sheet_name(sheets, 1, "exact") == 1
    assert match_sheet_name(sheets, -1, "glob") == -1
    assert match_sheet_name(sheets, 5, "exact") is None

def test_parse_drive_item():
    item = {
        "id": "123",
        "name": "Q1.xlsx",
        "size": 1024,
        "lastModifiedDateTime": "2026-05-01T12:00:00Z",
        "webUrl": "https://tenant.sharepoint.com/sites/site/_layouts/15/Doc.aspx?sourcedoc=%7B123%7D",
        "parentReference": {
            "path": "/drive/root:/Reports/2026-Q1"
        }
    }
    fi = parse_drive_item(item)
    assert fi.name == "Q1.xlsx"
    assert fi.id == "123"
    assert fi.size == 1024
    assert not fi.is_folder
    assert fi.path == "Reports/2026-Q1/Q1.xlsx"
    assert fi.modified == datetime(2026, 5, 1, 12, 0, tzinfo=timezone.utc)
    assert fi.webUrl == "https://tenant.sharepoint.com/sites/site/_layouts/15/Doc.aspx?sourcedoc=%7B123%7D"

    folder = {
        "id": "456",
        "name": "EmptyFolder",
        "folder": {},
    }
    fi2 = parse_drive_item(folder, parent_path="Reports")
    assert fi2.is_folder
    assert fi2.path == "Reports/EmptyFolder"
    assert fi2.webUrl is None
