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


# --- parse_folder_url / match_drive -------------------------------------

from gex_msgraph._files import (  # noqa: E402
    candidate_site_paths,
    match_drive,
    parse_folder_url,
    shared_item_folder_path,
)

VN_FOLDER = "08. Ứng dụng - Công cụ KTNB"
VN_ENCODED = "08.%20%E1%BB%A8ng%20d%E1%BB%A5ng%20-%20C%C3%B4ng%20c%E1%BB%A5%20KTNB"


def test_parse_view_link_id_param():
    url = (
        "https://contoso.sharepoint.com/sites/Audit/Shared%20Documents/Forms/"
        "AllItems.aspx?id=%2Fsites%2FAudit%2FShared%20Documents%2F"
        f"{VN_ENCODED}%2FSub&viewid=abc-123&FolderCTID=0x0120"
    )
    parsed = parse_folder_url(url)
    assert parsed.kind == "path"
    assert parsed.host == "contoso.sharepoint.com"
    assert parsed.segments == ("sites", "Audit", "Shared Documents", VN_FOLDER, "Sub")


def test_parse_view_link_rootfolder_param():
    url = (
        "https://contoso.sharepoint.com/sites/Audit/Shared%20Documents/Forms/"
        "AllItems.aspx?RootFolder=%2Fsites%2FAudit%2FShared%20Documents%2FA&at=1"
    )
    assert parse_folder_url(url).segments == ("sites", "Audit", "Shared Documents", "A")


def test_parse_bare_library_view_points_at_library_root():
    url = "https://contoso.sharepoint.com/sites/Audit/Shared%20Documents/Forms/AllItems.aspx"
    assert parse_folder_url(url).segments == ("sites", "Audit", "Shared Documents")


@pytest.mark.parametrize("suffix", ["", "/"])
def test_parse_direct_folder_link(suffix):
    url = f"https://contoso.sharepoint.com/sites/Audit/Shared%20Documents/{VN_ENCODED}{suffix}"
    parsed = parse_folder_url(url)
    assert parsed.segments == ("sites", "Audit", "Shared Documents", VN_FOLDER)
    assert parsed.url == f"https://contoso.sharepoint.com/sites/Audit/Shared Documents/{VN_FOLDER}"


def test_parse_teams_and_personal_links():
    teams = parse_folder_url("https://contoso.sharepoint.com/teams/Risk/Documents/A%20B")
    assert teams.segments == ("teams", "Risk", "Documents", "A B")
    personal = parse_folder_url(
        "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Reports"
    )
    assert personal.host == "contoso-my.sharepoint.com"
    assert personal.segments[:2] == ("personal", "user_contoso_com")


def test_parse_normalizes_to_nfc():
    import unicodedata

    nfd = unicodedata.normalize("NFD", VN_FOLDER)
    parsed = parse_folder_url(f"https://contoso.sharepoint.com/sites/A/Docs/{nfd}")
    assert parsed.segments[-1] == VN_FOLDER


@pytest.mark.parametrize("kind", ["f", "r", "g"])
def test_parse_sharing_link(kind):
    url = f"https://contoso.sharepoint.com/:f:/{kind}/Audit/EaBcToken?e=xyz"
    parsed = parse_folder_url(url)
    assert parsed.kind == "share"
    assert parsed.url == url


@pytest.mark.parametrize(
    "url",
    [
        "not a url",
        "https://example.com/sites/A/Docs/x",
        "ftp://contoso.sharepoint.com/sites/A/Docs",
        "https://contoso.sharepoint.com/sites/A",
        "https://contoso.sharepoint.com/Shared%20Documents/x",
        "",
    ],
)
def test_parse_rejects_garbage(url):
    with pytest.raises(ValueError, match="Supported SharePoint folder links"):
        parse_folder_url(url)


def test_parse_error_never_echoes_url():
    secret = "https://evil.example.com/:f:/s/Audit/SECRETTOKEN"
    with pytest.raises(ValueError) as exc:
        parse_folder_url(secret)
    assert "SECRETTOKEN" not in str(exc.value)


def test_candidate_site_paths():
    assert candidate_site_paths(("sites", "A", "Docs")) == ["/sites/A"]
    assert candidate_site_paths(("sites", "A", "B", "C", "D", "E", "F")) == [
        "/sites/A",
        "/sites/A/B",
        "/sites/A/B/C",
    ]


def _drive(name, web_url):
    return {"id": f"id-{name}", "name": name, "webUrl": web_url}


def test_match_drive_longest_prefix_and_segment_boundary():
    base = "https://contoso.sharepoint.com/sites/A"
    docs = _drive("Documents", f"{base}/Documents")
    archive = _drive("Documents Archive", f"{base}/Documents%20Archive")
    segs = ("sites", "A", "Documents Archive", "2024")

    assert match_drive([docs, archive], "contoso.sharepoint.com", segs) == (archive, "2024")
    # Only "Documents" visible: a string prefix, not a segment prefix — no match.
    assert match_drive([docs], "contoso.sharepoint.com", segs) is None


def test_match_drive_uses_web_url_not_name():
    drive = _drive("Documents", "https://contoso.sharepoint.com/sites/A/Shared%20Documents/")
    segs = ("sites", "A", "shared documents", VN_FOLDER)
    assert match_drive([drive], "contoso.sharepoint.com", segs) == (drive, VN_FOLDER)


def test_match_drive_library_root_and_other_host():
    drive = _drive("Documents", "https://contoso.sharepoint.com/sites/A/Shared%20Documents")
    segs = ("sites", "A", "Shared Documents")
    assert match_drive([drive], "contoso.sharepoint.com", segs) == (drive, "")
    assert match_drive([drive], "other.sharepoint.com", segs) is None


def test_shared_item_folder_path():
    item = {"name": "Q1", "parentReference": {"path": "/drives/b!x/root:/A%20B/C"}}
    assert shared_item_folder_path(item) == "A B/C/Q1"
    top = {"name": "Q1", "parentReference": {"path": "/drives/b!x/root:"}}
    assert shared_item_folder_path(top) == "Q1"
    assert shared_item_folder_path({"name": "Q1", "parentReference": {}}) is None
    assert shared_item_folder_path({"name": "root", "root": {}}) == ""
