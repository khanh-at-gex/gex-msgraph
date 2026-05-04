import pytest
import httpx
import respx
import asyncio
from gex_msgraph import GraphClient

async def test_init_missing_env(env_vars, monkeypatch):
    monkeypatch.delenv("MS_DAS_U1_CLIENT_ID")
    with pytest.raises(KeyError):
        GraphClient("das_u1")

async def test_download(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt").mock(
            return_value=httpx.Response(200, json={"@microsoft.graph.downloadUrl": "https://download.url"})
        )
        rm.get("https://download.url").mock(
            return_value=httpx.Response(200, content=b"hello bytes")
        )
        
        client = GraphClient("das_u1")
        async with client:
            res = await client.download(item_path="file.txt")
            assert res == b"hello bytes"

async def test_retry_429(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.side_effect = [
            httpx.Response(429, headers={"Retry-After": "0"}),
            httpx.Response(200, json={"@microsoft.graph.downloadUrl": "https://download.url"})
        ]
        rm.get("https://download.url").mock(
            return_value=httpx.Response(200, content=b"ok")
        )
        
        client = GraphClient("das_u1")
        async with client:
            res = await client.download(item_path="file.txt")
            assert res == b"ok"
            assert route.call_count == 2

async def test_retry_503(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(503))
        
        client = GraphClient("das_u1")
        
        async with client:
            with pytest.raises(httpx.HTTPStatusError):
                # Will retry 3 times plus initial = 4 calls.
                # However, backoff can take a bit if it uses real sleep.
                # We can mock sleep.
                async def mock_sleep(x):
                    pass
                with pytest.MonkeyPatch.context() as m:
                    m.setattr(asyncio, "sleep", mock_sleep)
                    await client.download(item_path="file.txt")
            
            assert route.call_count == 4

async def test_no_retry_403(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(403))
        
        client = GraphClient("das_u1")
        async with client:
            with pytest.raises(httpx.HTTPStatusError):
                await client.download(item_path="file.txt")
            
            assert route.call_count == 1

async def test_walk_pagination(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route1 = rm.get("https://graph.microsoft.com/v1.0/me/drive/root/children")
        route1.mock(return_value=httpx.Response(200, json={
            "value": [{"name": "file1.txt", "size": 10}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/page2"
        }))
        
        route2 = rm.get("https://graph.microsoft.com/v1.0/page2")
        route2.mock(return_value=httpx.Response(200, json={
            "value": [{"name": "file2.txt", "size": 20}]
        }))
        
        client = GraphClient("das_u1")
        async with client:
            files = await client.walk()
            assert len(files) == 2
            assert files[0].name == "file1.txt"
            assert files[1].name == "file2.txt"

async def test_send_mail(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.post("https://graph.microsoft.com/v1.0/me/sendMail")
        route.mock(return_value=httpx.Response(202))
        
        client = GraphClient("das_u1")
        async with client:
            await client.send_mail(to="test@test.com", subject="Subj", body="Hello")
            
        req = route.calls[0].request
        payload = req.read().decode("utf-8")
        assert "test@test.com" in payload
        assert "Subj" in payload

async def test_send_teams_message(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.post("https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages")
        route.mock(return_value=httpx.Response(201))
        
        client = GraphClient("das_u1")
        async with client:
            await client.send_teams_message(team_id="t1", channel_id="c1", text="Hello Teams")
            
        req = route.calls[0].request
        payload = req.read().decode("utf-8")
        assert "Hello Teams" in payload

def test_sync_wrappers(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt").mock(
            return_value=httpx.Response(200, json={"@microsoft.graph.downloadUrl": "https://download.url"})
        )
        rm.get("https://download.url").mock(
            return_value=httpx.Response(200, content=b"hello bytes")
        )
        
        client = GraphClient("das_u1")
        res = client.download_sync(item_path="file.txt")
        assert res == b"hello bytes"

async def test_get_metadata(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(200, json={"name": "file.txt", "size": 123}))
        
        client = GraphClient("das_u1")
        async with client:
            meta = await client.get_metadata(item_path="file.txt")
            assert meta.name == "file.txt"
            assert meta.size == 123

async def test_delete_file(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.delete("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(204))
        
        client = GraphClient("das_u1")
        async with client:
            await client.delete_file(item_path="file.txt")
            assert route.called

async def test_move_file(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.patch("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(200, json={"name": "new.txt"}))
        
        client = GraphClient("das_u1")
        async with client:
            res = await client.move_file("file.txt", "dest/folder", new_name="new.txt")
            assert res.name == "new.txt"
            payload = route.calls[0].request.read().decode()
            assert "dest/folder" in payload
            assert "new.txt" in payload

async def test_create_folder(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.post("https://graph.microsoft.com/v1.0/me/drive/root:/parent:/children")
        route.mock(return_value=httpx.Response(201, json={"name": "child", "folder": {}}))
        
        client = GraphClient("das_u1")
        async with client:
            res = await client.create_folder("parent/child")
            assert res.name == "child"
            assert res.is_folder

async def test_list_excel_sheets(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.xlsx:/workbook/worksheets")
        route.mock(return_value=httpx.Response(200, json={"value": [{"name": "Sheet1"}, {"name": "Sheet2"}]}))
        
        client = GraphClient("das_u1")
        async with client:
            res = await client.list_excel_sheets(item_path="file.xlsx")
            assert res == ["Sheet1", "Sheet2"]

async def test_get_folder_tree(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/root").mock(
            return_value=httpx.Response(200, json={"name": "root", "folder": {}})
        )
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/root:/children").mock(
            return_value=httpx.Response(200, json={"value": [
                {"name": "file.txt"},
                {"name": "sub", "folder": {}}
            ]})
        )
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/root/sub:/children").mock(
            return_value=httpx.Response(200, json={"value": []})
        )
        
        client = GraphClient("das_u1")
        async with client:
            tree = await client.get_folder_tree("root")
            assert tree.item is not None and tree.item.name == "root"
            assert len(tree.children) == 2
            assert tree.children[0].item.name == "file.txt"
            assert tree.children[1].item.name == "sub"

async def test_get_teams_messages(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages?$top=10")
        route.mock(return_value=httpx.Response(200, json={"value": [{"id": "m1"}]}))

        client = GraphClient("das_u1")
        async with client:
            res = await client.get_teams_messages("t1", "c1")
            assert len(res) == 1
            assert res[0]["id"] == "m1"


async def test_explicit_drive_id_url_format(env_vars_with_drive, mock_token):
    """Bug 1 regression: explicit DEFAULT_DRIVE_ID emits /drives/{id}/..., not /me/drive/..."""
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/drives/b!abc/root:/file.txt")
        route.mock(return_value=httpx.Response(200, json={"name": "file.txt", "size": 1}))

        client = GraphClient("das_u1")
        async with client:
            meta = await client.get_metadata(item_path="file.txt")
            assert meta.name == "file.txt"
            assert route.called


async def test_read_excel_many_missing_sheet_raise_overrides_on_error(env_vars, mock_token):
    """Bug 4 regression: on_missing_sheet='raise' must raise even when on_error='skip'."""
    import io
    import openpyxl

    # Build a real one-sheet xlsx in memory so pd.ExcelFile can open it.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OnlySheet"
    ws["A1"] = "hello"
    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/f.xlsx").mock(
            return_value=httpx.Response(200, json={"@microsoft.graph.downloadUrl": "https://dl.url"})
        )
        rm.get("https://dl.url").mock(return_value=httpx.Response(200, content=xlsx_bytes))

        client = GraphClient("das_u1")
        async with client:
            with pytest.raises(ValueError, match="not found"):
                await client.read_excel_many(
                    ["f.xlsx"],
                    sheet="DoesNotExist",
                    on_missing_sheet="raise",
                    on_error="skip",
                )


async def test_list_excel_sheets_by_item_id(env_vars, mock_token):
    """Bug 3 regression: id-addressed list_excel_sheets uses /items/{id}/workbook/worksheets (no colon)."""
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/items/xyz/workbook/worksheets")
        route.mock(return_value=httpx.Response(200, json={"value": [{"name": "S1"}]}))

        client = GraphClient("das_u1")
        async with client:
            res = await client.list_excel_sheets(item_id="xyz")
            assert res == ["S1"]


async def test_list_excel_sheets_by_share_url(env_vars, mock_token):
    """Bug 3 regression: share-addressed list_excel_sheets uses /shares/{enc}/driveItem/workbook/worksheets."""
    from gex_msgraph._files import encode_share_url
    share = "https://tenant.sharepoint.com/sites/x/Shared%20Documents/f.xlsx"
    encoded = encode_share_url(share)

    with respx.mock(assert_all_called=False) as rm:
        route = rm.get(
            f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem/workbook/worksheets"
        )
        route.mock(return_value=httpx.Response(200, json={"value": [{"name": "S1"}]}))

        client = GraphClient("das_u1")
        async with client:
            res = await client.list_excel_sheets(share_url=share)
            assert res == ["S1"]
