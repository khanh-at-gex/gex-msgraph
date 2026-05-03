import pytest
import httpx
import respx
import pandas as pd
import asyncio
from gex_msgraph import GraphClient

async def test_init_missing_env(env_vars, monkeypatch):
    monkeypatch.delenv("MS_DAS_U1_CLIENT_ID")
    with pytest.raises(KeyError):
        GraphClient("das_u1")

async def test_download(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/drives/me/root:/file.txt").mock(
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
        route = rm.get("https://graph.microsoft.com/v1.0/drives/me/root:/file.txt")
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
        route = rm.get("https://graph.microsoft.com/v1.0/drives/me/root:/file.txt")
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
        route = rm.get("https://graph.microsoft.com/v1.0/drives/me/root:/file.txt")
        route.mock(return_value=httpx.Response(403))
        
        client = GraphClient("das_u1")
        async with client:
            with pytest.raises(httpx.HTTPStatusError):
                await client.download(item_path="file.txt")
            
            assert route.call_count == 1

async def test_walk_pagination(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route1 = rm.get("https://graph.microsoft.com/v1.0/drives/me/root/children")
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
        route = rm.post("https://graph.microsoft.com/v1.0/users/me/sendMail")
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
        rm.get("https://graph.microsoft.com/v1.0/drives/me/root:/file.txt").mock(
            return_value=httpx.Response(200, json={"@microsoft.graph.downloadUrl": "https://download.url"})
        )
        rm.get("https://download.url").mock(
            return_value=httpx.Response(200, content=b"hello bytes")
        )
        
        client = GraphClient("das_u1")
        res = client.download_sync(item_path="file.txt")
        assert res == b"hello bytes"
