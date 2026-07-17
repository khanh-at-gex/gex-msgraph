import asyncio

import httpx
import pytest
import respx

from gex_msgraph import (
    AuthError,
    GraphClient,
    GraphError,
    NotFoundError,
    PermissionDeniedError,
    RateLimitExhaustedError,
)


def test_graph_error_is_httpx_http_status_error_subclass():
    # Backward-compat contract: `except httpx.HTTPStatusError` keeps working.
    assert issubclass(GraphError, httpx.HTTPStatusError)
    for cls in (AuthError, PermissionDeniedError, NotFoundError, RateLimitExhaustedError):
        assert issubclass(cls, GraphError)


@pytest.mark.parametrize(
    "status,exc_cls",
    [(401, AuthError), (403, PermissionDeniedError), (404, NotFoundError), (400, GraphError)],
)
async def test_request_4xx_raises_mapped_error(env_vars, mock_token, status, exc_cls):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt").mock(
            return_value=httpx.Response(status)
        )
        async with GraphClient("das_u1") as client:
            with pytest.raises(exc_cls) as exc_info:
                await client.get_metadata(item_path="file.txt")

    # request/response stay attached, like httpx.HTTPStatusError.
    assert exc_info.value.response.status_code == status
    assert exc_info.value.request.method == "GET"


async def test_request_exhausted_retries_raises_rate_limit_exhausted(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        route = rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt")
        route.mock(return_value=httpx.Response(503))

        async with GraphClient("das_u1") as client:
            with pytest.MonkeyPatch.context() as m:
                async def mock_sleep(x):
                    pass
                m.setattr(asyncio, "sleep", mock_sleep)
                with pytest.raises(RateLimitExhaustedError):
                    await client.get_metadata(item_path="file.txt")
        assert route.call_count == 4  # initial + 3 retries


async def test_download_url_fetch_raises_mapped_error(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt").mock(
            return_value=httpx.Response(
                200, json={"@microsoft.graph.downloadUrl": "https://download.url"}
            )
        )
        rm.get("https://download.url").mock(return_value=httpx.Response(404))

        async with GraphClient("das_u1") as client:
            with pytest.raises(NotFoundError):
                await client.download(item_path="file.txt")


async def test_exists_false_on_404_true_on_200(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/missing.txt").mock(
            return_value=httpx.Response(404)
        )
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/present.txt").mock(
            return_value=httpx.Response(200, json={"name": "present.txt"})
        )
        async with GraphClient("das_u1") as client:
            assert await client.exists(item_path="missing.txt") is False
            assert await client.exists(item_path="present.txt") is True


async def test_exists_reraises_non_404(env_vars, mock_token):
    with respx.mock(assert_all_called=False) as rm:
        rm.get("https://graph.microsoft.com/v1.0/me/drive/root:/secret.txt").mock(
            return_value=httpx.Response(403)
        )
        async with GraphClient("das_u1") as client:
            with pytest.raises(PermissionDeniedError):
                await client.exists(item_path="secret.txt")


def test_token_failure_raises_graph_authentication_error(env_vars, monkeypatch):
    from gex_msgraph import GraphAuthenticationError
    from gex_msgraph._core import _TokenProvider

    class FakeApp:
        def get_accounts(self, username=None):
            return []

        def acquire_token_by_username_password(self, **kw):
            return {"error": "invalid_grant", "error_description": "bad creds"}

    monkeypatch.setattr(
        "gex_msgraph._core.msal.ConfidentialClientApplication",
        lambda *a, **kw: FakeApp(),
    )
    provider = _TokenProvider("cid", "sec", "tid", "user", "pass")
    with pytest.raises(GraphAuthenticationError, match="bad creds"):
        provider.get_token()
