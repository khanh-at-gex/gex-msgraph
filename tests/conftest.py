import pytest
import respx

@pytest.fixture
def env_vars(monkeypatch):
    monkeypatch.setenv("MS_DAS_U1_CLIENT_ID", "fake_client_id")
    monkeypatch.setenv("MS_DAS_U1_CLIENT_SECRET", "fake_client_secret")
    monkeypatch.setenv("MS_DAS_U1_TENANT_ID", "fake_tenant_id")
    monkeypatch.setenv("MS_DAS_U1_USERNAME", "fake_user")
    monkeypatch.setenv("MS_DAS_U1_PASSWORD", "fake_pass")

@pytest.fixture
def mock_token(monkeypatch):
    from gex_msgraph._core import _TokenProvider
    monkeypatch.setattr(_TokenProvider, "__init__", lambda self, *args, **kwargs: None)
    monkeypatch.setattr(_TokenProvider, "get_token", lambda self: "fake-token")

@pytest.fixture
def mock_graph():
    with respx.mock(base_url="https://graph.microsoft.com/v1.0") as respx_mock:
        yield respx_mock
