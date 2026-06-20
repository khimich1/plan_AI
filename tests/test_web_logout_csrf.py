from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_login
from app.security.csrf import CSRF_COOKIE_NAME
from app.security.session import SESSION_COOKIE_NAME
from tests.helpers.csrf import CsrfAwareTestClient

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
_LOGIN_JSON = {"username": "admin", "password": "StrongPassword123!"}


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _mock_authenticate_success(monkeypatch: pytest.MonkeyPatch) -> None:
    patch_auth_login(monkeypatch)


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> CsrfAwareTestClient:
    _mock_authenticate_success(monkeypatch)
    return CsrfAwareTestClient(create_app())


def test_legacy_get_logout_returns_405_without_session_teardown(client: CsrfAwareTestClient) -> None:
    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200
    assert client.cookies.get(SESSION_COOKIE_NAME)

    response = client.get("/web/logout")

    assert response.status_code == 405
    assert response.headers.get("deprecation") == "true"
    assert "POST /api/v1/auth/logout" in response.json()["detail"]
    assert client.cookies.get(SESSION_COOKIE_NAME)

    me = client.get("/api/v1/auth/me")
    assert me.status_code == 200


def test_api_logout_without_csrf_is_rejected(client: CsrfAwareTestClient) -> None:
    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200

    raw_client = TestClient(create_app())
    raw_client.cookies.update(client.cookies)
    blocked = raw_client.post("/api/v1/auth/logout")

    assert blocked.status_code == 403
    assert blocked.json()["detail"] == "CSRF token missing or invalid."

    me = client.get("/api/v1/auth/me")
    assert me.status_code == 200


def test_api_logout_with_csrf_clears_session(client: CsrfAwareTestClient) -> None:
    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200
    assert client.cookies.get(CSRF_COOKIE_NAME)

    logout = client.post("/api/v1/auth/logout")
    assert logout.status_code == 200
    assert logout.json() == {"ok": True}

    raw_cookie = logout.headers.get("set-cookie", "")
    assert SESSION_COOKIE_NAME in raw_cookie
    assert CSRF_COOKIE_NAME in raw_cookie

    me = client.get("/api/v1/auth/me")
    assert me.status_code == 401
