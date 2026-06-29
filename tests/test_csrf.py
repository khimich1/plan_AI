from __future__ import annotations

import re
from unittest.mock import MagicMock

import pytest
from fastapi import Response
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_login
from app.security.csrf import (
    CSRF_COOKIE_NAME,
    CSRF_HEADER_NAME,
    clear_csrf_cookie,
    generate_csrf_token,
    set_csrf_cookie,
    tokens_match,
)
from app.security.session import SESSION_COOKIE_NAME
from tests.helpers.csrf import CsrfAwareTestClient, csrf_headers, ensure_csrf_cookie

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
_LOGIN_JSON = {"username": "admin", "password": "StrongPassword123!"}
_LOGIN_FORM = {"username": "admin", "password": "StrongPassword123!"}


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


def test_tokens_match_requires_equal_length() -> None:
    token = generate_csrf_token()
    assert tokens_match(token, token) is True
    assert tokens_match(token, token + "x") is False
    assert tokens_match(None, token) is False
    assert tokens_match(token, None) is False


def test_set_csrf_cookie_is_not_httponly(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("COOKIE_SAMESITE", "lax")
    monkeypatch.setenv("COOKIE_SECURE", "false")
    get_settings.cache_clear()
    response = MagicMock(spec=Response)

    set_csrf_cookie(response, "csrf-value")

    response.set_cookie.assert_called_once()
    args, kwargs = response.set_cookie.call_args
    assert args[0] == CSRF_COOKIE_NAME
    assert kwargs["httponly"] is False
    assert kwargs["samesite"] == "lax"
    assert kwargs["path"] == "/"


def test_clear_csrf_cookie_uses_shared_policy(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("COOKIE_SAMESITE", "strict")
    monkeypatch.setenv("COOKIE_SECURE", "true")
    get_settings.cache_clear()
    response = MagicMock(spec=Response)

    clear_csrf_cookie(response)

    response.delete_cookie.assert_called_once_with(
        CSRF_COOKIE_NAME,
        httponly=False,
        samesite="strict",
        secure=True,
        path="/",
    )


def test_get_health_sets_csrf_cookie(client: CsrfAwareTestClient) -> None:
    response = client.get("/health")

    assert response.status_code == 200
    assert client.cookies.get(CSRF_COOKIE_NAME)


def test_post_without_csrf_is_rejected() -> None:
    raw_client = TestClient(create_app())
    response = raw_client.post("/api/v1/auth/login", json=_LOGIN_JSON)

    assert response.status_code == 403
    assert response.json()["detail"] == "CSRF token missing or invalid."


def test_api_login_with_csrf_header_succeeds(client: CsrfAwareTestClient) -> None:
    response = client.post("/api/v1/auth/login", json=_LOGIN_JSON)

    assert response.status_code == 200
    assert response.json()["user"]["username"] == "admin"
    assert client.cookies.get(SESSION_COOKIE_NAME)
    assert client.cookies.get(CSRF_COOKIE_NAME)


def test_api_logout_requires_csrf(client: CsrfAwareTestClient) -> None:
    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200

    raw_client = TestClient(create_app())
    raw_client.cookies.update(client.cookies)
    blocked = raw_client.post("/api/v1/auth/logout")
    assert blocked.status_code == 403

    logout = client.post("/api/v1/auth/logout")
    assert logout.status_code == 200
    raw_cookie = logout.headers.get("set-cookie", "")
    assert CSRF_COOKIE_NAME in raw_cookie


def test_web_login_with_form_csrf_succeeds(client: CsrfAwareTestClient) -> None:
    response = client.post("/web/login", data=_LOGIN_FORM, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == "/commercial-offer/new"
    assert client.cookies.get(SESSION_COOKIE_NAME)


def test_web_login_without_csrf_is_rejected() -> None:
    raw_client = TestClient(create_app())
    response = raw_client.post("/web/login", data=_LOGIN_FORM, follow_redirects=False)

    assert response.status_code == 403


def test_login_rotates_csrf_token(client: CsrfAwareTestClient) -> None:
    bootstrap = client.get("/health")
    assert bootstrap.status_code == 200
    old_token = client.cookies.get(CSRF_COOKIE_NAME)
    assert old_token

    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200
    new_token = client.cookies.get(CSRF_COOKIE_NAME)
    assert new_token
    assert new_token != old_token

    set_cookie_headers = login.headers.get_list("set-cookie")
    csrf_set_cookie = next(h for h in set_cookie_headers if h.startswith(f"{CSRF_COOKIE_NAME}="))
    assert re.search(rf"{CSRF_COOKIE_NAME}=[^;]+", csrf_set_cookie)


def test_ensure_csrf_cookie_helper(client: CsrfAwareTestClient) -> None:
    token = ensure_csrf_cookie(client)
    assert token == client.cookies.get(CSRF_COOKIE_NAME)
    assert CSRF_HEADER_NAME in csrf_headers(client)
