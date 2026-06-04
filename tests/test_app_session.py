from __future__ import annotations

import re
from unittest.mock import MagicMock

import pytest
from fastapi import Response
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import (
    SESSION_COOKIE_NAME,
    clear_session_cookie,
    create_session_token,
    decode_session_token,
    session_cookie_policy,
    set_session_cookie,
)

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _configure_cookie_settings(
    monkeypatch: pytest.MonkeyPatch,
    *,
    app_env: str = "development",
    cookie_secure: str | None = None,
    cookie_samesite: str = "lax",
    session_max_age: str = "7200",
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("APP_ENV", app_env)
    monkeypatch.setenv("COOKIE_SAMESITE", cookie_samesite)
    monkeypatch.setenv("SESSION_COOKIE_MAX_AGE", session_max_age)
    if cookie_secure is None:
        monkeypatch.delenv("COOKIE_SECURE", raising=False)
    else:
        monkeypatch.setenv("COOKIE_SECURE", cookie_secure)
    if app_env.lower() == "production":
        monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
        monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    else:
        monkeypatch.delenv("BOT_TELEGRAM_ALLOWLIST", raising=False)
    get_settings.cache_clear()


def _parse_set_cookie_header(header: str) -> dict[str, str | bool | int]:
    """Parse ``Set-Cookie`` attributes (case-insensitive)."""
    attrs: dict[str, str | bool | int] = {}
    parts = [part.strip() for part in header.split(";")]
    if parts:
        name_value = parts[0]
        if "=" in name_value:
            name, value = name_value.split("=", 1)
            attrs["name"] = name
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] == '"':
                value = value[1:-1]
            attrs["value"] = value
    for part in parts[1:]:
        lower = part.lower()
        if lower == "httponly":
            attrs["httponly"] = True
        elif lower == "secure":
            attrs["secure"] = True
        elif lower.startswith("samesite="):
            attrs["samesite"] = part.split("=", 1)[1].lower()
        elif lower.startswith("max-age="):
            attrs["max_age"] = int(part.split("=", 1)[1])
    return attrs


def _mock_authenticate_success(
    monkeypatch: pytest.MonkeyPatch,
    *,
    username: str = "admin",
    password: str = "StrongPassword123!",
) -> None:
    def fake_authenticate(self, user: str, pwd: str) -> dict | None:
        if user == username and pwd == password:
            return {"id": 1, "username": username, "role": "admin"}
        return None

    monkeypatch.setattr(AuthRepository, "authenticate", fake_authenticate)


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    _mock_authenticate_success(monkeypatch)
    return TestClient(create_app())


def test_session_roundtrip() -> None:
    token = create_session_token({"id": 1, "username": "demo", "role": "admin"}, ttl_seconds=60)

    payload = decode_session_token(token)

    assert payload is not None
    assert payload["id"] == 1
    assert payload["username"] == "demo"
    assert payload["role"] == "admin"


def test_session_cookie_policy_reflects_settings(monkeypatch: pytest.MonkeyPatch) -> None:
    _configure_cookie_settings(
        monkeypatch,
        app_env="production",
        cookie_secure="true",
        cookie_samesite="strict",
        session_max_age="3600",
    )

    policy = session_cookie_policy()

    assert policy == {
        "httponly": True,
        "samesite": "strict",
        "secure": True,
        "max_age": 3600,
    }


def test_session_cookie_policy_secure_defaults_from_app_env(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(monkeypatch, app_env="development", cookie_samesite="lax")

    assert session_cookie_policy()["secure"] is False

    _configure_cookie_settings(monkeypatch, app_env="production", cookie_samesite="lax")

    assert session_cookie_policy()["secure"] is True


def test_set_session_cookie_applies_policy(monkeypatch: pytest.MonkeyPatch) -> None:
    _configure_cookie_settings(
        monkeypatch,
        cookie_secure="true",
        cookie_samesite="none",
        session_max_age="1800",
    )
    response = MagicMock(spec=Response)

    set_session_cookie(response, "signed-token")

    response.set_cookie.assert_called_once_with(
        SESSION_COOKIE_NAME,
        "signed-token",
        httponly=True,
        samesite="none",
        secure=True,
        max_age=1800,
    )


def test_clear_session_cookie_applies_policy(monkeypatch: pytest.MonkeyPatch) -> None:
    _configure_cookie_settings(
        monkeypatch,
        cookie_secure="false",
        cookie_samesite="lax",
        session_max_age="7200",
    )
    response = MagicMock(spec=Response)

    clear_session_cookie(response)

    response.delete_cookie.assert_called_once_with(
        SESSION_COOKIE_NAME,
        httponly=True,
        samesite="lax",
        secure=False,
    )


def test_api_login_set_cookie_attributes(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(
        monkeypatch,
        app_env="development",
        cookie_secure="true",
        cookie_samesite="strict",
        session_max_age="43200",
    )

    response = client.post(
        "/api/v1/auth/login",
        json={"username": "admin", "password": "StrongPassword123!"},
    )

    assert response.status_code == 200
    raw_cookie = response.headers.get("set-cookie", "")
    assert raw_cookie.startswith(f"{SESSION_COOKIE_NAME}=")
    attrs = _parse_set_cookie_header(raw_cookie)
    assert attrs.get("httponly") is True
    assert attrs.get("secure") is True
    assert attrs.get("samesite") == "strict"
    assert attrs.get("max_age") == 43200
    assert attrs.get("value")
    assert decode_session_token(str(attrs["value"])) is not None


def test_api_logout_clears_session_cookie(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(monkeypatch, cookie_samesite="lax", cookie_secure="false")
    login = client.post(
        "/api/v1/auth/login",
        json={"username": "admin", "password": "StrongPassword123!"},
    )
    assert login.status_code == 200

    logout = client.post("/api/v1/auth/logout")

    assert logout.status_code == 200
    raw_cookie = logout.headers.get("set-cookie", "")
    assert SESSION_COOKIE_NAME in raw_cookie
    assert re.search(rf"{SESSION_COOKIE_NAME}=;", raw_cookie) or '=""' in raw_cookie.replace(" ", "")


def test_web_login_set_cookie_attributes(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(
        monkeypatch,
        app_env="production",
        cookie_samesite="lax",
        session_max_age="86400",
    )

    response = client.post(
        "/web/login",
        data={"username": "admin", "password": "StrongPassword123!"},
        follow_redirects=False,
    )

    assert response.status_code == 303
    raw_cookie = response.headers.get("set-cookie", "")
    attrs = _parse_set_cookie_header(raw_cookie)
    assert attrs.get("name") == SESSION_COOKIE_NAME
    assert attrs.get("httponly") is True
    assert attrs.get("secure") is True
    assert attrs.get("samesite") == "lax"
    assert attrs.get("max_age") == 86400


def test_web_logout_clears_session_cookie(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(monkeypatch, cookie_samesite="strict", cookie_secure="true")
    login = client.post(
        "/web/login",
        data={"username": "admin", "password": "StrongPassword123!"},
        follow_redirects=False,
    )
    assert login.status_code == 303
    client.cookies.update(login.cookies)

    logout = client.get("/web/logout", follow_redirects=False)

    assert logout.status_code == 303
    raw_cookie = logout.headers.get("set-cookie", "")
    assert SESSION_COOKIE_NAME in raw_cookie
    assert "samesite=strict" in raw_cookie.lower()
    assert "secure" in raw_cookie.lower()
