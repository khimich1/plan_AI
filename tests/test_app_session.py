from __future__ import annotations

import re
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi import Response
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from tests.helpers.auth_fixtures import patch_auth_login
from app.security.session import (
    SESSION_COOKIE_NAME,
    clear_session_cookie,
    create_session_token,
    decode_session_token,
    is_session_active,
    session_claims_from_user,
    session_cookie_policy,
    set_session_cookie,
    token_session_version,
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
    session_version: int = 0,
) -> None:
    patch_auth_login(
        monkeypatch,
        username=username,
        password=password,
        session_version=session_version,
    )


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    _mock_authenticate_success(monkeypatch)
    return CsrfAwareTestClient(create_app())


def test_session_roundtrip() -> None:
    token = create_session_token(
        session_claims_from_user(
            {"id": 1, "username": "demo", "role": "admin", "session_version": 2}
        ),
        ttl_seconds=60,
    )

    payload = decode_session_token(token)

    assert payload is not None
    assert payload["id"] == 1
    assert payload["username"] == "demo"
    assert payload["role"] == "admin"
    assert token_session_version(payload) == 2


def test_is_session_active_matches_user_session_version() -> None:
    payload = {"sv": 3}
    user = {"session_version": 3}
    stale_user = {"session_version": 4}

    assert is_session_active(payload, user) is True
    assert is_session_active(payload, stale_user) is False
    assert is_session_active({"id": 1}, {"session_version": 0}) is True


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
    assert response.headers["location"] == "/commercial-offer/new"
    raw_cookie = response.headers.get("set-cookie", "")
    attrs = _parse_set_cookie_header(raw_cookie)
    assert attrs.get("name") == SESSION_COOKIE_NAME
    assert attrs.get("httponly") is True
    assert attrs.get("secure") is True
    assert attrs.get("samesite") == "lax"
    assert attrs.get("max_age") == 86400


def test_web_logout_get_does_not_clear_session(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _configure_cookie_settings(monkeypatch, cookie_samesite="lax", cookie_secure="false")
    login = client.post(
        "/api/v1/auth/login",
        json={"username": "admin", "password": "StrongPassword123!"},
    )
    assert login.status_code == 200

    legacy_logout = client.get("/web/logout", follow_redirects=False)

    assert legacy_logout.status_code == 405
    assert "POST /api/v1/auth/logout" in legacy_logout.json()["detail"]
    assert legacy_logout.headers.get("allow") == "POST"
    assert legacy_logout.headers.get("deprecation") == "true"

    me = client.get("/api/v1/auth/me")
    assert me.status_code == 200
    assert me.json()["user"]["username"] == "admin"


def test_stale_session_cookie_rejected_after_logout(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    login = client.post(
        "/api/v1/auth/login",
        json={"username": "admin", "password": "StrongPassword123!"},
    )
    assert login.status_code == 200
    stale_cookie = client.cookies.get(SESSION_COOKIE_NAME)
    assert stale_cookie

    logout = client.post("/api/v1/auth/logout")
    assert logout.status_code == 200

    replay = TestClient(create_app())
    replay.cookies.set(SESSION_COOKIE_NAME, stale_cookie)
    me = replay.get("/api/v1/auth/me")
    assert me.status_code == 401
    assert me.json()["detail"] == "Session expired"


def test_password_change_invalidates_old_session_cookie(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    db_path = str(tmp_path / "auth.db")
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()

    repository = AuthRepository(db_path=db_path)
    repository.create_or_update_user(
        username="alice",
        password="CurrentPassword123!",
        role="admin",
    )

    client = CsrfAwareTestClient(create_app())
    login = client.post(
        "/api/v1/auth/login",
        json={"username": "alice", "password": "CurrentPassword123!"},
    )
    assert login.status_code == 200
    old_cookie = client.cookies.get(SESSION_COOKIE_NAME)
    assert old_cookie

    change = client.post(
        "/api/v1/auth/change-password",
        json={
            "current_password": "CurrentPassword123!",
            "new_password": "UpdatedPassword123!",
        },
    )
    assert change.status_code == 200
    new_cookie = client.cookies.get(SESSION_COOKIE_NAME)
    assert new_cookie
    assert new_cookie != old_cookie

    replay = TestClient(create_app())
    replay.cookies.set(SESSION_COOKIE_NAME, old_cookie)
    me = replay.get("/api/v1/auth/me")
    assert me.status_code == 401
    assert me.json()["detail"] == "Session expired"

    me_new = client.get("/api/v1/auth/me")
    assert me_new.status_code == 200
    assert me_new.json()["user"]["username"] == "alice"
