from __future__ import annotations

import time
from typing import Any

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.login_rate_limit import reset_password_change_rate_limiter_for_tests
from tests.helpers.auth_fixtures import patch_auth_login

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
_LOGIN_JSON = {"username": "admin", "password": "StrongPassword123!"}
_CHANGE_PASSWORD_JSON = {
    "current_password": "WrongCurrentPassword1!",
    "new_password": "DifferentPassword123!",
}


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _logged_in_client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    patch_auth_login(monkeypatch)
    client = CsrfAwareTestClient(create_app())
    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200
    return client


def test_password_change_rate_limit_allows_three_attempts_then_429(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_PASSWORD_CHANGE_ATTEMPTS", "3")
    get_settings.cache_clear()
    reset_password_change_rate_limiter_for_tests()
    client = _logged_in_client(monkeypatch)

    for _ in range(3):
        response = client.post("/api/v1/auth/change-password", json=_CHANGE_PASSWORD_JSON)
        assert response.status_code == 401
        assert response.status_code != 429

    blocked = client.post("/api/v1/auth/change-password", json=_CHANGE_PASSWORD_JSON)
    assert blocked.status_code == 429
    assert blocked.json()["detail"] == "Слишком много попыток смены пароля. Повторите позже."
    retry_after = blocked.headers.get("Retry-After")
    assert retry_after is not None
    assert int(retry_after) >= 1


def test_password_change_rate_limit_is_per_ip(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_PASSWORD_CHANGE_ATTEMPTS", "2")
    monkeypatch.setenv("TRUSTED_PROXY_IPS", "testclient")
    get_settings.cache_clear()
    reset_password_change_rate_limiter_for_tests()
    client = _logged_in_client(monkeypatch)

    for ip in ("203.0.113.1", "203.0.113.1"):
        response = client.post(
            "/api/v1/auth/change-password",
            json=_CHANGE_PASSWORD_JSON,
            headers={"X-Forwarded-For": ip},
        )
        assert response.status_code == 401

    blocked = client.post(
        "/api/v1/auth/change-password",
        json=_CHANGE_PASSWORD_JSON,
        headers={"X-Forwarded-For": "203.0.113.1"},
    )
    assert blocked.status_code == 429

    allowed = client.post(
        "/api/v1/auth/change-password",
        json=_CHANGE_PASSWORD_JSON,
        headers={"X-Forwarded-For": "203.0.113.2"},
    )
    assert allowed.status_code == 401
    assert allowed.status_code != 429


def test_password_change_rate_limit_is_per_user(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_PASSWORD_CHANGE_ATTEMPTS", "2")
    get_settings.cache_clear()
    reset_password_change_rate_limiter_for_tests()

    users: list[dict[str, Any]] = [
        {
            "id": 1,
            "username": "alice",
            "role": "admin",
            "manager_id": None,
            "is_active": 1,
            "session_version": 0,
            "created_at": "2026-01-01 00:00:00",
        },
        {
            "id": 2,
            "username": "bob",
            "role": "admin",
            "manager_id": None,
            "is_active": 1,
            "session_version": 0,
            "created_at": "2026-01-01 00:00:00",
        },
    ]
    passwords = {"alice": "AlicePassword123!", "bob": "BobPassword123!"}

    def fake_authenticate(
        self: AuthRepository,
        username: str,
        password: str,
    ) -> dict[str, Any] | None:
        if passwords.get(username) == password:
            user = next(item for item in users if item["username"] == username)
            return dict(user)
        return None

    def fake_get_user_by_id(self: AuthRepository, user_id: int) -> dict[str, Any] | None:
        user = next((item for item in users if int(item["id"]) == int(user_id)), None)
        if user is None:
            return None
        return dict(user)

    monkeypatch.setattr(AuthRepository, "authenticate", fake_authenticate)
    monkeypatch.setattr(AuthRepository, "get_user_by_id", fake_get_user_by_id)

    alice_client = CsrfAwareTestClient(create_app())
    assert (
        alice_client.post(
            "/api/v1/auth/login",
            json={"username": "alice", "password": passwords["alice"]},
        ).status_code
        == 200
    )

    for _ in range(2):
        response = alice_client.post(
            "/api/v1/auth/change-password",
            json={
                "current_password": "WrongCurrentPassword1!",
                "new_password": "DifferentPassword123!",
            },
        )
        assert response.status_code == 401

    blocked = alice_client.post(
        "/api/v1/auth/change-password",
        json={
            "current_password": "WrongCurrentPassword1!",
            "new_password": "DifferentPassword123!",
        },
    )
    assert blocked.status_code == 429

    bob_client = CsrfAwareTestClient(create_app())
    assert (
        bob_client.post(
            "/api/v1/auth/login",
            json={"username": "bob", "password": passwords["bob"]},
        ).status_code
        == 200
    )
    allowed = bob_client.post(
        "/api/v1/auth/change-password",
        json={
            "current_password": "WrongCurrentPassword1!",
            "new_password": "DifferentPassword123!",
        },
    )
    assert allowed.status_code == 401
    assert allowed.status_code != 429


def test_password_change_rate_limit_window_resets(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_PASSWORD_CHANGE_ATTEMPTS", "2")
    monkeypatch.setenv("AUTH_PASSWORD_CHANGE_WINDOW_SECONDS", "1")
    get_settings.cache_clear()
    reset_password_change_rate_limiter_for_tests()
    client = _logged_in_client(monkeypatch)

    for _ in range(2):
        assert (
            client.post("/api/v1/auth/change-password", json=_CHANGE_PASSWORD_JSON).status_code
            == 401
        )

    assert client.post("/api/v1/auth/change-password", json=_CHANGE_PASSWORD_JSON).status_code == 429

    time.sleep(1.1)

    allowed = client.post("/api/v1/auth/change-password", json=_CHANGE_PASSWORD_JSON)
    assert allowed.status_code == 401
    assert allowed.status_code != 429
