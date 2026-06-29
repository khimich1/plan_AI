from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from app.security.login_rate_limit import reset_login_rate_limiter_for_tests
from tests.helpers.auth_fixtures import patch_auth_login

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
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    _mock_authenticate_success(monkeypatch)
    return CsrfAwareTestClient(create_app())


def test_login_rate_limit_allows_five_attempts_then_429(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_LOGIN_ATTEMPTS_PER_MINUTE", "5")
    get_settings.cache_clear()
    reset_login_rate_limiter_for_tests()

    for _ in range(5):
        response = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
        assert response.status_code == 200

    blocked = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert blocked.status_code == 429
    assert blocked.json()["detail"] == "Слишком много попыток входа. Повторите позже."
    retry_after = blocked.headers.get("Retry-After")
    assert retry_after is not None
    assert int(retry_after) >= 1


def test_login_rate_limit_is_per_ip(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_LOGIN_ATTEMPTS_PER_MINUTE", "2")
    monkeypatch.setenv("TRUSTED_PROXY_IPS", "testclient")
    get_settings.cache_clear()
    reset_login_rate_limiter_for_tests()

    assert (
        client.post(
            "/api/v1/auth/login",
            json=_LOGIN_JSON,
            headers={"X-Forwarded-For": "203.0.113.1"},
        ).status_code
        == 200
    )
    assert (
        client.post(
            "/api/v1/auth/login",
            json=_LOGIN_JSON,
            headers={"X-Forwarded-For": "203.0.113.1"},
        ).status_code
        == 200
    )
    assert (
        client.post(
            "/api/v1/auth/login",
            json=_LOGIN_JSON,
            headers={"X-Forwarded-For": "203.0.113.1"},
        ).status_code
        == 429
    )
    assert (
        client.post(
            "/api/v1/auth/login",
            json=_LOGIN_JSON,
            headers={"X-Forwarded-For": "203.0.113.2"},
        ).status_code
        == 200
    )


def test_login_rate_limit_does_not_apply_to_me_or_logout(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_LOGIN_ATTEMPTS_PER_MINUTE", "1")
    get_settings.cache_clear()
    reset_login_rate_limiter_for_tests()

    login = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert login.status_code == 200

    blocked = client.post("/api/v1/auth/login", json=_LOGIN_JSON)
    assert blocked.status_code == 429

    me = client.get("/api/v1/auth/me")
    assert me.status_code == 200
    assert me.status_code != 429

    logout = client.post("/api/v1/auth/logout")
    assert logout.status_code == 200
    assert logout.status_code != 429
