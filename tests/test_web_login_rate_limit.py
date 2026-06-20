from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.login_rate_limit import reset_login_rate_limiter_for_tests

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
_LOGIN_FORM = {"username": "admin", "password": "StrongPassword123!"}


@pytest.fixture(autouse=True)
def _valid_secret_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _mock_authenticate_success(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_authenticate(self, user: str, pwd: str) -> dict | None:
        if user == _LOGIN_FORM["username"] and pwd == _LOGIN_FORM["password"]:
            return {"id": 1, "username": user, "role": "admin"}
        return None

    monkeypatch.setattr(AuthRepository, "authenticate", fake_authenticate)


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    _mock_authenticate_success(monkeypatch)
    return CsrfAwareTestClient(create_app())


def test_web_login_rate_limit_blocks_sixth_attempt(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("AUTH_LOGIN_ATTEMPTS_PER_MINUTE", "5")
    get_settings.cache_clear()
    reset_login_rate_limiter_for_tests()

    for _ in range(5):
        response = client.post("/web/login", data=_LOGIN_FORM, follow_redirects=False)
        assert response.status_code == 303

    blocked = client.post("/web/login", data=_LOGIN_FORM, follow_redirects=False)
    assert blocked.status_code == 429
    assert blocked.json()["detail"] == "Слишком много попыток входа. Повторите позже."
    retry_after = blocked.headers.get("Retry-After")
    assert retry_after is not None
    assert int(retry_after) >= 1
