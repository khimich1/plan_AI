"""Shared TestClient fixtures for commercial HTTP tests.

Keep these out of ``test_*.py`` modules so pytest can collect
``test_commercial_web_flow.py`` in the same session without losing fixtures.
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.security.session import create_session_token
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "true")
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    return CsrfAwareTestClient(create_app())


@pytest.fixture()
def auth_cookie(client: TestClient) -> dict[str, str]:
    token = create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    return {"app_session": token}


@pytest.fixture()
def client_two_users(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "alice",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            },
            {
                "id": 2,
                "username": "bob",
                "role": "manager",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            },
        ],
    )
    return CsrfAwareTestClient(create_app())
