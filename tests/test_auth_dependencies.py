from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi import Request
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.dependencies.auth import get_current_user
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from tests.helpers.auth_fixtures import patch_auth_users

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"

USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


def test_get_current_user_uses_get_user_by_id_not_list_users(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    db_path = str(tmp_path / "auth.db")
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, USERS)

    list_users_mock = MagicMock(side_effect=AssertionError("list_users must not be called"))
    monkeypatch.setattr(AuthRepository, "list_users", list_users_mock)

    token = create_session_token(
        {"id": 1, "username": "admin", "role": "admin", "sv": 0},
        ttl_seconds=300,
    )
    request = Request(
        {
            "type": "http",
            "headers": [],
            "method": "GET",
            "path": "/",
        }
    )
    request._cookies = {"app_session": token}

    user = get_current_user(request, AuthRepository(db_path))

    assert user["username"] == "admin"
    list_users_mock.assert_not_called()


def test_get_current_user_via_api_session(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, USERS)

    client = TestClient(create_app())
    token = create_session_token(
        {"id": 1, "username": "admin", "role": "admin", "sv": 0},
        ttl_seconds=300,
    )
    response = client.get("/api/v1/health", cookies={"app_session": token})

    assert response.status_code == 200
