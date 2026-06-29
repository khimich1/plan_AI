"""HTTP tests for paginated admin user listing (S16)."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from tests.helpers.auth_fixtures import patch_auth_users

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
ADMIN_USERS_PATH = "/api/v1/admin/users"

ADMIN_USER = {
    "id": 1,
    "username": "admin",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}
MANAGER_USER = {
    "id": 2,
    "username": "manager",
    "role": "manager",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}


@pytest.fixture()
def users_db(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> AuthRepository:
    db_path = tmp_path / "plita.db"
    repository = AuthRepository(db_path=str(db_path))
    repository.create_or_update_user(
        username="admin",
        password="AdminTestPass12!",
        role="admin",
    )
    for index in range(4):
        repository.create_or_update_user(
            username=f"manager_{index:02d}",
            password="ManagerTestPass1!",
            role="manager",
        )

    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    monkeypatch.setenv("PB_DB_PATH", str(db_path))
    get_settings.cache_clear()
    return repository


@pytest.fixture()
def client(users_db: AuthRepository) -> TestClient:
    del users_db
    return TestClient(create_app())


def _session_cookie(user: dict[str, object]) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {
                "id": user["id"],
                "username": user["username"],
                "role": user["role"],
            },
            ttl_seconds=300,
        )
    }


def test_admin_users_endpoint_returns_paginated_page(client: TestClient) -> None:
    response = client.get(
        ADMIN_USERS_PATH,
        params={"limit": 2, "offset": 1},
        cookies=_session_cookie(ADMIN_USER),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["total"] == 5
    assert payload["limit"] == 2
    assert payload["offset"] == 1
    assert len(payload["items"]) == 2
    assert payload["items"][0]["username"] == "manager_00"
    assert "password_hash" not in payload["items"][0]


def test_admin_users_endpoint_defaults_to_limit_50(client: TestClient) -> None:
    response = client.get(ADMIN_USERS_PATH, cookies=_session_cookie(ADMIN_USER))

    assert response.status_code == 200
    payload = response.json()
    assert payload["limit"] == 50
    assert payload["offset"] == 0
    assert payload["total"] == 5
    assert len(payload["items"]) == 5


def test_admin_users_endpoint_requires_admin_role(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    patch_auth_users(monkeypatch, [ADMIN_USER, MANAGER_USER])

    response = client.get(ADMIN_USERS_PATH, cookies=_session_cookie(MANAGER_USER))

    assert response.status_code == 403


def test_admin_users_endpoint_requires_authentication(client: TestClient) -> None:
    response = client.get(ADMIN_USERS_PATH)

    assert response.status_code == 401
