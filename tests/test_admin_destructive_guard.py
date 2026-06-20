"""HTTP-level tests for destructive admin reset guards (WP2 S6)."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.http_errors import MSG_DESTRUCTIVE_DB_BLOCKED
from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token
from core import kp_db

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
ADMIN_RESET_FULL = "/api/v1/admin/db/reset/full"
ADMIN_RESET_KP = "/api/v1/admin/db/reset/kp-only"
ADMIN_RESET_PLANS = "/api/v1/admin/db/reset/plans-only"
ADMIN_RESET_CALENDAR = "/api/v1/admin/db/reset/calendar-only"

ALL_RESET_ENDPOINTS = (
    ADMIN_RESET_FULL,
    ADMIN_RESET_KP,
    ADMIN_RESET_PLANS,
    ADMIN_RESET_CALENDAR,
)

ADMIN_USER = {
    "id": 1,
    "username": "admin",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}


@pytest.fixture()
def production_admin_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_path = tmp_path / "plita.db"
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    calendar_path = tmp_path / "work_calendar.json"
    calendar_path.write_text(
        '{"extra_holidays": ["2026-01-01"], "extra_workdays": []}',
        encoding="utf-8",
    )

    kp_db.init_schema(str(db_path))
    AuthRepository(db_path=str(db_path)).create_or_update_user(
        username="admin",
        password="AdminTestPass12!",
        role="admin",
    )

    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    monkeypatch.setenv("PB_DB_PATH", str(db_path))
    monkeypatch.setenv("WORK_CALENDAR_PATH", str(calendar_path))
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("BOT_AUTH_ENABLED", "true")
    monkeypatch.setenv("BOT_TELEGRAM_ALLOWLIST", "1:admin")
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)
    get_settings.cache_clear()

    patch_auth_users(monkeypatch, [ADMIN_USER])
    return db_path


@pytest.fixture()
def client(production_admin_env: Path) -> TestClient:
    del production_admin_env
    return CsrfAwareTestClient(create_app())


@pytest.fixture()
def admin_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": ADMIN_USER["id"], "username": ADMIN_USER["username"], "role": ADMIN_USER["role"]},
            ttl_seconds=300,
        )
    }


@pytest.mark.parametrize("endpoint", ALL_RESET_ENDPOINTS)
def test_reset_blocked_in_production_without_flag(
    client: TestClient,
    admin_cookie: dict[str, str],
    endpoint: str,
) -> None:
    response = client.post(endpoint, cookies=admin_cookie)

    assert response.status_code == 403
    assert response.json()["detail"] == MSG_DESTRUCTIVE_DB_BLOCKED


@pytest.mark.parametrize("endpoint", ALL_RESET_ENDPOINTS)
def test_reset_blocked_in_production_with_allow_only(
    client: TestClient,
    admin_cookie: dict[str, str],
    endpoint: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    monkeypatch.delenv("DESTRUCTIVE_DB_RESET_BREAK_GLASS", raising=False)

    response = client.post(endpoint, cookies=admin_cookie)

    assert response.status_code == 403
    assert response.json()["detail"] == MSG_DESTRUCTIVE_DB_BLOCKED
