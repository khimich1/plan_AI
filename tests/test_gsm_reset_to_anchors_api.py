"""HTTP tests for POST /api/v1/gsm/reset-to-anchors."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.http_errors import MSG_DESTRUCTIVE_DB_BLOCKED
from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie
from tests.test_reset_gsm_to_anchors import _fresh_db, _seed_two_vehicles_with_anchors

RESET_ENDPOINT = "/api/v1/gsm/reset-to-anchors"

TEST_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


@pytest.fixture()
def gsm_reset_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_path = _fresh_db(tmp_path, "gsm_reset_api.db")
    _seed_two_vehicles_with_anchors(db_path)

    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    monkeypatch.setenv("PB_DB_PATH", str(db_path))
    monkeypatch.setenv("APP_ENV", "development")
    monkeypatch.setenv("ALLOW_DESTRUCTIVE_DB_RESET", "1")
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return db_path


@pytest.fixture()
def client(gsm_reset_env: Path) -> CsrfAwareTestClient:
    del gsm_reset_env
    return CsrfAwareTestClient(create_app())


def _count(db_path: Path, table: str) -> int:
    with sqlite3.connect(str(db_path)) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    return int(row[0])


def test_reset_to_anchors_success_for_admin(client: CsrfAwareTestClient, gsm_reset_env: Path) -> None:
    response = client.post(
        RESET_ENDPOINT,
        cookies=session_cookie(1, "admin", "admin"),
    )

    assert response.status_code == 200
    body = response.json()
    assert body["anchors_kept"] == 2
    assert body["waybills_deleted"] == 3
    assert body["transactions_deleted"] == 2
    assert body["import_batches_deleted"] == 2
    assert len(body["anchors"]) == 2
    assert "bak-before-gsm-test-" in body["backup_path"]
    assert _count(gsm_reset_env, "gsm_waybill") == 2
    assert _count(gsm_reset_env, "gsm_transaction") == 0
    assert _count(gsm_reset_env, "gsm_import_batch") == 0


def test_reset_to_anchors_forbidden_for_accountant(client: CsrfAwareTestClient) -> None:
    response = client.post(
        RESET_ENDPOINT,
        cookies=session_cookie(5, "accountant", "accountant_user"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_reset_to_anchors_blocked_without_destructive_flag(
    client: CsrfAwareTestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("ALLOW_DESTRUCTIVE_DB_RESET", raising=False)
    get_settings.cache_clear()

    response = client.post(
        RESET_ENDPOINT,
        cookies=session_cookie(1, "admin", "admin"),
    )

    assert response.status_code == 403
    assert response.json()["detail"] == MSG_DESTRUCTIVE_DB_BLOCKED


def test_reset_to_anchors_missing_anchor_returns_422(
    client: CsrfAwareTestClient,
    gsm_reset_env: Path,
) -> None:
    with sqlite3.connect(str(gsm_reset_env)) as conn:
        conn.execute("DELETE FROM gsm_waybill WHERE vehicle_id = 2 AND source = 'imported'")
        conn.commit()

    response = client.post(
        RESET_ENDPOINT,
        cookies=session_cookie(1, "admin", "admin"),
    )

    assert response.status_code == 422
    body = response.json()["detail"]
    assert body["code"] == "gsm_reset_no_anchors"
    assert "нет imported-якоря" in body["message"]
