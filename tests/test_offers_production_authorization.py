from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token
from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"

USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
]

ORDER_DATA = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "length_dm_raw": "60",
    }
]


def _session_cookie(user_id: int, role: str, username: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


def _seed_in_production_kp(db_path: str) -> int:
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.03.2026",
        ORDER_DATA,
        customer_name="Производственный клиент",
        manager_name="Менеджер",
        status="в работе",
        owner_user_id=3,
        db_path=db_path,
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 1, 'ПБ 60-12-8п', 6.0, 1.2, 800, 1, 'в производстве')
            """,
            (kp_id,),
        )
        conn.commit()
    return kp_id


@pytest.fixture()
def auth_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[TestClient, str]:
    db_path = fx.make_iso_db(tmp_path)
    _seed_in_production_kp(db_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, USERS)
    client = TestClient(create_app())
    return client, db_path


def test_production_forbidden_on_offers_list(auth_client: tuple[TestClient, str]) -> None:
    client, _ = auth_client
    response = client.get(
        "/api/v1/offers",
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_production_forbidden_on_offer_get(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    kp_id = _seed_in_production_kp(db_path)
    response = client.get(
        f"/api/v1/offers/{kp_id}",
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 403


def test_production_forbidden_on_offer_pdf(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    kp_id = _seed_in_production_kp(db_path)
    response = client.get(
        f"/api/v1/offers/{kp_id}/pdf",
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 403


def test_production_allowed_on_kp_candidates(auth_client: tuple[TestClient, str]) -> None:
    client, _ = auth_client
    response = client.get(
        "/api/v1/production/kp-candidates",
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 200
    payload = response.json()
    assert "items" in payload
    assert len(payload["items"]) >= 1


def test_manager_can_list_all_offers(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    KpPersistenceService.save_kp_to_db(
        "01.03.2026",
        ORDER_DATA,
        customer_name="Менеджерский клиент",
        manager_name="Менеджер",
        status="в архиве",
        owner_user_id=3,
        db_path=db_path,
    )
    KpPersistenceService.save_kp_to_db(
        "01.03.2026",
        ORDER_DATA,
        customer_name="Чужой архивный",
        manager_name="Менеджер",
        status="в архиве",
        owner_user_id=1,
        db_path=db_path,
    )
    response = client.get(
        "/api/v1/offers?status=archived",
        cookies=_session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 200
    assert response.json()["count"] >= 2


def test_admin_can_list_offers(auth_client: tuple[TestClient, str]) -> None:
    client, _ = auth_client
    response = client.get(
        "/api/v1/offers",
        cookies=_session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 200
    assert response.json()["count"] >= 1
