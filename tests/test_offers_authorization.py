from __future__ import annotations

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
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_b",
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


def _save_owned_offer(db_path: str, *, owner_user_id: int, customer_name: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.03.2026",
        ORDER_DATA,
        customer_name=customer_name,
        manager_name="Менеджер",
        status="в архиве",
        owner_user_id=owner_user_id,
        db_path=db_path,
    )


@pytest.fixture()
def auth_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[TestClient, str]:
    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, USERS)
    client = TestClient(create_app())
    return client, db_path


def test_manager_list_excludes_other_managers_offers(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    _save_owned_offer(db_path, owner_user_id=2, customer_name="Клиент A")
    _save_owned_offer(db_path, owner_user_id=3, customer_name="Клиент B")

    response = client.get(
        "/api/v1/offers?status=archived",
        cookies=_session_cookie(2, "manager", "manager_a"),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["count"] == 1
    assert payload["items"][0]["customer_name"] == "Клиент A"


def test_manager_cannot_get_foreign_offer(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    foreign_kp_id = _save_owned_offer(db_path, owner_user_id=3, customer_name="Чужой")

    response = client.get(
        f"/api/v1/offers/{foreign_kp_id}",
        cookies=_session_cookie(2, "manager", "manager_a"),
    )

    assert response.status_code == 403


def test_manager_cannot_patch_foreign_offer_discount(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    foreign_kp_id = _save_owned_offer(db_path, owner_user_id=3, customer_name="Чужой")

    response = client.patch(
        f"/api/v1/offers/{foreign_kp_id}/discount",
        json={"discount_percent": 10},
        cookies=_session_cookie(2, "manager", "manager_a"),
    )

    assert response.status_code == 403


def test_manager_cannot_delete_foreign_offer(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    foreign_kp_id = _save_owned_offer(db_path, owner_user_id=3, customer_name="Чужой")

    response = client.delete(
        f"/api/v1/offers/{foreign_kp_id}",
        cookies=_session_cookie(2, "manager", "manager_a"),
    )

    assert response.status_code == 403


def test_admin_sees_all_offers(auth_client: tuple[TestClient, str]) -> None:
    client, db_path = auth_client
    _save_owned_offer(db_path, owner_user_id=2, customer_name="Клиент A")
    _save_owned_offer(db_path, owner_user_id=3, customer_name="Клиент B")

    response = client.get(
        "/api/v1/offers?status=archived",
        cookies=_session_cookie(1, "admin", "admin"),
    )

    assert response.status_code == 200
    assert response.json()["count"] == 2
