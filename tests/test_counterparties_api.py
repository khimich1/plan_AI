"""CTR-004: HTTP API /api/v1/counterparties."""

from __future__ import annotations

import pytest

from app.repositories.counterparties_repository import CounterpartiesRepository
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

API = "/api/v1/counterparties"

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
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


@pytest.fixture()
def counterparties_db(tmp_path, monkeypatch: pytest.MonkeyPatch) -> str:
    from app.core.settings import get_settings
    from tests.helpers.auth_fixtures import patch_auth_users

    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    repo = CounterpartiesRepository(db_path=db_path)
    repo.upsert(
        [
            {
                "code_1c": "00-pref",
                "name": "СТОЛИЦА ООО",
                "inn": "7701000001",
                "kpp": "770101001",
                "is_client": True,
            },
            {
                "code_1c": "00-sub",
                "name": "ТРЕСТ СТО ООО",
                "inn": "7702000002",
                "is_client": True,
            },
            {
                "code_1c": "00-sup",
                "name": "СТО Поставщик",
                "inn": "9900000000",
                "is_client": False,
            },
        ]
    )
    return db_path


@pytest.fixture()
def client(counterparties_db: str):
    from app.main import create_app
    from tests.helpers.csrf import CsrfAwareTestClient

    del counterparties_db
    return CsrfAwareTestClient(create_app())


def _admin() -> dict[str, str]:
    return session_cookie(1, "admin", "admin")


def _manager() -> dict[str, str]:
    return session_cookie(3, "manager", "manager_a")


def _production() -> dict[str, str]:
    return session_cookie(2, "production", "prod_user")


def test_search_returns_ranked_clients(client) -> None:
    response = client.get(f"{API}/search", params={"q": "сто"}, cookies=_manager())
    assert response.status_code == 200
    body = response.json()
    assert body["count"] <= 10
    codes = [item["code_1c"] for item in body["items"]]
    assert codes[:2] == ["00-pref", "00-sub"]
    assert "00-sup" not in codes


def test_search_q_too_short_is_422(client) -> None:
    response = client.get(f"{API}/search", params={"q": "с"}, cookies=_manager())
    assert response.status_code == 422


def test_search_forbidden_for_production(client) -> None:
    response = client.get(f"{API}/search", params={"q": "сто"}, cookies=_production())
    assert response.status_code == 403


def test_search_manager_clients_only_false_still_hides_suppliers(client) -> None:
    response = client.get(
        f"{API}/search",
        params={"q": "сто", "clients_only": "false"},
        cookies=_manager(),
    )
    assert response.status_code == 200
    codes = [item["code_1c"] for item in response.json()["items"]]
    assert "00-sup" not in codes


def test_search_admin_clients_only_false_sees_suppliers(client) -> None:
    response = client.get(
        f"{API}/search",
        params={"q": "сто", "clients_only": "false"},
        cookies=_admin(),
    )
    assert response.status_code == 200
    codes = [item["code_1c"] for item in response.json()["items"]]
    assert "00-sup" in codes


def test_search_q_too_long_is_422(client) -> None:
    response = client.get(
        f"{API}/search", params={"q": "а" * 129}, cookies=_manager()
    )
    assert response.status_code == 422


def test_create_name_too_long_is_422(client) -> None:
    response = client.post(
        API,
        json={"name": "Н" * 256, "code_1c": "00-long"},
        cookies=_manager(),
    )
    assert response.status_code == 422


def test_search_empty_items(client) -> None:
    response = client.get(
        f"{API}/search", params={"q": "неттакого"}, cookies=_admin()
    )
    assert response.status_code == 200
    assert response.json() == {"items": [], "count": 0}


def test_create_201_and_immediately_searchable(client) -> None:
    response = client.post(
        API,
        json={"name": "НОВАЯ РОМАШКА", "code_1c": "00-new", "inn": "7800000000"},
        cookies=_manager(),
    )
    assert response.status_code == 201
    item = response.json()["item"]
    assert item["code_1c"] == "00-new"
    assert response.json().get("warning") is None

    found = client.get(
        f"{API}/search", params={"q": "ромашка"}, cookies=_manager()
    )
    assert found.status_code == 200
    assert any(row["code_1c"] == "00-new" for row in found.json()["items"])


def test_create_duplicate_code_409(client) -> None:
    response = client.post(
        API,
        json={"name": "ДРУГАЯ", "code_1c": "00-pref"},
        cookies=_admin(),
    )
    assert response.status_code == 409
    payload = response.json()
    detail = payload.get("detail", payload)
    if isinstance(detail, dict) and "existing" in detail:
        existing = detail["existing"]
        message = detail.get("detail") or detail.get("message") or ""
    else:
        existing = payload["existing"]
        message = payload.get("detail") or ""
    assert "уже есть" in str(message)
    assert existing["code_1c"] == "00-pref"


def test_create_duplicate_inn_201_with_warning(client) -> None:
    response = client.post(
        API,
        json={
            "name": "КЛОН",
            "code_1c": "00-clone",
            "inn": "7701000001",
        },
        cookies=_manager(),
    )
    assert response.status_code == 201
    assert response.json()["warning"]
    assert "ИНН" in response.json()["warning"]
