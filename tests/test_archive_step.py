"""Archive API: product_type filter and step KP details."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.services.archive_service import ArchiveService
from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"

ADMIN_USER = {
    "id": 1,
    "username": "admin",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}

STEP_ORDER = [
    {
        "product_kind": "step",
        "name": "ЛС11",
        "mark": "ЛС11",
        "qty": 3,
        "unit_price": 1409.91,
    }
]


def _admin_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": 1, "username": "admin", "role": "admin"},
            ttl_seconds=300,
        )
    }


@pytest.fixture()
def archive_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[TestClient, str]:
    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, [ADMIN_USER])
    return TestClient(create_app()), db_path


def _save_step_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        STEP_ORDER,
        customer_name="Step client",
        status="в архиве",
        product_type="steps",
        db_path=db_path,
    )


def test_archive_list_filter_product_type_steps(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    step_id = _save_step_kp(db_path)

    steps_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=steps",
        cookies=_admin_cookie(),
    )
    assert steps_resp.status_code == 200
    step_items = steps_resp.json()
    assert len(step_items) == 1
    assert step_items[0]["kp_id"] == step_id
    assert step_items[0]["product_type"] == "steps"


def test_archive_detail_includes_steps(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    step_id = _save_step_kp(db_path)

    response = client.get(
        f"/api/v1/commercial/archive/{step_id}",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["product_type"] == "steps"
    assert payload["plates"] == []
    assert payload["piles"] == []
    assert len(payload["steps"]) == 1
    step = payload["steps"][0]
    assert step["mark"] == "ЛС11"
    assert step["qty"] == 3
    assert step["unit_price"] == pytest.approx(1409.91)


def test_archive_service_maps_step_details(tmp_path: Path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    step_id = _save_step_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path,
    )
    details = service.get_details(step_id, user={"id": 1, "role": "admin"})

    assert details.product_type == "steps"
    assert len(details.steps) == 1
    assert details.steps[0].mark == "ЛС11"
    assert details.plates == []
    assert details.piles == []
