"""Archive API: product_type filter and march KP details."""

from __future__ import annotations

import asyncio
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

MARCH_ORDER = [
    {
        "product_kind": "march",
        "name": "1ЛМ 27-11-14-4",
        "mark": "1ЛМ 27-11-14-4",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 14391.41,
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


def _save_march_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        MARCH_ORDER,
        customer_name="March client",
        status="в архиве",
        product_type="marches",
        db_path=db_path,
    )


def test_archive_list_filter_marches(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    march_id = _save_march_kp(db_path)

    response = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=marches",
        cookies=_admin_cookie(),
    )
    assert response.status_code == 200
    items = response.json()
    assert len(items) == 1
    assert items[0]["kp_id"] == march_id
    assert items[0]["product_type"] == "marches"


def test_archive_detail_includes_marches(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    march_id = _save_march_kp(db_path)

    response = client.get(
        f"/api/v1/commercial/archive/{march_id}",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["product_type"] == "marches"
    assert payload["plates"] == []
    assert len(payload["marches"]) == 1
    march = payload["marches"][0]
    assert march["mark"] == "1ЛМ 27-11-14-4"
    assert march["concrete_grade"] == "B25"
    assert march["qty"] == 3
    assert march["unit_price"] == pytest.approx(14391.41)


def test_archive_service_maps_march_details(tmp_path: Path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    march_id = _save_march_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path,
    )
    details = service.get_details(march_id, user={"id": 1, "role": "admin"})

    assert details.product_type == "marches"
    assert len(details.marches) == 1
    assert details.marches[0].mark == "1ЛМ 27-11-14-4"
    assert details.plates == []


def test_archive_generate_pdf_for_saved_march_kp(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
) -> None:
    client, db_path = archive_client
    march_id = _save_march_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "outputs",
    )
    (tmp_path / "outputs").mkdir(exist_ok=True)

    path = asyncio.run(
        service.generate_document(march_id, "pdf", user={"id": 1, "role": "admin"})
    )

    assert path.exists()
    assert path.name == f"КП_{march_id}.pdf"
    assert path.stat().st_size > 100
