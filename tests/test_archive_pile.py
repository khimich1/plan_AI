"""Archive API: product_type filter and pile KP details (AC-10, AC-14, AC-15)."""

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

PLATE_ORDER = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
    }
]

PILE_ORDER = [
    {
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 44634.03,
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


def _save_plate_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        PLATE_ORDER,
        customer_name="Plate client",
        status="в архиве",
        db_path=db_path,
    )


def _save_pile_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        PILE_ORDER,
        customer_name="Pile client",
        status="в архиве",
        product_type="piles",
        db_path=db_path,
    )


def test_archive_list_includes_product_type(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    _save_plate_kp(db_path)
    _save_pile_kp(db_path)

    response = client.get(
        "/api/v1/commercial/archive?section=archived",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    by_type = {item["product_type"] for item in response.json()}
    assert by_type == {"plates", "piles"}


def test_archive_list_filter_product_type(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    plate_id = _save_plate_kp(db_path)
    pile_id = _save_pile_kp(db_path)

    piles_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=piles",
        cookies=_admin_cookie(),
    )
    assert piles_resp.status_code == 200
    pile_items = piles_resp.json()
    assert len(pile_items) == 1
    assert pile_items[0]["kp_id"] == pile_id
    assert pile_items[0]["product_type"] == "piles"

    plates_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=plates",
        cookies=_admin_cookie(),
    )
    assert plates_resp.status_code == 200
    plate_items = plates_resp.json()
    assert len(plate_items) == 1
    assert plate_items[0]["kp_id"] == plate_id
    assert plate_items[0]["product_type"] == "plates"

    all_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=all",
        cookies=_admin_cookie(),
    )
    assert all_resp.status_code == 200
    assert len(all_resp.json()) == 2


def test_archive_detail_includes_piles(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    response = client.get(
        f"/api/v1/commercial/archive/{pile_id}",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["product_type"] == "piles"
    assert payload["plates"] == []
    assert len(payload["piles"]) == 1
    pile = payload["piles"][0]
    assert pile["mark"] == "С120.35-12"
    assert pile["concrete_grade"] == "B25"
    assert pile["qty"] == 3
    assert pile["unit_price"] == pytest.approx(44634.03)


def test_archive_service_maps_pile_details(tmp_path: Path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    pile_id = _save_pile_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path,
    )
    details = service.get_details(pile_id, user={"id": 1, "role": "admin"})

    assert details.product_type == "piles"
    assert len(details.piles) == 1
    assert details.piles[0].mark == "С120.35-12"
    assert details.plates == []


def test_archive_generate_pdf_for_saved_pile_kp(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "outputs",
    )
    (tmp_path / "outputs").mkdir(exist_ok=True)

    path = asyncio.run(
        service.generate_document(pile_id, "pdf", user={"id": 1, "role": "admin"})
    )

    assert path.exists()
    assert path.name == f"КП_{pile_id}.pdf"
    assert path.stat().st_size > 100


def test_archive_download_pdf_http_for_pile_kp(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    outputs_dir = tmp_path / "archive_outputs"
    outputs_dir.mkdir()
    monkeypatch.setenv("OUTPUTS_DIR", str(outputs_dir))
    get_settings.cache_clear()

    response = client.get(
        f"/api/v1/commercial/archive/{pile_id}/files/pdf",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200, response.text
    assert response.headers["content-type"] == "application/pdf"
    assert len(response.content) > 100
