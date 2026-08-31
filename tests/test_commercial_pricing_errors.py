from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token
from app.services.draft_store import DraftStore
from core import commercial_offer, commercial_offer_xlsx
from core.commercial_pricing import ensure_order_priced, lookup_plate_price
from core.exceptions import PriceNotFoundError, UnpricedPlatesError


def _create_empty_price_db(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        cur = conn.cursor()
        cur.execute(
            "CREATE TABLE prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))"
        )
        conn.commit()
    finally:
        conn.close()


def _insert_price(path: Path, *, length_dm: int, load_code: int, price: float) -> None:
    conn = sqlite3.connect(path)
    try:
        conn.execute(
            "INSERT INTO prices (length_dm, load_code, price) VALUES (?, ?, ?)",
            (length_dm, load_code, price),
        )
        conn.commit()
    finally:
        conn.close()


def _unpriced_order_item() -> dict:
    return {
        "name": "ПБ 99-12-8п",
        "qty": 1,
        "length_m": 9.9,
        "width_m": 1.2,
        "load_class": 800,
    }


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    return CsrfAwareTestClient(create_app())


@pytest.fixture()
def auth_cookie(client: TestClient) -> dict[str, str]:
    token = create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    return {"app_session": token}


def test_lookup_plate_price_raises_when_db_price_is_zero(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_empty_price_db(db_path)
    _insert_price(db_path, length_dm=75, load_code=12, price=0.0)

    with pytest.raises(PriceNotFoundError):
        lookup_plate_price(7.5, 1.2, 1200, db_path=str(db_path))


def test_lookup_plate_price_returns_positive_price(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_empty_price_db(db_path)
    _insert_price(db_path, length_dm=70, load_code=12, price=29210.0)

    assert lookup_plate_price(7.0, 1.2, 1200, db_path=str(db_path)) == pytest.approx(29210.0)


def test_lookup_plate_price_raises_when_row_missing(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_empty_price_db(db_path)

    with pytest.raises(PriceNotFoundError):
        lookup_plate_price(9.9, 1.2, 800, db_path=str(db_path))


def test_ensure_order_priced_raises_with_position_labels(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    db_path = tmp_path / "pb.db"
    _create_empty_price_db(db_path)
    monkeypatch.setattr(commercial_offer, "DB_PATH", str(db_path))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(db_path))

    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced([_unpriced_order_item()], db_path=str(db_path))

    assert exc_info.value.positions == ["ПБ 99-12-8п"]


def test_generate_files_returns_unpriced_plates_and_skips_documents(
    client: TestClient,
    auth_cookie: dict[str, str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    db_path = tmp_path / "pb.db"
    _create_empty_price_db(db_path)
    outputs_dir = tmp_path / "outputs"
    drafts_dir = tmp_path / "drafts"
    outputs_dir.mkdir()
    drafts_dir.mkdir()

    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    monkeypatch.setenv("OUTPUTS_DIR", str(outputs_dir))
    get_settings.cache_clear()
    monkeypatch.setattr(commercial_offer, "DB_PATH", str(db_path))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(db_path))

    draft_id = "b" * 32
    store = DraftStore()
    order = PlateOrder()
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[_unpriced_order_item()],
        metadata={
            "owner_user_id": 1,
            "manager_id": 1,
            "manager_name": "Иван",
            "client_name": "ООО Тест",
            "conditions_mode": "standard",
            "wide_plates_resolved": True,
            "current_step": "result",
        },
    )

    before = {p.name for p in outputs_dir.iterdir()}

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/generate-files",
        json={"file_types": ["pdf", "xlsx"]},
    )

    after = {p.name for p in outputs_dir.iterdir()}

    assert response.status_code == 422
    body = response.json()
    assert body["detail"]["code"] == "unpriced_plates"
    assert body["detail"]["details"]["positions"] == ["ПБ 99-12-8п"]
    assert after == before
