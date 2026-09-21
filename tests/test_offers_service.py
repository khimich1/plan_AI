"""Unit tests for OffersService PDF/XLSX generation and GPS-022 archive notify."""

from __future__ import annotations

import json
import sqlite3
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.core.settings import get_settings
from app.repositories.auth_repository import AuthRepository
from app.repositories.kp_repository import KpRepository
from app.repositories.promise_repository import PromiseRepository
from app.schemas.offers import CreateOfferRequest, OfferOrderItem
from app.services.kp_guid_notify import NOTIFICATION_KIND
from app.services.offers_service import OffersService
from core.nomenclature_guid import ensure_schema, upsert
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY


def _offer_without_creation_date() -> dict:
    return {
        "kp_id": 42,
        "creation_date": None,
        "customer_name": "Клиент",
        "manager_name": "Менеджер",
        "discount_percent": 0,
        "delivery_conditions": None,
        "payment_conditions": None,
        "plates": [
            {
                "plate_name": "ПБ 60-12-8п",
                "length_m": 6.0,
                "width_m": 1.2,
                "load_class": 800,
                "qty": 1,
                "unit_price": 1000.0,
                "unit_weight": 500.0,
                "total_weight": 500.0,
            }
        ],
    }


@pytest.fixture()
def offers_service() -> OffersService:
    repo = MagicMock()
    repo.get_offer.return_value = _offer_without_creation_date()
    return OffersService(kp_repository=repo)


def test_generate_pdf_without_creation_date_no_nameerror(
    offers_service: OffersService, monkeypatch: pytest.MonkeyPatch
) -> None:
    captured: dict = {}

    def fake_pdf(**kwargs):
        captured.update(kwargs)
        buf = BytesIO(b"%PDF")
        return buf

    monkeypatch.setattr(
        "app.services.offers_service.generate_commercial_offer_pdf", fake_pdf
    )
    user = {"id": 1, "role": "admin", "username": "admin"}
    filename, data = offers_service.generate_pdf(42, user=user)
    assert filename == "KP_42.pdf"
    assert data == b"%PDF"
    assert "offer_date" in captured
    assert isinstance(captured["offer_date"], str)
    assert len(captured["offer_date"]) > 0


def test_generate_xlsx_without_creation_date_no_nameerror(
    offers_service: OffersService, monkeypatch: pytest.MonkeyPatch
) -> None:
    captured: dict = {}

    def fake_xlsx(**kwargs):
        captured.update(kwargs)
        buf = BytesIO(b"PK")
        return buf

    monkeypatch.setattr(
        "app.services.offers_service.generate_commercial_offer_xlsx", fake_xlsx
    )
    user = {"id": 1, "role": "admin", "username": "admin"}
    filename, data = offers_service.generate_xlsx(42, user=user)
    assert filename == "KP_42.xlsx"
    assert data == b"PK"
    assert "offer_date" in captured
    assert isinstance(captured["offer_date"], str)
    assert len(captured["offer_date"]) > 0


_PWD = "StrongPassword123!"
_USER = {"id": 1, "role": "admin", "username": "admin"}


def _seed_client(db: str) -> int:
    from app.repositories.counterparties_repository import CounterpartiesRepository

    client = CounterpartiesRepository(db_path=db).insert(
        code_1c="00-client",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        is_client=True,
        source="import",
    )
    return int(client["id"])


def _payload(counterparty_id: int, *, save_mode: str = "archive") -> CreateOfferRequest:
    return CreateOfferRequest.model_validate(
        {
            "creation_date": "01.01.2026",
            "customer_name": "Имя из формы",
            "manager_name": "Иванов",
            "order_data": [
                OfferOrderItem.model_validate(
                    {
                        "name": "С70.35-9у",
                        "length_m": 7.0,
                        "width_m": 0.35,
                        "load_class": 800,
                        "qty": 2,
                        "unit_price": 1000.0,
                        "weight": 500.0,
                    }
                )
            ],
            "save_mode": save_mode,
            "counterparty_id": counterparty_id,
        }
    )


def _payload_without_id(*, save_mode: str = "archive") -> CreateOfferRequest:
    return CreateOfferRequest.model_validate(
        {
            "creation_date": "01.01.2026",
            "customer_name": "Имя из формы",
            "manager_name": "Иванов",
            "order_data": [
                OfferOrderItem.model_validate(
                    {
                        "name": "С70.35-9у",
                        "length_m": 7.0,
                        "width_m": 0.35,
                        "load_class": 800,
                        "qty": 2,
                        "unit_price": 1000.0,
                        "weight": 500.0,
                    }
                )
            ],
            "save_mode": save_mode,
        }
    )


def _patch_dbs(monkeypatch: pytest.MonkeyPatch, plita: str, pb: Path) -> None:
    import core.db_config as db_config

    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PB_DB_PATH", str(pb))
    monkeypatch.setenv("PLITA_DB_PATH", plita)
    get_settings.cache_clear()
    # PB_DB_PATH is bound at import time in core.db_config; env + cache_clear
    # do not update it. kp_persistence_service imports it at call time.
    monkeypatch.setattr(db_config, "PB_DB_PATH", str(pb))


def _init_pb_missing_pile(pb: Path) -> None:
    conn = sqlite3.connect(pb)
    try:
        ensure_schema(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        upsert(conn, "pile", "С70.35-9", match_status="missing")
        conn.commit()
    finally:
        conn.close()


def test_create_offer_archive_notifies_economist_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb_missing_pile(pb)
    _patch_dbs(monkeypatch, plita, pb)
    client_id = _seed_client(plita)
    econ, _created = AuthRepository(db_path=plita).create_or_update_user(
        username="econ",
        password=_PWD,
        role="economist",
    )
    service = OffersService(kp_repository=KpRepository(db_path=plita))

    created = service.create_offer(_payload(client_id), user=_USER)
    kp_id = int(created["kp_id"])
    assert created["status"] == "в архиве"

    repo = PromiseRepository(db_path=plita)
    rows = repo.list_notifications(user_id=int(econ["id"]), kind=NOTIFICATION_KIND)
    assert len(rows) == 1
    payload = json.loads(rows[0]["payload_json"])
    assert payload["kp_id"] == kp_id
    assert payload["seq"] == kp_id
    assert "С70.35-9у" in payload["marks"]
    assert payload["link"] == "/prices"

    get_settings.cache_clear()


def test_create_offer_archive_without_id_writes_form_name(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb_missing_pile(pb)
    _patch_dbs(monkeypatch, plita, pb)
    service = OffersService(kp_repository=KpRepository(db_path=plita))
    created = service.create_offer(_payload_without_id(), user=_USER)
    assert created["status"] == "в архиве"
    with sqlite3.connect(plita) as conn:
        row = conn.execute(
            "SELECT customer_name, counterparty_id FROM KP_offers WHERE kp_id = ?",
            (int(created["kp_id"]),),
        ).fetchone()
    assert row[0] == "Имя из формы"
    assert row[1] is None
    get_settings.cache_clear()


def test_create_offer_work_mode_does_not_notify(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb_missing_pile(pb)
    _patch_dbs(monkeypatch, plita, pb)
    client_id = _seed_client(plita)
    econ, _created = AuthRepository(db_path=plita).create_or_update_user(
        username="econ",
        password=_PWD,
        role="economist",
    )
    service = OffersService(kp_repository=KpRepository(db_path=plita))
    created = service.create_offer(_payload(client_id, save_mode="work"), user=_USER)
    assert created["status"] == "в работе"
    repo = PromiseRepository(db_path=plita)
    assert repo.list_notifications(user_id=int(econ["id"])) == []
    get_settings.cache_clear()

