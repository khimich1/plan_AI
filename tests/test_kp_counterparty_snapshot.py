"""CTR-005: снапшот контрагента в KP_offers."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from app.repositories.counterparties_repository import CounterpartiesRepository
from app.repositories.kp_repository import KpRepository
from app.services.offers_service import OffersService
from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx

_ORDER = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 2,
        "unit_price": 1000.0,
        "weight": 5000.0,
        "length_dm_raw": "60",
        "concrete_grade": "М500",
    }
]


def _select_snapshot(db: str, kp_id: int) -> tuple:
    with sqlite3.connect(db) as conn:
        return conn.execute(
            "SELECT counterparty_id, customer_inn, customer_kpp, customer_name "
            "FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()


def test_save_kp_writes_counterparty_snapshot(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    repo = CounterpartiesRepository(db_path=db)
    created = repo.insert(
        code_1c="00-1",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        source="import",
    )
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name="РОМАШКА ООО",
        db_path=db,
        counterparty_id=int(created["id"]),
        customer_inn="7701000001",
        customer_kpp="770101001",
    )
    row = _select_snapshot(db, kp_id)
    assert row == (created["id"], "7701000001", "770101001", "РОМАШКА ООО")


def test_save_kp_without_counterparty_keeps_nulls(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name="Старый клиент",
        db_path=db,
    )
    row = _select_snapshot(db, kp_id)
    assert row == (None, None, None, "Старый клиент")


def test_kp_repository_save_offer_proxies_snapshot(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    created = CounterpartiesRepository(db_path=db).insert(
        code_1c="00-1",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        source="import",
    )
    kp_id = KpRepository(db_path=db).save_offer(
        customer_name="РОМАШКА ООО",
        manager_name="Иванов",
        creation_date="01.01.2026",
        order_data=_ORDER,
        counterparty_id=int(created["id"]),
        customer_inn="7701000001",
        customer_kpp="770101001",
    )
    row = _select_snapshot(db, kp_id)
    assert row[0] == created["id"]
    assert row[1] == "7701000001"
    assert row[2] == "770101001"


def test_offer_read_models_include_snapshot_fields(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    created = CounterpartiesRepository(db_path=db).insert(
        code_1c="00-1",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        source="import",
    )
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name="РОМАШКА ООО",
        db_path=db,
        counterparty_id=int(created["id"]),
        customer_inn="7701000001",
        customer_kpp="770101001",
    )
    service = OffersService(kp_repository=KpRepository(db_path=db))
    user = {"id": 1, "role": "admin"}
    details = service.get_offer(kp_id, user=user)
    assert details is not None
    assert details["counterparty_id"] == created["id"]
    assert details["customer_inn"] == "7701000001"
    assert details["customer_kpp"] == "770101001"
    summary = service._to_offer_summary(details)
    assert summary["counterparty_id"] == created["id"]
    assert summary["customer_inn"] == "7701000001"
