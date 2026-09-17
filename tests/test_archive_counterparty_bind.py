"""CBP-004: PATCH привязки контрагента 1С к КП в архиве."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.repositories.counterparties_repository import CounterpartiesRepository
from app.repositories.kp_archive_repository import KpArchiveRepository
from app.services.archive_service import (
    ArchiveService,
    ArchiveValidationError,
)
from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx

ADMIN = {"id": 1, "role": "admin"}

_ORDER = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "length_dm_raw": "60",
        "concrete_grade": "М500",
    }
]


def _seed_clients(db: str) -> dict[str, dict]:
    repo = CounterpartiesRepository(db_path=db)
    client = repo.insert(
        code_1c="00-client",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        is_client=True,
        source="import",
    )
    other = repo.insert(
        code_1c="00-other",
        name="ЛЮТИК ООО",
        inn="7702000002",
        kpp="770202002",
        is_client=True,
        source="import",
    )
    supplier = repo.insert(
        code_1c="00-sup",
        name="ПОСТАВЩИК ООО",
        inn="9900000000",
        is_client=False,
        source="import",
    )
    inactive = repo.insert(
        code_1c="00-off",
        name="АРХИВ ООО",
        inn="8800000000",
        is_client=True,
        is_active=False,
        source="import",
    )
    return {"client": client, "other": other, "supplier": supplier, "inactive": inactive}


def _save_archived(db: str, *, customer_name: str = "ИП Петров") -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name=customer_name,
        status="в архиве",
        owner_user_id=1,
        db_path=db,
    )


def _snapshot(db: str, kp_id: int) -> tuple:
    with sqlite3.connect(db) as conn:
        return conn.execute(
            "SELECT customer_name, counterparty_id, customer_inn, customer_kpp, "
            "(SELECT status FROM kp_meta WHERE kp_id = KP_offers.kp_id) "
            "FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()


def _service(db: str, tmp_path: Path) -> ArchiveService:
    return ArchiveService(
        repository=KpArchiveRepository(db_path=db),
        outputs_dir=tmp_path / "out",
    )


def test_bind_active_client_overwrites_snapshot(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    kp_id = _save_archived(db)
    service = _service(db, tmp_path)

    details = service.bind_counterparty(
        kp_id, int(seeded["client"]["id"]), user=ADMIN
    )

    assert details.counterparty_id == seeded["client"]["id"]
    assert details.customer_name == "РОМАШКА ООО"
    assert details.customer_inn == "7701000001"
    assert details.customer_kpp == "770101001"
    row = _snapshot(db, kp_id)
    assert row == ("РОМАШКА ООО", seeded["client"]["id"], "7701000001", "770101001", "в архиве")


def test_bind_rejects_supplier_and_inactive(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    kp_id = _save_archived(db)
    service = _service(db, tmp_path)

    with pytest.raises(ArchiveValidationError, match="не отмечен как клиент"):
        service.bind_counterparty(kp_id, int(seeded["supplier"]["id"]), user=ADMIN)
    with pytest.raises(ArchiveValidationError, match="не найден"):
        service.bind_counterparty(kp_id, int(seeded["inactive"]["id"]), user=ADMIN)
    assert _snapshot(db, kp_id)[1] is None


def test_bind_rejects_when_not_archived(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name="ИП Петров",
        status="в работе",
        owner_user_id=1,
        db_path=db,
        counterparty_id=int(seeded["client"]["id"]),
        customer_inn="7701000001",
        customer_kpp="770101001",
    )
    service = _service(db, tmp_path)
    with pytest.raises(ArchiveValidationError, match="в архиве"):
        service.bind_counterparty(kp_id, int(seeded["other"]["id"]), user=ADMIN)
    assert _snapshot(db, kp_id)[1] == seeded["client"]["id"]


def test_rebind_in_archive_replaces_card(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        _ORDER,
        customer_name="РОМАШКА ООО",
        status="в архиве",
        owner_user_id=1,
        db_path=db,
        counterparty_id=int(seeded["client"]["id"]),
        customer_inn="7701000001",
        customer_kpp="770101001",
    )
    service = _service(db, tmp_path)
    details = service.bind_counterparty(kp_id, int(seeded["other"]["id"]), user=ADMIN)
    assert details.counterparty_id == seeded["other"]["id"]
    assert details.customer_name == "ЛЮТИК ООО"
    assert details.customer_inn == "7702000002"


def test_list_and_details_expose_counterparty_id(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    kp_id = _save_archived(db)
    service = _service(db, tmp_path)
    items = service.list_offers("archived", user=ADMIN)
    match = next(item for item in items if item.kp_id == kp_id)
    assert match.counterparty_id is None
    details = service.get_details(kp_id, user=ADMIN)
    assert details.counterparty_id is None

    seeded = _seed_clients(db)
    bound = service.bind_counterparty(kp_id, int(seeded["client"]["id"]), user=ADMIN)
    assert bound.counterparty_id == seeded["client"]["id"]
    items = service.list_offers("archived", user=ADMIN)
    match = next(item for item in items if item.kp_id == kp_id)
    assert match.counterparty_id == seeded["client"]["id"]


def test_bind_then_move_to_production(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    kp_id = _save_archived(db)
    service = _service(db, tmp_path)
    with pytest.raises(ArchiveValidationError, match="не найден"):
        service.move_to_production(kp_id, "15.10.2026", user=ADMIN)
    assert _snapshot(db, kp_id)[4] == "в архиве"

    service.bind_counterparty(kp_id, int(seeded["client"]["id"]), user=ADMIN)
    details = service.move_to_production(kp_id, "15.10.2026", user=ADMIN)
    assert details.status == "в работе"
    assert _snapshot(db, kp_id)[4] == "в работе"
