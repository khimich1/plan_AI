"""ProductionUrgentService (orch-2026-08-12-podlozhki Task 6)."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.domain.enums import PlateStatus
from app.services import production_urgent_service as urgent_svc_mod
from app.services.production_urgent_service import (
    ProductionUrgentService,
    UrgentPosition,
)
from core.production.urgent import UrgentPosition as CoreUrgentPosition
from tests.helpers import kp_db_fixtures as fx

FIXED_NOW = datetime(2026, 8, 1, 12, 0, 0)
_TS = "2026-08-01T12:00:00"


def _service(tmp_path: Path) -> ProductionUrgentService:
    return ProductionUrgentService(db_path=fx.make_iso_db(tmp_path))


def _set_execution_terms(db_path: str, kp_id: int, terms: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE KP_offers SET execution_terms = ? WHERE kp_id = ?",
            (terms, kp_id),
        )
        conn.commit()


def _seed_delivery_batch(
    db_path: str,
    *,
    kp_id: int,
    plate_id: int,
    produce_by: str,
    qty: int,
    batch_name: str = "П1",
) -> None:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO delivery_schedule (
                kp_id, invoice_number, contract_number, status, created_at, updated_at
            ) VALUES (?, 'СЧ-1', 'Д-1', 'draft', ?, ?)
            """,
            (kp_id, _TS, _TS),
        )
        schedule_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO delivery_batch (
                schedule_id, name, deliver_from, deliver_to, produce_by, sort_order
            ) VALUES (?, ?, '2026-09-01', '2026-09-10', ?, 1)
            """,
            (schedule_id, batch_name, produce_by),
        )
        batch_id = int(cur.lastrowid)
        cur.execute(
            "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) VALUES (?, ?, ?)",
            (batch_id, plate_id, qty),
        )
        conn.commit()


def test_urgent_position_reexported_from_core() -> None:
    assert UrgentPosition is CoreUrgentPosition


def test_batch_driven_urgent_included(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 1)
    _set_execution_terms(svc.db_path, 1, "31.12.2026")
    plate_id = fx.seed_plate(
        svc.db_path,
        kp_id=1,
        plate_name="ПБ 60",
        length_m=6.0,
        width_m=1.2,
        qty=5,
        status=PlateStatus.IN_PRODUCTION.value,
    )
    _seed_delivery_batch(
        svc.db_path,
        kp_id=1,
        plate_id=plate_id,
        produce_by="2026-08-10",
        qty=2,
        batch_name="1 этаж",
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )

    assert len(result) == 1
    pos = result[0]
    assert isinstance(pos, UrgentPosition)
    assert pos.plate_id == plate_id
    assert pos.kp_id == 1
    assert pos.deadline == date(2026, 8, 10)
    assert pos.deadline_source == "delivery_batch"
    assert pos.qty_remaining == 5
    assert any(d.get("batch_name") == "1 этаж" for d in pos.deadline_details)


def test_terms_only_urgent_included(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 2)
    _set_execution_terms(svc.db_path, 2, "10.08.2026")
    plate_id = fx.seed_plate(
        svc.db_path,
        kp_id=2,
        plate_name="ПБ 50",
        length_m=5.0,
        width_m=1.2,
        qty=3,
        status=PlateStatus.IN_PRODUCTION.value,
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 15),
        now=FIXED_NOW,
    )

    assert len(result) == 1
    pos = result[0]
    assert pos.plate_id == plate_id
    assert pos.deadline == date(2026, 8, 10)
    assert pos.deadline_source == "execution_terms"
    assert pos.qty_remaining == 3


def test_planned_plate_excluded(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 3)
    _set_execution_terms(svc.db_path, 3, "05.08.2026")
    fx.seed_plate(
        svc.db_path,
        kp_id=3,
        plate_name="ПБ 40",
        length_m=4.0,
        width_m=1.2,
        qty=4,
        status=PlateStatus.IN_PLAN.value,
        plan_id="plan-1",
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert result == []


def test_qty_remaining_reflects_planned_split(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 4)
    _set_execution_terms(svc.db_path, 4, "08.08.2026")
    planned_id = fx.seed_plate(
        svc.db_path,
        kp_id=4,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=2,
        status=PlateStatus.IN_PLAN.value,
        plan_id="plan-split",
        position_number=1,
    )
    remainder_id = fx.seed_plate(
        svc.db_path,
        kp_id=4,
        plate_name="ПК 60.15",
        length_m=6.0,
        width_m=1.5,
        qty=3,
        status=PlateStatus.IN_PRODUCTION.value,
        position_number=2,
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )

    assert len(result) == 1
    pos = result[0]
    assert pos.plate_id == remainder_id
    assert pos.qty_remaining == 3
    assert planned_id not in {p.plate_id for p in result}


def test_empty_backlog_returns_empty(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert result == []


def test_terms_beyond_deadline_excluded(tmp_path: Path) -> None:
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 5)
    _set_execution_terms(svc.db_path, 5, "30.09.2026")
    fx.seed_plate(
        svc.db_path,
        kp_id=5,
        plate_name="ПБ 55",
        length_m=5.5,
        width_m=1.2,
        qty=1,
        status=PlateStatus.IN_PRODUCTION.value,
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )
    assert result == []


def test_qty_remaining_wired_from_repository(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """AC: qty_remaining must come from KpRepository.get_plate_qty_remaining."""
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 6)
    _set_execution_terms(svc.db_path, 6, "12.08.2026")
    plate_id = fx.seed_plate(
        svc.db_path,
        kp_id=6,
        plate_name="ПБ 62",
        length_m=6.2,
        width_m=1.2,
        qty=4,
        status=PlateStatus.IN_PRODUCTION.value,
    )

    called: list[int] = []

    def fake_qty_remaining(pid: int) -> int:
        called.append(pid)
        return 99

    monkeypatch.setattr(
        svc.kp_repository, "get_plate_qty_remaining", fake_qty_remaining
    )

    result = svc.list_urgent_positions(
        deadline_until=date(2026, 8, 20),
        now=FIXED_NOW,
    )

    assert called == [plate_id]
    assert len(result) == 1
    assert result[0].qty_remaining == 99


def test_delegates_aggregation_to_core_collect(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """AC: service must aggregate via core.production.urgent.collect_urgent_positions."""
    svc = _service(tmp_path)
    fx.seed_kp_offer(svc.db_path, 7)
    _set_execution_terms(svc.db_path, 7, "11.08.2026")
    plate_id = fx.seed_plate(
        svc.db_path,
        kp_id=7,
        plate_name="ПБ 70",
        length_m=7.0,
        width_m=1.2,
        qty=2,
        status=PlateStatus.IN_PRODUCTION.value,
    )
    deadline_until = date(2026, 8, 20)
    sentinel = MagicMock(name="urgent_from_core")
    collect = MagicMock(return_value=[sentinel])
    monkeypatch.setattr(urgent_svc_mod, "collect_urgent_positions", collect)

    result = svc.list_urgent_positions(
        deadline_until=deadline_until,
        now=FIXED_NOW,
    )

    collect.assert_called_once()
    args, kwargs = collect.call_args
    plates, batches_by_plate, kp_meta, until = args
    assert until == deadline_until
    assert kwargs.get("now") == FIXED_NOW
    assert any(int(p["plate_id"]) == plate_id for p in plates)
    assert 7 in kp_meta
    assert isinstance(batches_by_plate, dict)
    assert result == [sentinel]
