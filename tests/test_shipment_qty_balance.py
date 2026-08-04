"""SHIP-203/204: инвариант количества плит.

Σ kp_plates + Σ completed_plates + Σ отгруженного (plate-строки done-рейсов)
остаётся константой при complete/cancel/мульти-КП сценариях.
"""

from __future__ import annotations

import sqlite3

import pytest

from app.schemas.logistics import ShipmentItemInput, ShipmentOrderPatch
from app.services.shipment_service import ShipmentError, ShipmentService
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"


@pytest.fixture
def db(tmp_path) -> str:
    return fx.make_iso_db(tmp_path)


@pytest.fixture
def svc(db: str) -> ShipmentService:
    return ShipmentService(db_path=db)


def _physics_total(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(SUM(qty), 0) FROM kp_plates")
        in_kp = int(cur.fetchone()[0])
        cur.execute("SELECT COALESCE(SUM(qty), 0) FROM completed_plates")
        on_sgp = int(cur.fetchone()[0])
        cur.execute(
            """
            SELECT COALESCE(SUM(si.qty), 0)
            FROM shipment_items si
            JOIN shipments s ON s.id = si.shipment_id
            WHERE s.status = 'done' AND si.item_type = 'plate'
            """
        )
        shipped = int(cur.fetchone()[0])
    return in_kp + on_sgp + shipped


def _seed_kp_with_sgp(db_path: str, kp_id: int, qty: int) -> int:
    fx.seed_kp_offer(db_path, kp_id)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ?", (qty, kp_id)
        )
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 6.0, 1.2, 800, ?, '27.07.2026', 1, 'plan-1')
            """,
            (kp_id, PLATE, qty),
        )
        conn.commit()
        return int(cur.lastrowid)


def _create(svc: ShipmentService, kp_ids: list[int]) -> int:
    return svc.create(
        shipment_date="2026-07-31", delivery_type="delivery", kp_ids=kp_ids
    ).id


def _confirm_and_complete(
    svc: ShipmentService,
    shipment_id: int,
    items: list[tuple[int, int]],
    kp_ids: list[int],
) -> None:
    svc.put_items(
        shipment_id,
        [
            ShipmentItemInput(item_type="plate", completed_plate_id=cp_id, qty=qty)
            for cp_id, qty in items
        ],
    )
    svc.patch(
        shipment_id,
        fields={},
        orders=[ShipmentOrderPatch(kp_id=kp_id, ya_order_no="ЯР-1") for kp_id in kp_ids],
    )
    svc.complete(shipment_id)


def test_balance_across_complete(svc: ShipmentService, db: str) -> None:
    cp_id = _seed_kp_with_sgp(db, 1, 5)
    baseline = _physics_total(db)
    assert baseline == 5

    shipment_id = _create(svc, [1])
    # Подтверждение состава (резерв) не меняет физику.
    svc.put_items(
        shipment_id,
        [ShipmentItemInput(item_type="plate", completed_plate_id=cp_id, qty=3)],
    )
    assert _physics_total(db) == baseline

    svc.patch(
        shipment_id,
        fields={},
        orders=[ShipmentOrderPatch(kp_id=1, ya_order_no="ЯР-1")],
    )
    svc.complete(shipment_id)
    after = _physics_total(db)
    assert after == baseline  # 2 на складе + 3 отгружено

    with sqlite3.connect(db) as conn:
        on_sgp = conn.execute(
            "SELECT qty FROM completed_plates WHERE id = ?", (cp_id,)
        ).fetchone()[0]
    assert on_sgp == 2


def test_balance_across_cancel(svc: ShipmentService, db: str) -> None:
    cp_id = _seed_kp_with_sgp(db, 1, 5)
    baseline = _physics_total(db)

    shipment_id = _create(svc, [1])
    svc.put_items(
        shipment_id,
        [ShipmentItemInput(item_type="plate", completed_plate_id=cp_id, qty=4)],
    )
    svc.cancel(shipment_id)
    assert _physics_total(db) == baseline

    # После отмены резерв свободен: можно отгрузить всё.
    second_id = _create(svc, [1])
    _confirm_and_complete(svc, second_id, [(cp_id, 5)], [1])
    assert _physics_total(db) == baseline


def test_balance_multi_kp_and_failed_complete(svc: ShipmentService, db: str) -> None:
    cp1 = _seed_kp_with_sgp(db, 1, 5)
    cp2 = _seed_kp_with_sgp(db, 2, 4)
    baseline = _physics_total(db)
    assert baseline == 9

    multi_id = _create(svc, [1, 2])
    svc.put_items(
        multi_id,
        [
            ShipmentItemInput(item_type="plate", completed_plate_id=cp1, qty=5),
            ShipmentItemInput(item_type="plate", completed_plate_id=cp2, qty=3),
        ],
    )
    assert _physics_total(db) == baseline

    # Неудачный complete (нет ЯР) — откат, физика не меняется.
    with pytest.raises(ShipmentError):
        svc.complete(multi_id)
    assert _physics_total(db) == baseline

    svc.patch(
        multi_id,
        fields={},
        orders=[
            ShipmentOrderPatch(kp_id=1, ya_order_no="ЯР-1"),
            ShipmentOrderPatch(kp_id=2, ya_order_no="ЯР-2"),
        ],
    )
    svc.complete(multi_id)
    assert _physics_total(db) == baseline

    # Довоз остатка КП 2 вторым рейсом.
    tail_id = _create(svc, [2])
    _confirm_and_complete(svc, tail_id, [(cp2, 1)], [2])
    assert _physics_total(db) == baseline

    with sqlite3.connect(db) as conn:
        shipped = conn.execute(
            """
            SELECT COALESCE(SUM(si.qty), 0)
            FROM shipment_items si
            JOIN shipments s ON s.id = si.shipment_id
            WHERE s.status = 'done' AND si.item_type = 'plate'
            """
        ).fetchone()[0]
        left = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates"
        ).fetchone()[0]
    assert shipped == 9
    assert left == 0
