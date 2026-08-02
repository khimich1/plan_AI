"""S3: KP membership — плита/свая рейса должна принадлежать КП из shipment_orders."""

from __future__ import annotations

import sqlite3

import pytest

from app.schemas.logistics import ShipmentItemInput, ShipmentOrderPatch
from app.services.shipment_service import ShipmentError, ShipmentService
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"
DATE = "2026-07-31"


@pytest.fixture
def db(tmp_path) -> str:
    return fx.make_iso_db(tmp_path)


@pytest.fixture
def svc(db: str) -> ShipmentService:
    return ShipmentService(db_path=db)


def _seed_completed(
    db_path: str,
    kp_id: int,
    qty: int,
    *,
    plate: str = PLATE,
    completed_date: str = "27.07.2026",
    day: int = 1,
) -> int:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 6.0, 1.2, 800, ?, ?, ?, 'plan-1')
            """,
            (kp_id, plate, qty, completed_date, day),
        )
        conn.commit()
        return int(cur.lastrowid)


def _create(svc: ShipmentService, kp_ids: list[int]) -> int:
    return svc.create(
        shipment_date=DATE, delivery_type="delivery", kp_ids=kp_ids, actor="tester"
    ).id


def _set_ya(svc: ShipmentService, shipment_id: int, kp_ids: list[int], ya: str) -> None:
    svc.patch(
        shipment_id,
        fields={},
        orders=[ShipmentOrderPatch(kp_id=kp_id, ya_order_no=ya) for kp_id in kp_ids],
    )


def _scalar(db_path: str, sql: str, params: tuple = ()):
    with sqlite3.connect(db_path) as conn:
        return conn.execute(sql, params).fetchone()[0]


def test_put_items_rejects_plate_from_foreign_kp(svc: ShipmentService, db: str) -> None:
    fx.seed_kp_offer(db, 1)
    fx.seed_kp_offer(db, 2)
    foreign_cp = _seed_completed(db, 2, 5)
    shipment_id = _create(svc, [1])

    with pytest.raises(ShipmentError) as exc:
        svc.put_items(
            shipment_id,
            [ShipmentItemInput(item_type="plate", completed_plate_id=foreign_cp, qty=1)],
        )
    assert exc.value.code == "shipment_plate_kp_mismatch"
    assert _scalar(db, "SELECT COUNT(*) FROM shipment_items WHERE shipment_id = ?", (shipment_id,)) == 0


def test_complete_rejects_sql_seeded_foreign_plate_without_qty_deduct(
    svc: ShipmentService, db: str
) -> None:
    """Обход put_items: чужой item уже в БД — complete должен отказать до списания."""
    fx.seed_kp_offer(db, 1)
    fx.seed_kp_offer(db, 2)
    own_cp = _seed_completed(db, 1, 5)
    foreign_cp = _seed_completed(db, 2, 5)
    shipment_id = _create(svc, [1])
    _set_ya(svc, shipment_id, [1], "ЯР-1")

    with sqlite3.connect(db) as conn:
        conn.execute(
            """
            INSERT INTO shipment_items (
                shipment_id, item_type, completed_plate_id, kp_id, mark,
                qty, unit_weight_kg, weight_kg, sort_order, note
            ) VALUES (?, 'plate', ?, 2, NULL, 1, 2040.0, 2040.0, 0, NULL)
            """,
            (shipment_id, foreign_cp),
        )
        conn.commit()

    with pytest.raises(ShipmentError) as exc:
        svc.complete(shipment_id)
    assert exc.value.code == "shipment_plate_kp_mismatch"

    assert _scalar(db, "SELECT qty FROM completed_plates WHERE id = ?", (own_cp,)) == 5
    assert _scalar(db, "SELECT qty FROM completed_plates WHERE id = ?", (foreign_cp,)) == 5
    assert _scalar(db, "SELECT status FROM shipments WHERE id = ?", (shipment_id,)) == "in_work"
    assert _scalar(db, "SELECT COUNT(*) FROM plate_status_log") == 0


def test_put_items_rejects_pile_with_foreign_kp(svc: ShipmentService, db: str) -> None:
    fx.seed_kp_offer(db, 1)
    fx.seed_kp_offer(db, 2)
    shipment_id = _create(svc, [1])

    with pytest.raises(ShipmentError) as exc:
        svc.put_items(
            shipment_id,
            [ShipmentItemInput(item_type="free", mark="С60.30", qty=1, kp_id=2)],
        )
    assert exc.value.code == "shipment_pile_kp_mismatch"


def test_put_items_allows_pile_without_kp(svc: ShipmentService, db: str) -> None:
    fx.seed_kp_offer(db, 1)
    shipment_id = _create(svc, [1])

    card = svc.put_items(
        shipment_id,
        [ShipmentItemInput(item_type="free", mark="С60.30", qty=2)],
    )
    assert len(card.items) == 1
    assert card.items[0].item_type == "free"
    assert card.items[0].kp_id is None
    assert card.items[0].qty == 2


def test_put_and_complete_own_plate_happy_path(svc: ShipmentService, db: str) -> None:
    fx.seed_kp_offer(db, 1)
    with sqlite3.connect(db) as conn:
        conn.execute("UPDATE kp_meta SET ordered_qty = 5 WHERE kp_id = 1")
        conn.commit()
    cp_id = _seed_completed(db, 1, 5)
    shipment_id = _create(svc, [1])
    _set_ya(svc, shipment_id, [1], "ЯР-1")

    card = svc.put_items(
        shipment_id,
        [ShipmentItemInput(item_type="plate", completed_plate_id=cp_id, qty=3)],
    )
    assert len(card.items) == 1
    assert card.items[0].kp_id == 1

    response = svc.complete(shipment_id, actor="logist")
    assert response.status == "done"
    assert _scalar(db, "SELECT qty FROM completed_plates WHERE id = ?", (cp_id,)) == 2
