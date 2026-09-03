"""Two slices of GET /production/kp-candidates: wizard vs in-work queue."""

from __future__ import annotations

import sqlite3

import pytest

from app.repositories.kp_repository import KpRepository
from app.services.production_service import ProductionService
from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService

PLATE = "ПБ 60-12-8п"


@pytest.fixture()
def db_path(tmp_path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _plate_line(*, qty: int, line_id: str = "ln_1") -> dict:
    return {
        "line_id": line_id,
        "name": PLATE,
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": qty,
        "unit_price": 1000.0,
        "weight": 500.0,
    }


def _save_plate_kp(
    db_path: str,
    *,
    qty: int,
    customer: str,
    terms: str,
    date: str = "01.03.2026",
) -> int:
    return KpPersistenceService.save_kp_to_db(
        date,
        [_plate_line(qty=qty, line_id=f"ln_{customer}")],
        customer_name=customer,
        execution_terms=terms,
        status="в работе",
        db_path=db_path,
    )


def _set_all_plates_status(db_path: str, kp_id: int, status: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_plates SET status = ? WHERE kp_id = ?",
            (status, kp_id),
        )
        conn.commit()


def _split_qty_in_plan(db_path: str, kp_id: int, *, in_plan: int) -> None:
    """Keep ``in_plan`` шт «в плане», remainder «в производстве» (one source row)."""
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, qty, plate_name, length_m, width_m, load_class, position_number
            FROM kp_plates WHERE kp_id = ? ORDER BY id
            """,
            (kp_id,),
        )
        row = cur.fetchone()
        assert row is not None
        plate_id, qty, name, length_m, width_m, load_class, pos = row
        remaining = int(qty) - in_plan
        assert remaining > 0
        cur.execute(
            "UPDATE kp_plates SET qty = ?, status = 'в производстве' WHERE id = ?",
            (remaining, plate_id),
        )
        cur.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'в плане')
            """,
            (kp_id, pos, name, length_m, width_m, load_class, in_plan),
        )
        conn.commit()


def _move_all_to_sgp(db_path: str, kp_id: int) -> None:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT plate_name, length_m, width_m, load_class, SUM(qty) FROM kp_plates WHERE kp_id = ?",
            (kp_id,),
        )
        name, length_m, width_m, load_class, qty = cur.fetchone()
        cur.execute("DELETE FROM kp_plates WHERE kp_id = ?", (kp_id,))
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (?, ?, ?, ?, ?, ?, '01.09.2026')
            """,
            (kp_id, name, length_m, width_m, load_class, qty),
        )
        conn.commit()


def _service(db_path: str) -> ProductionService:
    return ProductionService(kp_repository=KpRepository(db_path=db_path))


def test_plan_scope_hides_fully_scheduled_kp(db_path: str) -> None:
    kp_id = _save_plate_kp(db_path, qty=10, customer="Full plan", terms="01.09.2026")
    _set_all_plates_status(db_path, kp_id, "в плане")

    plan = _service(db_path).list_kp_candidates()
    in_work = _service(db_path).list_kp_candidates(scope="in_work")

    assert plan["count"] == 0
    assert in_work["count"] == 1
    item = in_work["items"][0]
    assert item["kp_id"] == kp_id
    assert item["remaining_qty"] == 0
    assert item["in_plan_qty"] == 10
    assert item["on_sgp_qty"] == 0
    assert all(p["bucket"] == "in_plan" for p in item["plates"])
    assert sum(p["qty"] for p in item["plates"]) == 10


def test_in_work_excludes_fully_on_sgp(db_path: str) -> None:
    kp_id = _save_plate_kp(db_path, qty=3, customer="Done SGP", terms="01.09.2026")
    _move_all_to_sgp(db_path, kp_id)

    plan = _service(db_path).list_kp_candidates()
    in_work = _service(db_path).list_kp_candidates(scope="in_work")

    assert plan["count"] == 0
    assert in_work["count"] == 0


def test_in_work_qty_and_buckets_for_partial_plan(db_path: str) -> None:
    kp_id = _save_plate_kp(db_path, qty=10, customer="Partial", terms="15.09.2026")
    _split_qty_in_plan(db_path, kp_id, in_plan=6)

    plan = _service(db_path).list_kp_candidates()
    in_work = _service(db_path).list_kp_candidates(scope="in_work")

    assert plan["count"] == 1
    plan_item = plan["items"][0]
    assert plan_item["kp_id"] == kp_id
    assert plan_item["remaining_qty"] == 4
    assert plan_item["in_plan_qty"] == 6
    assert all(p["bucket"] == "awaiting_plan" for p in plan_item["plates"])
    assert sum(p["qty"] for p in plan_item["plates"]) == 4

    work_item = in_work["items"][0]
    assert work_item["remaining_qty"] == 4
    assert work_item["in_plan_qty"] == 6
    assert work_item["on_sgp_qty"] == 0
    buckets = {p["bucket"] for p in work_item["plates"]}
    assert buckets == {"awaiting_plan", "in_plan"}
    by_bucket = {p["bucket"]: p["qty"] for p in work_item["plates"]}
    assert by_bucket["awaiting_plan"] == 4
    assert by_bucket["in_plan"] == 6


def test_in_work_sorts_hot_deadlines_first(db_path: str) -> None:
    later = _save_plate_kp(db_path, qty=2, customer="Later", terms="01.10.2026")
    sooner = _save_plate_kp(db_path, qty=2, customer="Sooner", terms="01.09.2026")
    _set_all_plates_status(db_path, later, "в плане")
    _set_all_plates_status(db_path, sooner, "в плане")

    items = _service(db_path).list_kp_candidates(scope="in_work")["items"]
    assert [row["kp_id"] for row in items] == [sooner, later]


def test_in_work_unparseable_terms_go_last(db_path: str) -> None:
    dated = _save_plate_kp(db_path, qty=1, customer="Dated", terms="01.09.2026")
    junk = _save_plate_kp(db_path, qty=1, customer="Junk", terms="когда получится")
    _set_all_plates_status(db_path, dated, "в плане")
    _set_all_plates_status(db_path, junk, "в плане")

    items = _service(db_path).list_kp_candidates(scope="in_work")["items"]
    assert [row["kp_id"] for row in items] == [dated, junk]


def test_default_list_kp_candidates_is_plan_scope(db_path: str) -> None:
    kp_id = _save_plate_kp(db_path, qty=2, customer="Open", terms="01.09.2026")
    assert _service(db_path).list_kp_candidates()["items"][0]["kp_id"] == kp_id
    _set_all_plates_status(db_path, kp_id, "в плане")
    assert _service(db_path).list_kp_candidates()["count"] == 0
