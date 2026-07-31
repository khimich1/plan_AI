"""Qty balance invariants for SGP operations (plate_loss style)."""

from __future__ import annotations

import sqlite3

from app.services.sgp_service import SgpService
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"


def _totals(db_path: str, kp_id: int) -> tuple[int, int, int]:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM kp_plates WHERE kp_id=?",
            (kp_id,),
        )
        demand = int(cur.fetchone()[0])
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id=?",
            (kp_id,),
        )
        linked = int(cur.fetchone()[0])
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id IS NULL"
        )
        free = int(cur.fetchone()[0])
    return demand, linked, free


def test_send_unlink_relink_preserves_total_physics(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    fx.seed_kp_offer(db, 2)
    with sqlite3.connect(db) as conn:
        conn.execute("UPDATE kp_meta SET ordered_qty=10 WHERE kp_id=1")
        conn.execute("UPDATE kp_meta SET ordered_qty=10 WHERE kp_id=2")
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day
            ) VALUES (1, ?, 6.0, 1.2, 800, 10, '27.07.2026', 1)
            """,
            (PLATE,),
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (2, 1, ?, 6.0, 1.2, 800, 4, 'в производстве')
            """,
            (PLATE,),
        )
        conn.commit()

    # Initial: demand1=0 linked1=10 free=0; demand2=4
    d1, l1, free = _totals(db, 1)
    assert d1 + l1 + free == 10

    svc = SgpService(db_path=db)
    sgp_id = svc.list_plates(filter="linked").items[0].id
    svc.unlink(sgp_id, 3)
    d1, l1, free = _totals(db, 1)
    assert d1 == 3 and l1 == 7 and free == 3
    assert d1 + l1 == 10

    with sqlite3.connect(db) as conn:
        total_physics = conn.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates"
        ).fetchone()[0]
    assert int(total_physics) == 10  # physics never shrinks on unlink

    free_id = svc.list_plates(filter="unlinked").items[0].id
    svc.relink(free_id, target_kp_id=2, qty=2)

    with sqlite3.connect(db) as conn:
        total_physics = int(
            conn.execute(
                "SELECT COALESCE(SUM(qty),0) FROM completed_plates"
            ).fetchone()[0]
        )
        demand2 = int(
            conn.execute(
                "SELECT COALESCE(SUM(qty),0) FROM kp_plates WHERE kp_id=2"
            ).fetchone()[0]
        )
        linked2 = int(
            conn.execute(
                "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id=2"
            ).fetchone()[0]
        )
    assert total_physics == 10
    assert demand2 == 2
    assert linked2 == 2
