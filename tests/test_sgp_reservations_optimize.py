"""SGP-401: reservations reduce optimizer demand before build."""

from __future__ import annotations

import sqlite3

import pytest

from app.services.sgp_service import SgpError, SgpService
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"


def test_reduce_selected_qty_demand_5_reserve_3(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    plate_id = fx.seed_plate(
        db,
        kp_id=1,
        plate_name=PLATE,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
        qty=5,
        status="в производстве",
    )
    with sqlite3.connect(db) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (NULL, ?, 6.0, 1.2, 800, 3, '27.07.2026')
            """,
            (PLATE,),
        )
        conn.commit()
        sgp_id = int(cur.lastrowid)

    svc = SgpService(db_path=db)
    reduced = svc.reduce_selected_qty_for_reservations(
        selected_plate_qty={1: {plate_id: 5}},
        sgp_reservations=[{"sgp_id": sgp_id, "target_kp_id": 1, "qty": 3}],
    )
    assert reduced[1][plate_id] == 2


def test_reduce_without_prior_qty_map(tmp_path) -> None:
    """filter=all: empty qty map still gets overrides for matched plates."""
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    plate_id = fx.seed_plate(
        db,
        kp_id=1,
        plate_name=PLATE,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
        qty=5,
        status="в производстве",
    )
    with sqlite3.connect(db) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (NULL, ?, 6.0, 1.2, 800, 3, '27.07.2026')
            """,
            (PLATE,),
        )
        conn.commit()
        sgp_id = int(cur.lastrowid)

    svc = SgpService(db_path=db)
    reduced = svc.reduce_selected_qty_for_reservations(
        selected_plate_qty=None,
        sgp_reservations=[{"sgp_id": sgp_id, "target_kp_id": 1, "qty": 3}],
    )
    assert reduced[1][plate_id] == 2


def test_reduce_rejects_no_demand(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    with sqlite3.connect(db) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (NULL, ?, 6.0, 1.2, 800, 2, '27.07.2026')
            """,
            (PLATE,),
        )
        conn.commit()
        sgp_id = int(cur.lastrowid)

    svc = SgpService(db_path=db)
    with pytest.raises(SgpError) as exc:
        svc.reduce_selected_qty_for_reservations(
            selected_plate_qty=None,
            sgp_reservations=[{"sgp_id": sgp_id, "target_kp_id": 1, "qty": 2}],
        )
    assert exc.value.code == "sgp_no_matching_demand"
