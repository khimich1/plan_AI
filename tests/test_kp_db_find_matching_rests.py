"""Regression tests for ``find_matching_rests``."""

from __future__ import annotations

import sqlite3

import pytest

from core import kp_db
from tests.helpers import kp_db_fixtures as fx

LONG_CUT = 460.0
TRANSVERSE = 1200.0


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def _seed_rest(
    db: str,
    *,
    kp_id: int = 1,
    length_m: float = 6.0,
    width_mm: int = 1200,
    qty: int = 1,
    status: str = "available",
) -> int:
    fx.seed_kp_offer(db, kp_id)
    rest_id = kp_db.create_plate_rest(
        kp_id,
        "ПБ 60-12-8п",
        width_mm,
        length_m,
        production_day=1,
        qty=qty,
        db_path=db,
    )
    if status != "available":
        with sqlite3.connect(db) as conn:
            conn.execute(
                "UPDATE plate_rests SET status = ? WHERE id = ?",
                (status, rest_id),
            )
            conn.commit()
    return rest_id


def test_r1_exact_match_zero_cut_cost(iso_db: str) -> None:
    _seed_rest(iso_db, length_m=6.0, width_mm=1200)

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert len(matches) == 1
    assert matches[0]["match_type"] == "exact"
    assert matches[0]["cut_cost"] == 0.0


def test_r2_width_cut_cost(iso_db: str) -> None:
    _seed_rest(iso_db, length_m=6.0, width_mm=1500)

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert matches[0]["match_type"] == "width_cut"
    assert matches[0]["cut_cost"] == pytest.approx(LONG_CUT * 6.0)


def test_r3_length_cut_cost(iso_db: str) -> None:
    _seed_rest(iso_db, length_m=7.0, width_mm=1200)

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert matches[0]["match_type"] == "length_cut"
    assert matches[0]["cut_cost"] == pytest.approx(TRANSVERSE)


def test_r4_both_cuts_cost(iso_db: str) -> None:
    _seed_rest(iso_db, length_m=7.0, width_mm=1500)

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert matches[0]["match_type"] == "both_cuts"
    assert matches[0]["cut_cost"] == pytest.approx(LONG_CUT * 6.0 + TRANSVERSE)


def test_r5_exact_ordered_before_heavier_rest(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    kp_db.create_plate_rest(1, "A", 1200, 6.0, 1, qty=1, db_path=iso_db)
    kp_db.create_plate_rest(1, "B", 1500, 7.0, 1, qty=1, db_path=iso_db)

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert matches[0]["match_type"] == "exact"
    assert matches[0]["rest_length"] == pytest.approx(6.0)


def test_r6_qty_needed_from_multiple_rests(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    kp_db.create_plate_rest(1, "A", 1200, 6.0, 1, qty=1, db_path=iso_db)
    kp_db.create_plate_rest(1, "B", 1200, 6.0, 1, qty=2, db_path=iso_db)

    matches = kp_db.find_matching_rests(6.0, 1200, 3, db_path=iso_db)

    assert sum(m["qty_to_use"] for m in matches) == 3


def test_r7_non_available_excluded(iso_db: str) -> None:
    _seed_rest(iso_db, status="used")

    matches = kp_db.find_matching_rests(6.0, 1200, 1, db_path=iso_db)

    assert matches == []
