"""Per-step tests for find_kp_plate_row matching strategies."""

from __future__ import annotations

import sqlite3

import pytest

from core.domain.plate_completion_matching import (
    LENGTH_TOLERANCE_M,
    _match_step_exact_name,
    _match_step_length_dm_raw,
    find_kp_plate_row,
)
from core.domain.plate_completion_matching import PlateMatchContext, _width_parts
from tests.helpers import kp_db_fixtures as fx


@pytest.fixture
def match_db(tmp_path):
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    return db


def _ctx(
    cur: sqlite3.Cursor,
    *,
    plate_name: str = "ПБ 60-12-8п",
    length_m: float = 6.0,
    width_m: float = 1.2,
    load_class: int = 800,
    prefer_kp_id: int = 1,
    length_dm_raw: str | None = None,
    allow_cross_kp: bool = False,
    plan_ids: list[str] | None = None,
) -> PlateMatchContext:
    width_clause, width_params = _width_parts(width_m)
    return PlateMatchContext(
        cur=cur,
        plate_name=plate_name,
        length_m=length_m,
        width_m=width_m,
        load_class=load_class,
        prefer_kp_id=prefer_kp_id,
        length_dm_raw=length_dm_raw,
        allow_cross_kp=allow_cross_kp,
        plan_ids=plan_ids,
        width_clause=width_clause,
        width_params=width_params,
    )


def test_step_length_dm_raw(match_db: str) -> None:
    fx.seed_plate(
        match_db,
        kp_id=1,
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в производстве",
        length_dm_raw="60,0",
    )
    with sqlite3.connect(match_db) as conn:
        cur = conn.cursor()
        row = _match_step_length_dm_raw(
            _ctx(cur, length_dm_raw="60,0"),
        )
    assert row is not None
    assert row[1] == 1


def test_step_exact_name(match_db: str) -> None:
    fx.seed_plate(
        match_db,
        kp_id=1,
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_m=1.2,
        qty=2,
        status="в плане",
    )
    with sqlite3.connect(match_db) as conn:
        cur = conn.cursor()
        row = _match_step_exact_name(_ctx(cur))
    assert row is not None
    assert row[2] == "ПБ 60-12-8п"


def test_find_kp_plate_row_integration(match_db: str) -> None:
    fx.seed_plate(
        match_db,
        kp_id=1,
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в производстве",
    )
    with sqlite3.connect(match_db) as conn:
        cur = conn.cursor()
        row = find_kp_plate_row(cur, "ПБ 60-12-8п", 6.0, 1.2, 800, 1)
    assert row is not None


def test_length_tolerance_constant() -> None:
    assert LENGTH_TOLERANCE_M == 0.02
