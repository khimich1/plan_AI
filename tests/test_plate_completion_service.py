"""Unit tests for PlateCompletionService orchestration (A2 stage 6)."""

from __future__ import annotations

import pytest

from app.services.plate_completion_service import PlateCompletionService
from core.kp_db_common import _connect
from tests.helpers import kp_db_fixtures as fx

PLATE_60 = "ПБ 60-12-8п"


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def _complete_on_cursor(
    db: str,
    kp_id: int,
    plate_name: str,
    *,
    qty: int = 1,
    length_m: float = 6.0,
    width_m: float = 1.2,
    load_class: int = 800,
    length_dm_raw: str = "",
    kp_plate_id: int | None = None,
    allow_cross_kp: bool = False,
    plan_ids: list[str] | None = None,
) -> dict:
    payload: dict = {
        "plate_name": plate_name,
        "qty": qty,
        "length_m": length_m,
        "width_m": width_m,
        "load_class": load_class,
        "length_dm_raw": length_dm_raw,
    }
    if kp_plate_id is not None:
        payload["kp_plate_id"] = kp_plate_id
    conn = _connect(db)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()
        result = PlateCompletionService.complete_plates_on_cursor(
            cur,
            kp_id,
            [payload],
            production_day=1,
            plan_ids=plan_ids,
            allow_cross_kp=allow_cross_kp,
        )
        conn.commit()
        return result
    finally:
        conn.close()


def test_service_length_dm_raw_step0(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=5.98,
        width_m=1.2,
        qty=2,
        status="в плане",
        length_dm_raw="598",
    )
    fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=5.99,
        width_m=1.2,
        qty=3,
        status="в плане",
        length_dm_raw="599",
    )

    result = _complete_on_cursor(
        iso_db,
        1,
        PLATE_60,
        qty=2,
        length_m=5.98,
        length_dm_raw="598",
    )

    assert result["completed_count"] == 2
    assert result["unmoved"] == []
    rows = fx.plates_snapshot(iso_db, 1)
    assert len(rows) == 1
    assert rows[0]["length_dm_raw"] == "599"
    assert rows[0]["qty"] == 3
    completed = fx.completed_snapshot(iso_db, 1)
    assert sum(r["qty"] for r in completed) == 2


def test_service_unmoved_when_row_missing(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)

    result = _complete_on_cursor(iso_db, 1, PLATE_60, qty=2)

    assert result["completed_count"] == 0
    assert len(result["unmoved"]) == 1
    assert result["unmoved"][0]["qty"] == 2


def test_service_kp_plate_id_direct_lookup(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    plate_id = fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE_60,
        length_m=6.0,
        width_m=1.2,
        qty=2,
        status="в плане",
        day_number=1,
        plan_id="plan-a",
    )

    result = _complete_on_cursor(
        iso_db,
        1,
        "wrong name",
        qty=1,
        kp_plate_id=plate_id,
        plan_ids=["plan-a"],
    )

    assert result["completed_count"] == 1
    assert result["unmoved"] == []
