"""Q1 guard: split rows must keep nomenclature_id and length_dm_raw."""

from __future__ import annotations

import pytest

from core import kp_db
from tests.helpers import kp_db_fixtures as fx

PLATE = "ПБ 60-12-8п"


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def test_mark_plates_as_planned_split_preserves_identity(iso_db: str) -> None:
    kp_id = kp_db.save_kp_to_db(
        "01.06.2026",
        [
            {
                "name": PLATE,
                "qty": 3,
                "unit_price": 100.0,
                "length_m": 6.0,
                "width_m": 1.2,
                "load_class": 800,
                "nomenclature_id": 42,
                "length_dm_raw": "600",
            }
        ],
        db_path=iso_db,
        status="в работе",
    )

    result = kp_db.mark_plates_as_planned(
        kp_id,
        PLATE,
        qty_to_plan=1,
        plan_id="plan_split_test",
        db_path=iso_db,
    )
    assert result["success"] is True
    assert result["split_count"] == 1

    rows = fx.plates_snapshot(iso_db, kp_id)
    in_production = [r for r in rows if r["status"] == "в производстве"]
    assert len(in_production) == 1
    assert int(in_production[0]["nomenclature_id"]) == 42
    assert in_production[0]["length_dm_raw"] == "600"

    in_plan = [r for r in rows if r["status"] == "в плане"]
    assert len(in_plan) == 1
    assert int(in_plan[0]["nomenclature_id"]) == 42
    assert in_plan[0]["length_dm_raw"] == "600"


def test_return_plates_to_production_split_preserves_identity(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    plate_id = fx.seed_plate(
        iso_db,
        kp_id=1,
        plate_name=PLATE,
        length_m=6.0,
        width_m=1.2,
        qty=3,
        status="в плане",
        length_dm_raw="600",
        nomenclature_id=42,
        plan_id="plan_ret",
    )

    ok = kp_db.return_plates_to_production(
        1,
        PLATE,
        qty=1,
        db_path=iso_db,
    )
    assert ok is True

    rows = fx.plates_snapshot(iso_db, 1)
    in_production = [r for r in rows if r["status"] == "в производстве"]
    assert len(in_production) == 1
    assert int(in_production[0]["nomenclature_id"]) == 42
    assert in_production[0]["length_dm_raw"] == "600"
    in_plan = [r for r in rows if r["id"] == plate_id]
    assert in_plan[0]["qty"] == 2
