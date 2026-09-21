"""Unit-тесты для :mod:`app.services.day_view_service` (агрегация plates_info)."""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.schemas.production import DayPlateInfo
from app.services.day_view_service import (
    aggregate_plates_for_track,
    aggregate_plates_for_track_from_db,
)


def test_aggregate_plates_for_track_from_db_write_off_completed_follows_snapshot_flag():
    """``write_off_completed`` зеркалит ``is_completed_snapshot`` в строке БД-агрегации."""
    track = {
        "items": [
            {
                "kp_plate_id": 101,
                "secondary_cuts": [],
            },
        ],
    }
    row_live = {
        "kp_id": 1,
        "plate_name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "length_dm_raw": "",
        "customer": "Клиент",
        "kp_date": "21.04.2026",
        "reinforcement": 0,
        "is_completed_snapshot": False,
    }
    plates = aggregate_plates_for_track_from_db(
        track,
        {101: row_live},
    )
    assert len(plates) == 1
    assert plates[0]["kp_plate_id"] == 101
    assert plates[0]["write_off_completed"] is False

    row_snap = {**row_live, "is_completed_snapshot": True}
    plates_done = aggregate_plates_for_track_from_db(track, {101: row_snap})
    assert plates_done[0]["write_off_completed"] is True


def test_day_plate_info_accepts_write_off_completed_field():
    """Схема API дня принимает флаг списания для фронта."""
    m = DayPlateInfo(
        customer="X",
        plate_name="ПБ 60-12-8п",
        kp_date="d",
        kp_id=1,
        length_m=6.0,
        width_mm=1200,
        qty=2,
        write_off_completed=True,
    )
    assert m.write_off_completed is True

    m_default = DayPlateInfo(
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_mm=1200,
        qty=1,
    )
    assert m_default.write_off_completed is False
    assert m_default.concrete_grade == "—"


def _lookup_with_grade(grade: str):
    def lookup(_length_m: float, _width_mm: int) -> dict:
        return {
            "kp_id": 1,
            "kp_date": "21.04.2026",
            "customer": "Клиент",
            "plate_name": "ПБ 60-12-8п",
            "reinforcement": 8.0,
            "concrete_grade": grade,
        }

    return lookup


def test_aggregate_plates_legacy_missing_grade_is_dash():
    track = {
        "items": [
            {
                "mode": "solid",
                "length": 6.0,
                "width": 1.2,
                "load_code": 8,
            }
        ],
    }
    plates = aggregate_plates_for_track(track, _lookup_with_grade(""))
    assert len(plates) == 1
    assert plates[0]["concrete_grade"] == "—"


def test_aggregate_plates_uses_real_grade_from_lookup():
    track = {
        "items": [
            {
                "mode": "solid",
                "length": 6.0,
                "width": 1.2,
                "load_code": 8,
            }
        ],
    }
    plates = aggregate_plates_for_track(track, _lookup_with_grade("М500"))
    assert plates[0]["concrete_grade"] == "М500"


def test_aggregate_plates_rescue_falls_back_to_parent_item_grade():
    track = {
        "label": "РЕСКЬЮ",
        "items": [
            {
                "mode": "solid",
                "length": 6.0,
                "width": 1.2,
                "load_code": 8,
                "plate_name": "ПБ 60-12-8п",
                "concrete_grade": "М500",
            }
        ],
    }
    plates = aggregate_plates_for_track(track, _lookup_with_grade(""))
    assert plates[0]["concrete_grade"] == "М500"
