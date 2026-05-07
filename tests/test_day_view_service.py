"""Unit-тесты для :mod:`app.services.day_view_service` (агрегация plates_info)."""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.schemas.production import DayPlateInfo
from app.services.day_view_service import _aggregate_plates_for_track_from_db


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
    plates = _aggregate_plates_for_track_from_db(
        track,
        {101: row_live},
    )
    assert len(plates) == 1
    assert plates[0]["kp_plate_id"] == 101
    assert plates[0]["write_off_completed"] is False

    row_snap = {**row_live, "is_completed_snapshot": True}
    plates_done = _aggregate_plates_for_track_from_db(track, {101: row_snap})
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
