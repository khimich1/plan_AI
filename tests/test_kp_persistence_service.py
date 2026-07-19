"""Tests for KpPersistenceService (A2 offers slice)."""

from __future__ import annotations

from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx


def test_save_kp_persists_plate_line(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    order = [
        {
            "name": "ПБ 60-12-8п",
            "length_m": 6.0,
            "width_m": 1.2,
            "load_class": 800,
            "qty": 2,
            "unit_price": 1000.0,
            "weight": 5000.0,
            "length_dm_raw": "60",
        }
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        order,
        customer_name="Test",
        db_path=db,
    )
    assert kp_id == 1
