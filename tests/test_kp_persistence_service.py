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
            # Explicit grade so resolve_concrete_grade short-circuits (no pb.db).
            "concrete_grade": "М500",
        }
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        order,
        customer_name="Test",
        db_path=db,
    )
    assert kp_id == 1


def test_saved_plate_concrete_spec_roundtrips_and_blank_stays_null(tmp_path) -> None:
    from core.kp.offers_read import get_kp_by_id
    from core.kp_order_data import order_data_from_kp_plates

    db = fx.make_iso_db(tmp_path)
    order = [
        {
            "name": "ПБ 60-12-8п",
            "length_m": 6.0,
            "width_m": 1.2,
            "load_class": 800,
            "qty": 1,
            "unit_price": 1000.0,
            "weight": 500.0,
            "length_dm_raw": "60",
            "concrete_grade": "М500",
            "line_id": "plate-spec",
            "frost_resistance": "F300",
            "waterproofness": "W12",
            "concrete_aggregate": "granite",
            "concrete_spec_source": "table",
        },
        {
            "name": "ПБ 54-12-8п",
            "length_m": 5.4,
            "width_m": 1.2,
            "load_class": 800,
            "qty": 1,
            "unit_price": 900.0,
            "weight": 400.0,
            "length_dm_raw": "54",
            "concrete_grade": "М400",
            "line_id": "plate-old",
        },
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        order,
        customer_name="Plates",
        product_type="plates",
        db_path=db,
    )
    rows = {row["name"]: row for row in order_data_from_kp_plates(get_kp_by_id(kp_id, db))}
    assert rows["ПБ 60-12-8п"]["frost_resistance"] == "F300"
    assert rows["ПБ 60-12-8п"]["waterproofness"] == "W12"
    assert rows["ПБ 60-12-8п"]["concrete_spec_source"] == "table"
    assert rows["ПБ 54-12-8п"].get("frost_resistance") in (None, "")
    assert rows["ПБ 54-12-8п"].get("concrete_spec_source") in (None, "")
