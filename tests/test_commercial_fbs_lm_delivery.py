"""Контракт котла доставки ФБС/ЛС/ЛМ и флага fbs_lm_delivery_enabled."""

from __future__ import annotations

import sqlite3

import pytest

from core.cargo_delivery_pricing import (
    CARGO_DELIVERY_TRUCK_CAPACITY_KG,
    FBS_LM_TRUCK_CAPACITY_KG,
    compute_fbs_lm_delivery,
)
from core.commercial_pricing import calculate_total_cost
from core.product_weight_catalog import (
    ProductWeightRecord,
    normalize_product_mark,
    upsert_product_weights,
)
from tests.helpers import kp_db_fixtures as fx

OLD_TOTAL_KEYS = (
    "total_qty",
    "subtotal",
    "vat_amount",
    "total_with_vat",
    "plate_delivery_total",
    "pile_delivery_total",
    "pile_trips",
    "pile_trip_pending_marks",
    "pile_delivery_ready",
    "plate_trips",
)


def _seed_weights(tmp_path, rows: list[tuple[str, str, float]]) -> str:
    db_path = fx.make_iso_db(tmp_path)
    records = [
        ProductWeightRecord(
            mark=mark,
            mark_norm=normalize_product_mark(mark),
            product_type=product_type,
            weight_kg=weight_kg,
            volume_m3=0.1,
            display_name=mark,
        )
        for mark, product_type, weight_kg in rows
    ]
    upsert_product_weights(db_path, records)
    return db_path


def _totals(order_data, *, db_path: str, enabled: bool, logistics_cost: float = 50_000.0):
    return calculate_total_cost(
        order_data,
        discount_percent=0,
        logistics_cost=logistics_cost,
        db_path=db_path,
        require_all_priced=False,
        pile_catalog_db_path=db_path,
        fbs_lm_delivery_enabled=enabled,
        weight_catalog_db_path=db_path,
    )


def test_fbs_lm_truck_capacity_is_20_tonnes() -> None:
    assert FBS_LM_TRUCK_CAPACITY_KG == 20_000.0
    assert CARGO_DELIVERY_TRUCK_CAPACITY_KG == 18_600.0


def test_compute_30t_times_trip_cost_is_two_trips(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "qty": 30,
            "unit_price": 10.0,
        }
    ]
    breakdown = compute_fbs_lm_delivery(
        order_data, trip_cost=50_000.0, catalog_db_path=db_path
    )
    assert breakdown.ready is True
    assert breakdown.cargo_kg == pytest.approx(30_000.0)
    assert breakdown.trips == 2
    assert breakdown.pending_marks == ()


def test_compute_boundary_20000_is_one_trip_20001_is_two(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 12.4.6-Т", "fbs", 1.0)])
    one_trip = compute_fbs_lm_delivery(
        [{"mark": "ФБС 12.4.6-Т", "product_type": "fbs", "qty": 20_000, "unit_price": 1}],
        trip_cost=50_000.0,
        catalog_db_path=db_path,
    )
    two_trips = compute_fbs_lm_delivery(
        [{"mark": "ФБС 12.4.6-Т", "product_type": "fbs", "qty": 20_001, "unit_price": 1}],
        trip_cost=50_000.0,
        catalog_db_path=db_path,
    )
    assert one_trip.trips == 1
    assert two_trips.trips == 2


def test_compute_qty_zero_lines_are_ignored(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 9.3.6-Т", "fbs", 350.0)])
    breakdown = compute_fbs_lm_delivery(
        [
            {"mark": "ФБС 9.3.6-Т", "product_type": "fbs", "qty": 0, "unit_price": 1},
            {"mark": "ФБС 9.3.6-Т", "product_type": "fbs", "qty": 10, "unit_price": 1},
        ],
        trip_cost=50_000.0,
        catalog_db_path=db_path,
    )
    assert breakdown.cargo_kg == pytest.approx(3500.0)
    assert breakdown.trips == 1


def test_compute_unknown_mark_is_pending_and_trips_zero(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 9.3.6-Т", "fbs", 350.0)])
    breakdown = compute_fbs_lm_delivery(
        [
            {"mark": "ФБС 9.3.6-Т", "product_type": "fbs", "qty": 10, "unit_price": 1},
            {"mark": "ФБС нет-в-каталоге", "product_type": "fbs", "qty": 2, "unit_price": 1},
        ],
        trip_cost=50_000.0,
        catalog_db_path=db_path,
    )
    assert breakdown.ready is False
    assert breakdown.trips == 0
    assert "ФБС нет-в-каталоге" in breakdown.pending_marks


def test_compute_three_types_share_one_boiler(tmp_path) -> None:
    db_path = _seed_weights(
        tmp_path,
        [
            ("ФБС 9.3.6-Т", "fbs", 1000.0),
            ("ЛС-14", "steps", 150.0),
            ("1ЛМ 27-11-14-4", "marches", 1330.0),
        ],
    )
    breakdown = compute_fbs_lm_delivery(
        [
            {"mark": "ФБС 9.3.6-Т", "product_type": "fbs", "qty": 10, "unit_price": 1},
            {"mark": "ЛС-14", "product_type": "steps", "qty": 40, "unit_price": 1},
            {"mark": "1ЛМ 27-11-14-4", "product_type": "marches", "qty": 3, "unit_price": 1},
        ],
        trip_cost=50_000.0,
        catalog_db_path=db_path,
    )
    # 10*1000 + 40*150 + 3*1330 = 10000+6000+3990 = 19990 кг → 1 рейс
    assert breakdown.ready is True
    assert breakdown.cargo_kg == pytest.approx(19_990.0)
    assert breakdown.trips == 1


def test_compute_accepts_product_kind_and_product_type(tmp_path) -> None:
    db_path = _seed_weights(
        tmp_path,
        [
            ("ФБС 9.3.6-Т", "fbs", 350.0),
            ("ЛС-12", "steps", 133.0),
            ("ЛМ 2,8", "marches", 975.0),
        ],
    )
    breakdown = compute_fbs_lm_delivery(
        [
            {"mark": "ФБС 9.3.6-Т", "product_kind": "fbs", "qty": 2, "unit_price": 1},
            {"mark": "ЛС-12", "product_kind": "step", "qty": 10, "unit_price": 1},
            {"mark": "ЛМ 2,8", "product_kind": "march", "qty": 1, "unit_price": 1},
        ],
        trip_cost=1.0,
        catalog_db_path=db_path,
    )
    assert breakdown.cargo_kg == pytest.approx(2 * 350 + 10 * 133 + 975)
    assert breakdown.ready is True


def test_flag_off_keeps_old_keys_and_zero_fbs_lm_totals(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "qty": 30,
            "unit_price": 10.0,
        }
    ]
    baseline = calculate_total_cost(
        order_data,
        discount_percent=0,
        logistics_cost=50_000.0,
        db_path=db_path,
        require_all_priced=False,
        pile_catalog_db_path=db_path,
    )
    disabled = _totals(order_data, db_path=db_path, enabled=False)
    for key in OLD_TOTAL_KEYS:
        assert disabled[key] == baseline[key]
    assert disabled["fbs_lm_delivery_total"] == 0.0
    assert disabled["fbs_lm_cargo_kg"] == 0.0
    assert disabled["fbs_lm_trips"] == 0
    assert disabled["fbs_lm_delivery_ready"] is True
    assert disabled["fbs_lm_pending_marks"] == []
    assert disabled["total_with_vat"] == 300.0


def test_flag_on_30t_adds_100000_delivery(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "qty": 30,
            "unit_price": 10.0,
        }
    ]
    totals = _totals(order_data, db_path=db_path, enabled=True, logistics_cost=50_000.0)
    assert totals["fbs_lm_delivery_ready"] is True
    assert totals["fbs_lm_cargo_kg"] == pytest.approx(30_000.0)
    assert totals["fbs_lm_trips"] == 2
    assert totals["fbs_lm_delivery_total"] == 100_000.0
    assert totals["fbs_lm_pending_marks"] == []
    assert totals["plate_delivery_total"] == 0.0
    assert totals["total_with_vat"] == 100_300.0  # 30*10 + 100000


def test_flag_on_pending_mark_zeroes_boiler(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 9.3.6-Т", "fbs", 350.0)])
    order_data = [
        {
            "mark": "ФБС нет-в-каталоге",
            "product_type": "fbs",
            "qty": 4,
            "unit_price": 20.0,
        }
    ]
    totals = _totals(order_data, db_path=db_path, enabled=True, logistics_cost=50_000.0)
    assert totals["fbs_lm_delivery_ready"] is False
    assert totals["fbs_lm_delivery_total"] == 0.0
    assert totals["fbs_lm_trips"] == 0
    assert "ФБС нет-в-каталоге" in totals["fbs_lm_pending_marks"]
    assert totals["total_with_vat"] == 80.0


def test_mixed_plates_and_fbs_boilers_are_independent(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 19_000.0)])
    order_data = [
        {
            "name": "ПБ",
            "product_type": "plates",
            "qty": 66,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        },
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "qty": 1,
            "unit_price": 100.0,
        },
    ]
    totals = _totals(order_data, db_path=db_path, enabled=True, logistics_cost=50_000.0)
    # Плиты: qty=66 × ~283 кг > 18 600 → 2 рейса; ФБС 19 000 кг / 20 000 → 1 рейс.
    assert totals["plate_trips"] == 2
    assert totals["fbs_lm_trips"] == 1
    assert totals["plate_delivery_total"] == 100_000.0
    assert totals["fbs_lm_delivery_total"] == 50_000.0
    assert totals["total_with_vat"] == 660.0 + 100.0 + 150_000.0


def test_schema_adds_fbs_lm_flag_and_is_idempotent(tmp_path) -> None:
    from core import kp_db_schema

    db_path = str(tmp_path / "legacy.db")
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL,
                customer_name TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL UNIQUE,
                status TEXT
            )
            """
        )
        conn.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) VALUES ('2026-01-01', 'old')"
        )
        conn.execute("INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')")
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema.ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(kp_meta)")}
        flag = conn.execute(
            "SELECT fbs_lm_delivery_enabled FROM kp_meta WHERE kp_id = 1"
        ).fetchone()[0]
    assert "fbs_lm_delivery_enabled" in cols
    assert flag in (0, 0.0, None)


def test_new_fbs_preview_metadata_enables_delivery_flag() -> None:
    from types import SimpleNamespace

    from app.services.commercial_draft_service import CommercialDraftService

    preview = SimpleNamespace(
        warnings=[],
        unparsed_lines=[],
        normalized_text="",
        normalized_lines=[],
        order_data=[],
        total_sum=0,
    )
    metadata = CommercialDraftService().build_fbs_preview_metadata(
        preview=preview,
        base_metadata={},
        source_type="text",
        original_text="",
        ocr_text="",
        input_text="ФБС 24.6.6-Т B25 1",
        last_source_filename="",
        fbs_batches=[],
        source_metadata={},
    )
    assert metadata["fbs_lm_delivery_enabled"] is True


def test_save_kp_persists_flag_and_includes_boiler_in_total(tmp_path) -> None:
    from core.kp_persistence_service import KpPersistenceService

    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "product_kind": "fbs",
            "qty": 30,
            "unit_price": 10.0,
            "concrete_grade": "B25",
        }
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "10.09.2026",
        order_data,
        customer_name="ФБС клиент",
        product_type="fbs",
        db_path=db_path,
        logistics_cost=50_000.0,
        fbs_lm_delivery_enabled=True,
    )
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        meta = conn.execute(
            "SELECT fbs_lm_delivery_enabled FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
        offer = conn.execute(
            "SELECT total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert int(meta["fbs_lm_delivery_enabled"]) == 1
    assert offer["total_amount"] == pytest.approx(100_300.0)


def test_save_kp_without_flag_does_not_add_boiler(tmp_path) -> None:
    from core.kp_persistence_service import KpPersistenceService

    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "product_kind": "fbs",
            "qty": 30,
            "unit_price": 10.0,
            "concrete_grade": "B25",
        }
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "10.09.2026",
        order_data,
        customer_name="Старый клиент",
        product_type="fbs",
        db_path=db_path,
        logistics_cost=50_000.0,
    )
    with sqlite3.connect(db_path) as conn:
        flag = conn.execute(
            "SELECT fbs_lm_delivery_enabled FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
        total = conn.execute(
            "SELECT total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
    assert int(flag or 0) == 0
    assert total == pytest.approx(300.0)


def test_export_lines_fbs_only_uses_generic_label() -> None:
    from core.commercial_pricing import kp_delivery_export_lines

    lines = kp_delivery_export_lines(
        {
            "plate_delivery_total": 0.0,
            "pile_delivery_total": 0.0,
            "pile_delivery_ready": True,
            "fbs_lm_delivery_total": 100_000.0,
            "fbs_lm_delivery_ready": True,
            "fbs_lm_trips": 2,
            "plate_trips": 0,
            "pile_trips": 0,
        },
        plate_trip_cost=50_000.0,
        pile_trip_cost=0.0,
    )
    assert len(lines) == 1
    assert lines[0]["label"] == "Услуга по доставке грузов"
    assert lines[0]["trips"] == 2
    assert lines[0]["amount"] == 100_000.0


def test_export_lines_plates_and_fbs_use_specific_labels() -> None:
    from core.commercial_pricing import kp_delivery_export_lines

    lines = kp_delivery_export_lines(
        {
            "plate_delivery_total": 50_000.0,
            "plate_trips": 1,
            "pile_delivery_total": 0.0,
            "pile_delivery_ready": True,
            "fbs_lm_delivery_total": 50_000.0,
            "fbs_lm_delivery_ready": True,
            "fbs_lm_trips": 1,
        },
        plate_trip_cost=50_000.0,
        pile_trip_cost=0.0,
    )
    labels = [line["label"] for line in lines]
    assert labels == ["Доставка плит", "Доставка ФБС/ЛС/ЛМ"]


def test_xlsx_mono_fbs_with_flag_contains_delivery_row(tmp_path) -> None:
    import pandas as pd

    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "product_kind": "fbs",
            "qty": 30,
            "unit_price": 10.0,
            "concrete_grade": "B25",
        }
    ]
    buf = generate_commercial_offer_xlsx(
        order_data=order,
        offer_number="FBS1",
        offer_date="10.09.2026",
        customer_name="ООО Тест",
        logistics_cost=50_000.0,
        fbs_lm_delivery_enabled=True,
        weight_catalog_db_path=db_path,
        pile_catalog_db_path=db_path,
    )
    df = pd.read_excel(buf, sheet_name="КП", header=None)
    names = " ".join(str(v) for v in df.values.ravel() if pd.notna(v))
    assert "Услуга по доставке грузов" in names


def test_xlsx_without_flag_has_no_fbs_delivery_row(tmp_path) -> None:
    import pandas as pd

    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "product_kind": "fbs",
            "qty": 30,
            "unit_price": 10.0,
            "concrete_grade": "B25",
        }
    ]
    buf = generate_commercial_offer_xlsx(
        order_data=order,
        offer_number="FBS0",
        offer_date="10.09.2026",
        customer_name="ООО Тест",
        logistics_cost=50_000.0,
        fbs_lm_delivery_enabled=False,
        weight_catalog_db_path=db_path,
        pile_catalog_db_path=db_path,
    )
    df = pd.read_excel(buf, sheet_name="КП", header=None)
    names = " ".join(str(v) for v in df.values.ravel() if pd.notna(v))
    assert "Доставка ФБС/ЛС/ЛМ" not in names
    assert "Услуга по доставке грузов" not in names


def test_discount_does_not_apply_to_fbs_lm_boiler(tmp_path) -> None:
    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "qty": 30,
            "unit_price": 10.0,
        }
    ]
    totals = calculate_total_cost(
        order_data,
        discount_percent=50,
        logistics_cost=50_000.0,
        db_path=db_path,
        require_all_priced=False,
        fbs_lm_delivery_enabled=True,
        weight_catalog_db_path=db_path,
        pile_catalog_db_path=db_path,
    )
    assert totals["fbs_lm_delivery_total"] == 100_000.0
    assert totals["total_with_vat"] == 150.0 + 100_000.0  # 30*10*0.5 + delivery


def test_update_logistics_cost_old_kp_does_not_enable_boiler(tmp_path) -> None:
    from core.kp.offers_write import update_kp_logistics_cost
    from core.kp_persistence_service import KpPersistenceService

    db_path = _seed_weights(tmp_path, [("ФБС 24.6.6-Т", "fbs", 1000.0)])
    order_data = [
        {
            "name": "ФБС 24.6.6-Т",
            "mark": "ФБС 24.6.6-Т",
            "product_type": "fbs",
            "product_kind": "fbs",
            "qty": 30,
            "unit_price": 10.0,
            "concrete_grade": "B25",
        }
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "10.09.2026",
        order_data,
        customer_name="Архив",
        product_type="fbs",
        db_path=db_path,
        logistics_cost=10_000.0,
        fbs_lm_delivery_enabled=False,
    )
    assert update_kp_logistics_cost(kp_id, 50_000.0, db_path)
    with sqlite3.connect(db_path) as conn:
        total = conn.execute(
            "SELECT total_amount FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
        flag = conn.execute(
            "SELECT fbs_lm_delivery_enabled FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
    assert int(flag or 0) == 0
    assert total == pytest.approx(300.0)
