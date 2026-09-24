"""Доставка свай длиннее 13 м: группы длин отдельно от короткого котла."""

from __future__ import annotations

import math

import pytest

from core.commercial_pricing import (
    calculate_total_cost,
    coerce_long_pile_delivery_enabled,
    kp_delivery_export_lines,
)
from core.pile_catalog import PileCatalogEntry, resolve_catalog_for_mark, upsert_pile_catalog
from core.pile_trip_pricing import (
    PILE_REMAINDER_TRUCK_CAPACITY_KG,
    coerce_long_pile_delivery,
    compute_long_pile_groups,
)
from tests.helpers import kp_db_fixtures as fx

# Числа из текущего pile_catalog. Прочерк pcs у С160 не заполняем.
_CATALOG = [
    PileCatalogEntry("С100.35", 10.0, 350, 1.24, 3100.0, 6),
    PileCatalogEntry("С130.35", 13.0, 350, 1.61, 4030.0, 5),
    PileCatalogEntry("С138.35", 13.8, 350, 1.71, 4280.0, 4),
    PileCatalogEntry("С140.35", 14.0, 350, 1.73, 4330.0, 4),
    PileCatalogEntry("С140.40", 14.0, 400, 2.26, 5650.0, 3),
    PileCatalogEntry("С150.35", 15.0, 350, 1.86, 4650.0, 4),
    PileCatalogEntry("С160.35", 16.0, 350, 1.98, 4950.0, None),
]


def _lookup(mark: str) -> PileCatalogEntry | None:
    return resolve_catalog_for_mark(mark, _CATALOG)


def _line(
    mark: str,
    qty: int,
    product_type: str = "piles",
    unit_price: float | None = None,
) -> dict:
    row = {"mark": mark, "name": mark, "qty": qty, "product_type": product_type}
    if unit_price is not None:
        row["unit_price"] = unit_price
    return row


def _length(split, key: int):
    matches = [group for group in split.lengths if group.length_key == key]
    assert len(matches) == 1, f"length {key} missing in {[g.length_key for g in split.lengths]}"
    return matches[0]


def test_s130_stays_in_short_boiler_s140_and_s150_do_not() -> None:
    split = compute_long_pile_groups(
        [
            _line("С130.35", 5),
            _line("С140.35", 4),
            _line("С150.35", 4),
        ],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert split.short.full_trips == 1
    assert split.short.remainder_kg == pytest.approx(0.0)
    assert split.short.pending_marks == ()
    assert split.short.total_trips == 1
    assert {group.length_key for group in split.lengths} == {140, 150}
    assert _length(split, 140).trips == 1
    assert _length(split, 150).trips == 1


def test_short_full_trip_does_not_absorb_long_pile_kilograms() -> None:
    split = compute_long_pile_groups(
        [
            _line("С100.35", 6),
            _line("С140.35", 1),
        ],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert split.short.full_trips == 1
    assert split.short.remainder_kg == pytest.approx(0.0)
    assert split.short.remainder_trips == 0
    assert split.short.total_trips == 1
    assert _length(split, 140).remainder_kg == pytest.approx(4330.0)


def test_s150_five_pieces_is_one_full_plus_remainder_trip() -> None:
    split = compute_long_pile_groups(
        [_line("С150.35", 5)],
        overrides=None,
        catalog_lookup=_lookup,
    )
    group = _length(split, 150)
    assert group.full_trips == 1
    assert group.remainder_kg == pytest.approx(4650.0)
    assert group.remainder_trips == 1
    assert group.trips == 2
    assert split.short.total_trips == 0
    assert math_ceil_ok(4650.0)


def test_s140_sections_share_one_remainder_and_not_15m() -> None:
    split = compute_long_pile_groups(
        [
            _line("С140.35", 1),
            _line("С140.40", 1),
            _line("С150.35", 1),
        ],
        overrides=None,
        catalog_lookup=_lookup,
    )
    length_140 = _length(split, 140)
    length_150 = _length(split, 150)
    assert length_140.remainder_kg == pytest.approx(4330.0 + 5650.0)
    assert length_140.remainder_trips == 1
    assert length_140.trips == 1
    assert length_150.remainder_kg == pytest.approx(4650.0)
    assert length_150.trips == 1
    assert len(split.lengths) == 2


def test_load_suffix_stays_in_length_key_140() -> None:
    split = compute_long_pile_groups(
        [_line("С140.35-12", 4)],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert [group.length_key for group in split.lengths] == [140]
    assert _length(split, 140).trips == 1
    assert split.short.total_trips == 0


def test_s160_without_n_does_not_zero_15m_trips() -> None:
    split = compute_long_pile_groups(
        [
            _line("С150.35", 5),
            _line("С160.35-12", 4),
        ],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert _length(split, 150).trips == 2
    assert _length(split, 150).pending_marks == ()
    length_160 = _length(split, 160)
    assert length_160.pending_marks == ("С160.35-12",)
    assert length_160.trips == 0
    assert length_160.ready is False


def test_explicit_zero_override_closes_mark() -> None:
    split = compute_long_pile_groups(
        [_line("С160.35-12", 4)],
        overrides={"С160.35-12": 0},
        catalog_lookup=_lookup,
    )
    length_160 = _length(split, 160)
    assert length_160.pending_marks == ()
    assert length_160.ready is True
    assert length_160.override_trips == 0
    assert length_160.trips == 0


def test_bridge_c14_joins_length_key_140() -> None:
    split = compute_long_pile_groups(
        [_line("C14-40T4", 3, "bridge_piles")],
        overrides=None,
        catalog_lookup=_lookup,
    )
    group = _length(split, 140)
    assert group.full_trips == 1
    assert group.remainder_kg == pytest.approx(0.0)
    assert group.trips == 1
    assert group.pending_marks == ()
    assert split.short.pending_marks == ()


def test_bridge_c18_is_long_group_without_norm() -> None:
    split = compute_long_pile_groups(
        [_line("C18-40T8", 49, "bridge_piles")],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert split.short.pending_marks == ()
    assert split.short.total_trips == 0
    group = _length(split, 180)
    assert group.pending_marks == ("C18-40T8",)
    assert group.trips == 0
    assert group.full_trips == 0
    assert group.ready is False


def test_mark_without_geometry_stays_pending_in_short_boiler() -> None:
    split = compute_long_pile_groups(
        [_line("XYZ-99", 3)],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert split.lengths == ()
    assert split.short.ready is False
    assert "XYZ-99" in split.short.pending_marks
    assert split.short.total_trips == 0


def test_length_138_is_long_and_exact_13m_is_short() -> None:
    split = compute_long_pile_groups(
        [
            _line("С130.35", 5),
            _line("С138.35", 4),
        ],
        overrides=None,
        catalog_lookup=_lookup,
    )
    assert split.short.total_trips == 1
    assert [group.length_key for group in split.lengths] == [138]
    assert _length(split, 138).trips == 1


def math_ceil_ok(kg: float) -> bool:
    assert math.ceil(kg / PILE_REMAINDER_TRUCK_CAPACITY_KG) == 1
    return True


def _seed_catalog(tmp_path) -> str:
    db_path = fx.make_iso_db(tmp_path)
    upsert_pile_catalog(db_path, list(_CATALOG))
    return db_path


def _money(
    order,
    *,
    db_path: str,
    enabled: bool,
    pile_logistics_cost: float = 0.0,
    tariffs: dict | None = None,
    overrides: dict | None = None,
    discount_percent: float = 0.0,
):
    return calculate_total_cost(
        order,
        discount_percent=discount_percent,
        logistics_cost=0.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=pile_logistics_cost,
        pile_trip_overrides=overrides,
        pile_catalog_db_path=db_path,
        long_pile_delivery_enabled=enabled,
        long_pile_delivery=tariffs,
    )


def _length_row(totals: dict, key: int) -> dict:
    rows = [row for row in totals["long_pile_lengths"] if row["length_key"] == key]
    assert len(rows) == 1, totals["long_pile_lengths"]
    return rows[0]


def test_short_boiler_tariff_times_one_trip(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [_line("С100.35", 6, unit_price=10.0)],
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
    )
    assert totals["pile_delivery_ready"] is True
    assert totals["pile_trips"] == 1
    assert totals["pile_delivery_total"] == 1000.0
    assert totals["long_pile_delivery_total"] == 0.0
    assert totals["long_pile_lengths"] == []
    assert totals["long_pile_pending_marks"] == []
    assert totals["total_with_vat"] == 60.0 + 1000.0


def test_length_15m_tariff_times_two_trips(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [_line("С150.35", 5, unit_price=10.0)],
        db_path=db_path,
        enabled=True,
        tariffs={"150": {"trip_cost": 5000}},
    )
    row = _length_row(totals, 150)
    assert row["trips"] == 2
    assert row["trip_cost"] == 5000.0
    assert row["ready"] is True
    assert row["amount"] == 10000.0
    assert row["pending_marks"] == []
    assert totals["pile_delivery_total"] == 0.0
    assert totals["long_pile_delivery_total"] == 10000.0
    assert totals["total_with_vat"] == 50.0 + 10000.0


def test_two_lengths_are_priced_separately(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [
            _line("С140.35", 4, unit_price=1.0),
            _line("С150.35", 5, unit_price=1.0),
        ],
        db_path=db_path,
        enabled=True,
        tariffs={"140": {"trip_cost": 8000}, "150": {"trip_cost": 5000}},
    )
    assert _length_row(totals, 140)["amount"] == 8000.0
    assert _length_row(totals, 150)["amount"] == 10000.0
    assert totals["long_pile_delivery_total"] == 18000.0
    assert totals["pile_delivery_total"] == 0.0


def test_missing_16m_tariff_does_not_zero_short_or_15m(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [
            _line("С100.35", 6, unit_price=10.0),
            _line("С150.35", 5, unit_price=10.0),
            _line("С160.35-12", 4, unit_price=10.0),
        ],
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
        tariffs={"150": {"trip_cost": 5000}},
        overrides={"С160.35-12": 2},
    )
    assert totals["pile_delivery_total"] == 1000.0
    assert _length_row(totals, 150)["amount"] == 10000.0
    row_160 = _length_row(totals, 160)
    assert row_160["trip_cost"] is None
    assert row_160["ready"] is False
    assert row_160["amount"] == 0.0
    assert totals["long_pile_delivery_total"] == 10000.0
    assert totals["total_with_vat"] == 150.0 + 11000.0


def test_s160_manual_trips_times_tariff(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [_line("С160.35-12", 4, unit_price=10.0)],
        db_path=db_path,
        enabled=True,
        tariffs={"160": {"trip_cost": 7000}},
        overrides={"С160.35-12": 2},
    )
    row = _length_row(totals, 160)
    assert row["trips"] == 2
    assert row["pending_marks"] == []
    assert row["ready"] is True
    assert row["amount"] == 14000.0
    assert totals["long_pile_pending_marks"] == []
    assert totals["long_pile_delivery_total"] == 14000.0


def test_s160_without_n_zeros_only_that_length(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [
            _line("С100.35", 6, unit_price=1.0),
            _line("С150.35", 4, unit_price=1.0),
            _line("С160.35-12", 4, unit_price=1.0),
        ],
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
        tariffs={"150": {"trip_cost": 5000}, "160": {"trip_cost": 7000}},
    )
    assert totals["pile_delivery_total"] == 1000.0
    assert _length_row(totals, 150)["amount"] == 5000.0
    assert _length_row(totals, 160)["amount"] == 0.0
    assert _length_row(totals, 160)["ready"] is False
    assert totals["long_pile_pending_marks"] == ["С160.35-12"]
    assert totals["long_pile_delivery_total"] == 5000.0
    assert totals["pile_delivery_ready"] is True


def test_explicit_zero_tariff_is_ready_without_export_line(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    order = [_line("С150.35", 4, unit_price=10.0)]
    totals = _money(
        order,
        db_path=db_path,
        enabled=True,
        tariffs={"150": {"trip_cost": 0}},
    )
    row = _length_row(totals, 150)
    assert row["trip_cost"] == 0.0
    assert row["ready"] is True
    assert row["amount"] == 0.0
    assert totals["long_pile_delivery_total"] == 0.0
    lines = kp_delivery_export_lines(
        totals, plate_trip_cost=0.0, pile_trip_cost=0.0
    )
    assert lines == []


def test_flag_off_keeps_one_boiler_and_zero_long_keys(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    order = [
        _line("С100.35", 6, unit_price=10.0),
        _line("С160.35-12", 4, unit_price=10.0),
    ]
    totals = _money(
        order,
        db_path=db_path,
        enabled=False,
        pile_logistics_cost=1000.0,
        tariffs={"160": {"trip_cost": 7000}},
    )
    assert totals["pile_delivery_ready"] is False
    assert totals["pile_delivery_total"] == 0.0
    assert totals["pile_trips"] == 0
    assert "С160.35-12" in totals["pile_trip_pending_marks"]
    assert totals["long_pile_delivery_total"] == 0.0
    assert totals["long_pile_lengths"] == []
    assert totals["long_pile_pending_marks"] == []
    assert totals["total_with_vat"] == 100.0


def test_default_signature_without_flag_matches_one_boiler(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    order = [_line("С100.35", 6, unit_price=10.0)]
    totals = calculate_total_cost(
        order,
        discount_percent=0,
        logistics_cost=0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=1000.0,
        pile_catalog_db_path=db_path,
    )
    assert totals["pile_delivery_total"] == 1000.0
    assert totals["long_pile_delivery_total"] == 0.0
    assert totals["long_pile_lengths"] == []
    assert totals["long_pile_pending_marks"] == []


def test_export_label_for_ready_length_14m(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [
            _line("С100.35", 6, unit_price=1.0),
            _line("С140.35", 4, unit_price=1.0),
        ],
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
        tariffs={"140": {"trip_cost": 5000}},
    )
    lines = kp_delivery_export_lines(
        totals, plate_trip_cost=0.0, pile_trip_cost=1000.0
    )
    labels = [line["label"] for line in lines]
    assert labels == ["Доставка свай", "Доставка свай 14,0 м"]
    long_line = lines[1]
    assert long_line["trips"] == 1
    assert long_line["unit_price"] == 5000.0
    assert long_line["amount"] == 5000.0


def test_export_label_uses_comma_for_13_8m(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [_line("С138.35", 4, unit_price=1.0)],
        db_path=db_path,
        enabled=True,
        tariffs={"138": {"trip_cost": 3000}},
    )
    lines = kp_delivery_export_lines(
        totals, plate_trip_cost=0.0, pile_trip_cost=0.0
    )
    assert [line["label"] for line in lines] == ["Доставка свай 13,8 м"]


def test_flag_on_short_piles_only_use_generic_delivery_label(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [_line("С100.35", 6, unit_price=1.0)],
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
    )
    lines = kp_delivery_export_lines(
        totals, plate_trip_cost=0.0, pile_trip_cost=1000.0
    )
    assert [line["label"] for line in lines] == ["Услуга по доставке грузов"]


def test_schema_adds_long_pile_columns_and_second_ensure_is_safe(tmp_path) -> None:
    import os
    import sqlite3

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
    kp_db_schema._schema_ready.discard(os.path.abspath(db_path))
    kp_db_schema.ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(kp_meta)")}
        flag, raw_json = conn.execute(
            """
            SELECT long_pile_delivery_enabled, long_pile_delivery_json
            FROM kp_meta WHERE kp_id = 1
            """
        ).fetchone()
    assert "long_pile_delivery_enabled" in cols
    assert "long_pile_delivery_json" in cols
    assert flag in (0, 0.0, None)
    assert raw_json in (None, "")


def test_legacy_raw_without_columns_is_flag_off_and_empty_tariffs() -> None:
    from core.commercial_pricing import coerce_long_pile_delivery_enabled
    from core.pile_trip_pricing import coerce_long_pile_delivery

    raw = {"kp_id": 1, "status": "в работе"}
    assert coerce_long_pile_delivery_enabled(raw.get("long_pile_delivery_enabled")) is False
    assert coerce_long_pile_delivery(raw.get("long_pile_delivery_json")) == {}


def test_explicit_zero_tariff_is_stored_and_empty_is_not() -> None:
    from core.pile_trip_pricing import coerce_long_pile_delivery, dumps_long_pile_delivery

    assert dumps_long_pile_delivery({}) is None
    assert dumps_long_pile_delivery({"140": {"trip_cost": ""}}) is None
    stored = dumps_long_pile_delivery({"140": {"trip_cost": 0}, "160": {"trip_cost": 7000}})
    assert stored is not None
    parsed = coerce_long_pile_delivery(stored)
    assert parsed[140] == 0.0
    assert parsed[160] == 7000.0
    assert "140" in stored


def test_garbage_tariff_json_does_not_price_that_length(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    totals = _money(
        [
            _line("С150.35", 4, unit_price=10.0),
            _line("С140.35", 4, unit_price=10.0),
        ],
        db_path=db_path,
        enabled=True,
        tariffs='{"150": {"trip_cost": "мусор"}, "140": {"trip_cost": 5000}}',
    )
    assert _length_row(totals, 150)["trip_cost"] is None
    assert _length_row(totals, 150)["amount"] == 0.0
    assert _length_row(totals, 140)["amount"] == 5000.0
    assert totals["long_pile_delivery_total"] == 5000.0

    broken = _money(
        [_line("С150.35", 4, unit_price=10.0)],
        db_path=db_path,
        enabled=True,
        tariffs="{not json",
    )
    assert _length_row(broken, 150)["trip_cost"] is None
    assert broken["long_pile_delivery_total"] == 0.0


def _priced_mix() -> list[dict]:
    return [
        _line("С100.35", 6, unit_price=10.0),
        _line("С150.35", 5, unit_price=10.0),
        _line("С160.35-12", 4, unit_price=10.0),
    ]


def test_new_pile_draft_enables_long_pile_delivery_flag() -> None:
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
    metadata = CommercialDraftService().build_pile_preview_metadata(
        preview=preview,
        base_metadata={},
        source_type="text",
        original_text="",
        ocr_text="",
        input_text="С150.35 B25 5",
        last_source_filename="",
        pile_batches=[],
        source_metadata={},
    )
    assert metadata["long_pile_delivery_enabled"] is True
    assert metadata["long_pile_delivery"] == {}


def test_save_roundtrip_keeps_flag_json_and_ready_lengths_in_total(tmp_path) -> None:
    import sqlite3

    from core.kp.offers_read import get_kp_by_id
    from core.kp_persistence_service import KpPersistenceService
    from core.pile_trip_pricing import coerce_pile_trip_overrides

    db_path = _seed_catalog(tmp_path)
    tariffs = {"150": {"trip_cost": 5000}, "160": {"trip_cost": 7000}}
    kp_id = KpPersistenceService.save_kp_to_db(
        "23.09.2026",
        _priced_mix(),
        customer_name="Длинные сваи",
        product_type="piles",
        db_path=db_path,
        pile_logistics_cost=1000.0,
        pile_trip_overrides={"С160.35-12": 2},
        long_pile_delivery_enabled=True,
        long_pile_delivery=tariffs,
    )
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        meta = conn.execute(
            """
            SELECT long_pile_delivery_enabled, long_pile_delivery_json, pile_trip_overrides_json
            FROM kp_meta WHERE kp_id = ?
            """,
            (kp_id,),
        ).fetchone()
        total = conn.execute(
            "SELECT total_amount, pile_logistics_cost FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert int(meta["long_pile_delivery_enabled"]) == 1
    stored = coerce_long_pile_delivery(meta["long_pile_delivery_json"])
    assert stored[150] == 5000.0
    assert stored[160] == 7000.0
    assert coerce_pile_trip_overrides(meta["pile_trip_overrides_json"])["С160.35-12"] == 2
    assert total["pile_logistics_cost"] == pytest.approx(1000.0)
    assert total["total_amount"] == pytest.approx(150.0 + 1000.0 + 10000.0 + 14000.0)

    raw = get_kp_by_id(kp_id, db_path)
    assert int(raw["long_pile_delivery_enabled"]) == 1
    assert coerce_long_pile_delivery(raw["long_pile_delivery_json"])[150] == 5000.0


def test_saved_without_flag_keeps_one_pile_boiler(tmp_path) -> None:
    from core.kp.offers_read import get_kp_by_id
    from core.kp_persistence_service import KpPersistenceService

    db_path = _seed_catalog(tmp_path)
    order = [
        _line("С100.35", 6, unit_price=10.0),
        _line("С160.35-12", 4, unit_price=10.0),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "23.09.2026",
        order,
        customer_name="Старый котёл",
        product_type="piles",
        db_path=db_path,
        pile_logistics_cost=1000.0,
    )
    raw = get_kp_by_id(kp_id, db_path)
    assert coerce_long_pile_delivery_enabled(raw.get("long_pile_delivery_enabled")) is False
    totals = calculate_total_cost(
        order,
        discount_percent=0,
        logistics_cost=0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=1000.0,
        pile_catalog_db_path=db_path,
        long_pile_delivery_enabled=coerce_long_pile_delivery_enabled(
            raw.get("long_pile_delivery_enabled")
        ),
        long_pile_delivery=raw.get("long_pile_delivery_json"),
    )
    assert totals["pile_delivery_ready"] is False
    assert totals["pile_delivery_total"] == 0.0
    assert totals["long_pile_delivery_total"] == 0.0
    assert raw["total_amount"] == pytest.approx(100.0)

    split_id = KpPersistenceService.save_kp_to_db(
        "23.09.2026",
        order,
        customer_name="Новый котёл",
        product_type="piles",
        db_path=db_path,
        pile_logistics_cost=1000.0,
        long_pile_delivery_enabled=True,
    )
    split_raw = get_kp_by_id(split_id, db_path)
    assert split_raw["total_amount"] == pytest.approx(100.0 + 1000.0)


def test_patch_length_tariff_changes_total_and_keeps_short_boiler(tmp_path) -> None:
    from core.kp.offers_read import get_kp_by_id
    from core.kp.offers_write import update_kp_logistics_cost
    from core.kp_persistence_service import KpPersistenceService
    from core.pile_trip_pricing import coerce_pile_trip_overrides

    db_path = _seed_catalog(tmp_path)
    kp_id = KpPersistenceService.save_kp_to_db(
        "23.09.2026",
        _priced_mix(),
        customer_name="Патч",
        product_type="piles",
        db_path=db_path,
        pile_logistics_cost=1000.0,
        pile_trip_overrides={"С160.35-12": 2},
        long_pile_delivery_enabled=True,
        long_pile_delivery={"150": {"trip_cost": 5000}, "160": {"trip_cost": 7000}},
    )
    before = get_kp_by_id(kp_id, db_path)
    assert update_kp_logistics_cost(
        kp_id,
        0.0,
        db_path,
        long_pile_delivery={"150": {"trip_cost": 8000}, "160": {"trip_cost": 7000}},
    )
    after = get_kp_by_id(kp_id, db_path)
    assert after["pile_logistics_cost"] == pytest.approx(1000.0)
    assert after["total_amount"] == pytest.approx(before["total_amount"] + 6000.0)
    stored = coerce_long_pile_delivery(after["long_pile_delivery_json"])
    assert stored[150] == 8000.0
    assert stored[160] == 7000.0
    assert coerce_pile_trip_overrides(after["pile_trip_overrides_json"])["С160.35-12"] == 2

    assert update_kp_logistics_cost(kp_id, 0.0, db_path, pile_logistics_cost=1000.0)
    kept = get_kp_by_id(kp_id, db_path)
    assert coerce_long_pile_delivery(kept["long_pile_delivery_json"])[150] == 8000.0


def _xlsx_names(buf) -> str:
    import pandas as pd

    frame = pd.read_excel(buf, sheet_name="КП", header=None)
    return " ".join(str(value) for value in frame.values.ravel() if pd.notna(value))


def _pdf_text(data: bytes) -> str:
    import subprocess
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".pdf") as handle:
        handle.write(data)
        handle.flush()
        result = subprocess.run(
            ["pdftotext", "-layout", handle.name, "-"],
            check=True,
            capture_output=True,
        )
    return result.stdout.decode("utf-8", errors="replace")


def test_xlsx_and_pdf_show_ready_lengths_only(tmp_path) -> None:
    from core.commercial_offer import generate_commercial_offer_pdf
    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    db_path = _seed_catalog(tmp_path)
    order = [
        _line("С100.35", 6, unit_price=10.0),
        _line("С140.35", 4, unit_price=10.0),
        _line("С150.35", 5, unit_price=10.0),
        _line("С160.35-12", 4, unit_price=10.0),
    ]
    common = dict(
        order_data=order,
        offer_number="L1",
        offer_date="23.09.2026",
        customer_name="ООО Тест",
        pile_logistics_cost=1000.0,
        pile_trip_overrides={"С160.35-12": 2},
        pile_catalog_db_path=db_path,
        long_pile_delivery_enabled=True,
        long_pile_delivery={"140": {"trip_cost": 5000}, "160": {"trip_cost": 7000}},
    )
    xlsx_names = _xlsx_names(generate_commercial_offer_xlsx(**common))
    assert "Доставка свай 14,0 м" in xlsx_names
    assert "Доставка свай 16,0 м" in xlsx_names
    assert "15,0 м" not in xlsx_names
    compact = xlsx_names.replace(" ", "").replace("\xa0", "")
    assert "5000" in compact
    assert "7000" in compact
    totals = _money(
        order,
        db_path=db_path,
        enabled=True,
        pile_logistics_cost=1000.0,
        tariffs={"140": {"trip_cost": 5000}, "160": {"trip_cost": 7000}},
        overrides={"С160.35-12": 2},
    )
    shown = [
        row
        for row in totals["long_pile_lengths"]
        if row["ready"] and row["amount"] > 0
    ]
    assert [row["length_key"] for row in shown] == [140, 160]
    assert sum(row["amount"] for row in shown) == totals["long_pile_delivery_total"]

    pdf_text = _pdf_text(generate_commercial_offer_pdf(**common).getvalue())
    assert "Доставка свай 14,0 м" in pdf_text
    assert "Доставка свай 16,0 м" in pdf_text
    assert "Доставка свай 15,0 м" not in pdf_text

    without_flag = dict(common)
    without_flag["long_pile_delivery_enabled"] = False
    plain_names = _xlsx_names(generate_commercial_offer_xlsx(**without_flag))
    assert "14,0 м" not in plain_names
    assert "16,0 м" not in plain_names
    plain_pdf = _pdf_text(generate_commercial_offer_pdf(**without_flag).getvalue())
    assert "Доставка свай 14,0 м" not in plain_pdf
    assert "16,0 м" not in plain_pdf


def test_discount_does_not_change_long_pile_delivery(tmp_path) -> None:
    db_path = _seed_catalog(tmp_path)
    order = [_line("С150.35", 5, unit_price=100.0)]
    totals = _money(
        order,
        db_path=db_path,
        enabled=True,
        tariffs={"150": {"trip_cost": 5000}},
        discount_percent=50,
    )
    assert totals["long_pile_delivery_total"] == 10000.0
    assert totals["total_with_vat"] == 250.0 + 10000.0
