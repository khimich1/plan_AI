"""GPS-005: sync 1C pricelist rows into nomenclature_guid."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.nomenclature_guid import ensure_schema, get_by_mark, upsert
from core.nomenclature_sync import (
    FIELD_PLAIN,
    FIELD_U,
    infer_product_kind,
    normalize_name,
    sync_pricelist,
)
from core.pricelist_1c_parser import PricelistRow

PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
OTHER_GUID = "22222222-2222-2222-2222-222222222222"
THIRD_GUID = "33333333-3333-3333-3333-333333333333"
U_GUID = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    try:
        yield connection
    finally:
        connection.close()


@pytest.fixture
def db(conn: sqlite3.Connection) -> sqlite3.Connection:
    ensure_schema(conn)
    return conn


def _row(
    guid: str,
    name: str,
    *,
    source: str = "Прайс сваи.xls",
    idx: int = 3,
    price: float | None = None,
) -> PricelistRow:
    return PricelistRow(
        guid=guid,
        name=name,
        price=price,
        unit="шт",
        source_file=source,
        row_index=idx,
    )


def _count(db: sqlite3.Connection) -> int:
    return db.execute("SELECT COUNT(*) FROM nomenclature_guid").fetchone()[0]


def test_unique_match_fills_guid_1c_status_auto(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С30.30-3", match_status="missing")

    report = sync_pricelist(
        db,
        [_row(PLAIN_GUID, "Сваи С 30.30-3")],
        product_kind="pile",
    )

    assert len(report.new_guids) == 1
    assert report.new_guids[0].mark == "С30.30-3"
    assert report.new_guids[0].field == FIELD_PLAIN
    assert report.new_guids[0].guid == PLAIN_GUID
    row = get_by_mark(db, "pile", "С30.30-3")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.guid_1c_u is None
    assert row.match_status == "auto"


def test_second_run_adds_zero_guids(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С30.30-3", match_status="missing")
    rows = [_row(PLAIN_GUID, "Сваи С 30.30-3")]
    first = sync_pricelist(db, rows, product_kind="pile")
    assert len(first.new_guids) == 1

    second = sync_pricelist(db, rows, product_kind="pile")
    assert second.new_guids == ()
    assert second.updated_guids == ()
    assert second.unchanged == 1
    row = get_by_mark(db, "pile", "С30.30-3")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.match_status == "auto"


def test_reinforced_u_fills_guid_1c_u_of_plain_mark(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С40.30-6", match_status="missing")

    report = sync_pricelist(
        db,
        [_row(U_GUID, "Сваи С 40.30-6у")],
        product_kind="pile",
    )

    assert len(report.new_guids) == 1
    assert report.new_guids[0].field == FIELD_U
    row = get_by_mark(db, "pile", "С40.30-6")
    assert row is not None
    assert row.guid_1c is None
    assert row.guid_1c_u == U_GUID
    assert row.match_status == "auto"
    assert get_by_mark(db, "pile", "С40.30-6у") is None
    assert _count(db) == 1


def test_two_guids_same_name_ambiguous_no_auto_pick(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С50.30-8",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )

    report = sync_pricelist(
        db,
        [
            _row(OTHER_GUID, "Сваи С 50.30-8", idx=3),
            _row(THIRD_GUID, "Сваи C 50.30-8", idx=4),
        ],
        product_kind="pile",
    )

    assert report.new_guids == ()
    assert len(report.ambiguous) == 1
    assert report.ambiguous[0].mark == "С50.30-8"
    assert set(report.ambiguous[0].guids) == {OTHER_GUID, THIRD_GUID}
    row = get_by_mark(db, "pile", "С50.30-8")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.match_status == "ambiguous"
    assert OTHER_GUID in (row.match_note or "")
    assert THIRD_GUID in (row.match_note or "")


def test_plates_unknown_kind_not_inserted(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С30.30-3", match_status="missing")
    plate_rows = [
        _row(
            PLAIN_GUID,
            "Плиты ПБ 51-12-8",
            source="Прайс плиты.xls",
        )
    ]

    skipped = sync_pricelist(db, plate_rows, product_kind="plates")
    inferred = sync_pricelist(db, plate_rows)

    assert _count(db) == 1
    assert get_by_mark(db, "pile", "С30.30-3") is not None
    kinds = {
        row[0]
        for row in db.execute("SELECT DISTINCT product_kind FROM nomenclature_guid")
    }
    assert kinds == {"pile"}
    assert len(skipped.unmatched_1c) == 1
    assert len(inferred.unmatched_1c) == 1
    assert infer_product_kind("Прайс плиты.xls") is None


def test_latin_c_matches_cyrillic_mark(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С60.30-6", match_status="missing")

    report = sync_pricelist(
        db,
        [_row(PLAIN_GUID, "Сваи C 60.30-6")],
        product_kind="pile",
    )

    assert normalize_name("Сваи C 60.30-6") == normalize_name("Сваи С 60.30-6")
    assert len(report.new_guids) == 1
    row = get_by_mark(db, "pile", "С60.30-6")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.match_status == "auto"


def test_disappeared_seeded_row_not_in_file(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С70.30-8",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )
    upsert(db, "pile", "С30.30-3", match_status="missing")

    report = sync_pricelist(
        db,
        [_row(OTHER_GUID, "Сваи С 30.30-3")],
        product_kind="pile",
    )

    assert [item.mark for item in report.disappeared] == ["С70.30-8"]
    leftover = get_by_mark(db, "pile", "С70.30-8")
    assert leftover is not None
    assert leftover.guid_1c == PLAIN_GUID
    assert leftover.match_status == "auto"
    filled = get_by_mark(db, "pile", "С30.30-3")
    assert filled is not None
    assert filled.guid_1c == OTHER_GUID


def test_plain_and_u_together_then_idempotent(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С80.30-10", match_status="missing")
    rows = [
        _row(PLAIN_GUID, "Сваи С 80.30-10", idx=3),
        _row(U_GUID, "Сваи С 80.30-10у", idx=4),
    ]
    first = sync_pricelist(db, rows, product_kind="pile")
    assert {item.field for item in first.new_guids} == {FIELD_PLAIN, FIELD_U}

    second = sync_pricelist(db, rows, product_kind="pile")
    assert second.new_guids == ()
    assert second.unchanged == 2
    row = get_by_mark(db, "pile", "С80.30-10")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.guid_1c_u == U_GUID


def test_manual_guid_not_overwritten(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С90.30-12",
        guid_1c=PLAIN_GUID,
        match_status="manual",
        match_note="выбрала бухгалтерия",
    )

    report = sync_pricelist(
        db,
        [_row(OTHER_GUID, "Сваи С 90.30-12")],
        product_kind="pile",
    )

    assert report.new_guids == ()
    assert report.updated_guids == ()
    row = get_by_mark(db, "pile", "С90.30-12")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert row.match_status == "manual"
    assert row.match_note == "выбрала бухгалтерия"


def test_auto_unique_guid_change_updates(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С100.30-12",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )

    report = sync_pricelist(
        db,
        [_row(OTHER_GUID, "Сваи С 100.30-12")],
        product_kind="pile",
    )

    assert report.new_guids == ()
    assert len(report.updated_guids) == 1
    row = get_by_mark(db, "pile", "С100.30-12")
    assert row is not None
    assert row.guid_1c == OTHER_GUID
    assert row.match_status == "auto"


def test_fbs_and_bridge_and_stairs_name_mapping(db: sqlite3.Connection) -> None:
    upsert(db, "fbs", "ФБС 9.3.6-Т", match_status="missing")
    upsert(db, "bridge_pile", "С13-40T7", match_status="missing")
    upsert(db, "stair_flight", "1ЛМ 30-11-15-4", match_status="missing")
    upsert(db, "stair_step", "ЛС11", match_status="missing")

    fbs = sync_pricelist(
        db,
        [_row(PLAIN_GUID, "Блоки ФБС 9.3.6-Т", source="Прайс блоки.xls")],
        product_kind="fbs",
    )
    bridge = sync_pricelist(
        db,
        [
            _row(
                OTHER_GUID,
                "Сваи мостовые С 13.40-Т7",
                source="Прайс сваи мостовые.xls",
            )
        ],
        product_kind="bridge_pile",
    )
    march = sync_pricelist(
        db,
        [_row(THIRD_GUID, "Лестничные марши 1ЛМ 30-11-15-4", source="lm.xls")],
        product_kind="stair_flight",
    )
    step = sync_pricelist(
        db,
        [_row(U_GUID, "Лестничные ступени ЛС-11", source="lm.xls")],
        product_kind="stair_step",
    )

    assert len(fbs.new_guids) == 1
    assert len(bridge.new_guids) == 1
    assert len(march.new_guids) == 1
    assert len(step.new_guids) == 1
    assert get_by_mark(db, "fbs", "ФБС 9.3.6-Т").guid_1c == PLAIN_GUID
    assert get_by_mark(db, "bridge_pile", "С13-40T7").guid_1c == OTHER_GUID
    assert get_by_mark(db, "stair_flight", "1ЛМ 30-11-15-4").guid_1c == THIRD_GUID
    assert get_by_mark(db, "stair_step", "ЛС11").guid_1c == U_GUID


def test_unmatched_1c_not_inserted(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С30.30-3", match_status="missing")

    report = sync_pricelist(
        db,
        [_row(PLAIN_GUID, "Сваи С 30.30-3"), _row(OTHER_GUID, "Сваи С 999.30-1")],
        product_kind="pile",
    )

    assert [item.name for item in report.unmatched_1c] == ["Сваи С 999.30-1"]
    assert _count(db) == 1
    assert get_by_mark(db, "pile", "С999.30-1") is None


def test_unknown_product_kind_raises(db: sqlite3.Connection) -> None:
    with pytest.raises(ValueError, match="Unknown product_kind"):
        sync_pricelist(db, [_row(PLAIN_GUID, "X")], product_kind="widget")
