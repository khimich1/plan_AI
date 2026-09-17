"""GPS-012: персистентные кандидаты дублей GUID не-плит."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.duplicate_candidates import (
    clear_candidates,
    ensure_schema,
    list_candidates,
)
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.nomenclature_sync import sync_pricelist
from core.pricelist_1c_parser import PricelistRow

GUID_A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
GUID_B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
MARK = "С50.30-8"


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    try:
        yield connection
    finally:
        connection.close()


@pytest.fixture
def db(conn: sqlite3.Connection) -> sqlite3.Connection:
    ensure_guid_schema(conn)
    ensure_schema(conn)
    return conn


def _row(guid: str, name: str, *, price: float | None, idx: int) -> PricelistRow:
    return PricelistRow(
        guid=guid,
        name=name,
        price=price,
        unit="шт",
        source_file="Прайс сваи.xls",
        row_index=idx,
    )


def test_ensure_schema_is_idempotent(conn: sqlite3.Connection) -> None:
    ensure_schema(conn)
    ensure_schema(conn)
    conn.row_factory = sqlite3.Row
    cols = {row["name"] for row in conn.execute("PRAGMA table_info(duplicate_candidates)")}
    conn.row_factory = None
    assert cols == {
        "kind",
        "mark_norm",
        "guid",
        "name",
        "price",
        "first_seen",
        "active",
    }


def test_sync_ambiguous_writes_candidates(db: sqlite3.Connection) -> None:
    upsert(db, "pile", MARK, match_status="missing")

    report = sync_pricelist(
        db,
        [
            _row(GUID_A, "Сваи С 50.30-8", price=1200.0, idx=3),
            _row(GUID_B, "Сваи C 50.30-8", price=1200.0, idx=4),
        ],
        product_kind="pile",
    )

    assert len(report.ambiguous) == 1
    found = list_candidates(db, kind="pile", mark=MARK)
    assert {item.guid for item in found} == {GUID_A, GUID_B}
    by_guid = {item.guid: item for item in found}
    assert by_guid[GUID_A].name == "Сваи С 50.30-8"
    assert by_guid[GUID_A].price == 1200.0
    assert by_guid[GUID_B].name == "Сваи C 50.30-8"
    assert all(item.active for item in found)
    assert all(item.kind == "pile" for item in found)


def test_second_sync_is_idempotent(db: sqlite3.Connection) -> None:
    upsert(db, "pile", MARK, match_status="missing")
    rows = [
        _row(GUID_A, "Сваи С 50.30-8", price=1100.0, idx=3),
        _row(GUID_B, "Сваи С 50.30-8", price=1300.0, idx=4),
    ]
    first = sync_pricelist(db, rows, product_kind="pile")
    first_seen = {item.guid: item.first_seen for item in list_candidates(db)}

    second = sync_pricelist(
        db,
        [
            _row(GUID_A, "Сваи С 50.30-8", price=1500.0, idx=3),
            _row(GUID_B, "Сваи С 50.30-8", price=1600.0, idx=4),
        ],
        product_kind="pile",
    )

    assert first.ambiguous and second.ambiguous
    found = list_candidates(db, kind="pile", mark=MARK)
    assert len(found) == 2
    by_guid = {item.guid: item for item in found}
    assert by_guid[GUID_A].price == 1500.0
    assert by_guid[GUID_A].first_seen == first_seen[GUID_A]
    assert by_guid[GUID_B].first_seen == first_seen[GUID_B]


def test_clear_candidates_hides_rows(db: sqlite3.Connection) -> None:
    upsert(db, "pile", MARK, match_status="missing")
    sync_pricelist(
        db,
        [
            _row(GUID_A, "Сваи С 50.30-8", price=1.0, idx=3),
            _row(GUID_B, "Сваи С 50.30-8", price=2.0, idx=4),
        ],
        product_kind="pile",
    )
    clear_candidates(db, "pile", MARK)
    assert list_candidates(db, kind="pile", mark=MARK) == []
