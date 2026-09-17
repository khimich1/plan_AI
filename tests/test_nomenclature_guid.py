"""GPS-002: nomenclature_guid table and CRUD helpers."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.nomenclature_guid import (
    ensure_schema,
    get_by_mark,
    get_guid_for_invoice,
    iter_by_state,
    set_status,
    upsert,
)

EXPECTED_COLUMNS = {
    "product_kind": (1, 1),  # notnull, pk position
    "mark": (1, 2),
    "guid_1c": (0, 0),
    "guid_1c_u": (0, 0),
    "match_status": (1, 0),
    "match_note": (0, 0),
    "updated_at": (1, 0),
}

PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
REINFORCED_GUID = "22222222-2222-2222-2222-222222222222"
UPDATED_GUID = "33333333-3333-3333-3333-333333333333"


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


def _table_info(conn: sqlite3.Connection) -> dict[str, sqlite3.Row]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute("PRAGMA table_info(nomenclature_guid)").fetchall()
    conn.row_factory = None
    return {row["name"]: row for row in rows}


def test_ensure_schema_creates_table(conn: sqlite3.Connection) -> None:
    ensure_schema(conn)
    ensure_schema(conn)

    info = _table_info(conn)
    assert set(info) == set(EXPECTED_COLUMNS)
    for name, (notnull, pk) in EXPECTED_COLUMNS.items():
        assert info[name]["type"] == "TEXT"
        assert info[name]["notnull"] == notnull
        assert info[name]["pk"] == pk

    indexes = conn.execute("PRAGMA index_list(nomenclature_guid)").fetchall()
    pk_indexes = [row for row in indexes if row[2] == 1]
    assert pk_indexes


def test_upsert_insert_and_idempotent_guid_update(db: sqlite3.Connection) -> None:
    first = upsert(
        db,
        "pile",
        "С60.30-6",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )
    assert first.product_kind == "pile"
    assert first.mark == "С60.30-6"
    assert first.guid_1c == PLAIN_GUID
    assert first.guid_1c_u is None
    assert first.match_status == "auto"
    assert first.updated_at

    second = upsert(
        db,
        "pile",
        "С60.30-6",
        guid_1c=UPDATED_GUID,
        match_status="auto",
    )
    assert second.guid_1c == UPDATED_GUID
    assert second.guid_1c_u is None
    assert second.match_status == "auto"

    count = db.execute("SELECT COUNT(*) FROM nomenclature_guid").fetchone()[0]
    assert count == 1


def test_get_by_mark(db: sqlite3.Connection) -> None:
    upsert(db, "fbs", "ФБС 9.3.6-Т", guid_1c=PLAIN_GUID, match_status="auto")

    row = get_by_mark(db, "fbs", "ФБС 9.3.6-Т")
    assert row is not None
    assert row.guid_1c == PLAIN_GUID
    assert get_by_mark(db, "fbs", "ФБС 24.4.6-Т") is None


def test_set_status(db: sqlite3.Connection) -> None:
    upsert(db, "bridge_pile", "С13-40T7", match_status="missing")

    updated = set_status(db, "bridge_pile", "С13-40T7", "ambiguous", note="duplicate in 1C")
    assert updated.match_status == "ambiguous"
    assert updated.match_note == "duplicate in 1C"

    kept_note = set_status(db, "bridge_pile", "С13-40T7", "manual")
    assert kept_note.match_status == "manual"
    assert kept_note.match_note == "duplicate in 1C"


def test_iter_by_state(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С60.30-6", match_status="auto", guid_1c=PLAIN_GUID)
    upsert(db, "fbs", "ФБС 9.3.6-Т", match_status="missing")
    upsert(db, "stair_step", "ЛС11", match_status="auto", guid_1c=UPDATED_GUID)
    upsert(db, "stair_flight", "1ЛМ 30-11-15-4", match_status="ambiguous")

    auto_rows = list(iter_by_state(db, "auto"))
    assert [row.mark for row in auto_rows] == ["С60.30-6", "ЛС11"]
    assert list(iter_by_state(db, "manual")) == []
    missing = list(iter_by_state(db, "missing"))
    assert len(missing) == 1
    assert missing[0].product_kind == "fbs"


def test_guid_1c_vs_guid_1c_u_for_piles(db: sqlite3.Connection) -> None:
    upsert(
        db,
        "pile",
        "С60.30-6",
        guid_1c=PLAIN_GUID,
        guid_1c_u=REINFORCED_GUID,
        match_status="auto",
    )
    upsert(db, "fbs", "ФБС 9.3.6-Т", guid_1c=UPDATED_GUID, match_status="auto")

    pile = get_by_mark(db, "pile", "С60.30-6")
    assert pile is not None
    assert pile.guid_1c == PLAIN_GUID
    assert pile.guid_1c_u == REINFORCED_GUID
    assert get_guid_for_invoice(db, "pile", "С60.30-6", reinforced=False) == PLAIN_GUID
    assert get_guid_for_invoice(db, "pile", "С60.30-6", reinforced=True) == REINFORCED_GUID

    fbs = get_by_mark(db, "fbs", "ФБС 9.3.6-Т")
    assert fbs is not None
    assert fbs.guid_1c_u is None
    assert get_guid_for_invoice(db, "fbs", "ФБС 9.3.6-Т", reinforced=False) == UPDATED_GUID
    assert get_guid_for_invoice(db, "fbs", "ФБС 9.3.6-Т", reinforced=True) is None


def test_primary_key_uniqueness(db: sqlite3.Connection) -> None:
    upsert(db, "pile", "С80.35-8", guid_1c=PLAIN_GUID, match_status="auto")
    upsert(db, "pile", "С80.35-8", guid_1c=UPDATED_GUID, match_status="manual")
    upsert(db, "bridge_pile", "С80.35-8", guid_1c=REINFORCED_GUID, match_status="auto")

    pile_count = db.execute(
        "SELECT COUNT(*) FROM nomenclature_guid WHERE product_kind = ? AND mark = ?",
        ("pile", "С80.35-8"),
    ).fetchone()[0]
    assert pile_count == 1
    assert get_by_mark(db, "pile", "С80.35-8").guid_1c == UPDATED_GUID
    assert get_by_mark(db, "bridge_pile", "С80.35-8").guid_1c == REINFORCED_GUID

    with pytest.raises(sqlite3.IntegrityError):
        db.execute(
            """
            INSERT INTO nomenclature_guid (
                product_kind, mark, guid_1c, guid_1c_u,
                match_status, match_note, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            ("pile", "С80.35-8", PLAIN_GUID, None, "auto", None, "2026-09-15T00:00:00Z"),
        )
