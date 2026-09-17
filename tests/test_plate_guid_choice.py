"""GPS-011: override-таблица выбора GUID для плит."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.plate_guid_choice import (
    clear_choice,
    ensure_schema,
    get_choice,
    set_choice,
)

GUID_A = "11111111-1111-1111-1111-111111111111"
GUID_B = "22222222-2222-2222-2222-222222222222"
NAME_A = "ЛВ 60.12-4"


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
    rows = conn.execute("PRAGMA table_info(plate_guid_choice)").fetchall()
    conn.row_factory = None
    return {row["name"]: row for row in rows}


def test_get_choice_creates_table_if_missing(conn: sqlite3.Connection) -> None:
    assert get_choice(conn, NAME_A) is None
    names = {
        row[0]
        for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    assert "plate_guid_choice" in names


def test_ensure_schema_is_idempotent(conn: sqlite3.Connection) -> None:
    ensure_schema(conn)
    ensure_schema(conn)

    info = _table_info(conn)
    assert set(info) == {
        "plate_name_norm",
        "chosen_guid",
        "chosen_name",
        "decided_by",
        "decided_at",
        "note",
    }
    assert info["plate_name_norm"]["pk"] == 1
    assert info["chosen_guid"]["notnull"] == 1


def test_set_then_get_returns_choice(db: sqlite3.Connection) -> None:
    saved = set_choice(
        db,
        NAME_A,
        chosen_guid=GUID_A,
        chosen_name="ЛВ60.12-4",
        decided_by="economist",
        note="по решению бухгалтерии",
    )
    assert saved.chosen_guid == GUID_A
    assert saved.chosen_name == "ЛВ60.12-4"
    assert saved.decided_by == "economist"
    assert saved.note == "по решению бухгалтерии"
    assert saved.decided_at

    loaded = get_choice(db, NAME_A)
    assert loaded is not None
    assert loaded.chosen_guid == GUID_A
    assert loaded.plate_name_norm == saved.plate_name_norm


def test_spaced_and_compact_plate_names_share_key(db: sqlite3.Connection) -> None:
    set_choice(db, "ЛВ 60.12-4", chosen_guid=GUID_A, chosen_name="ЛВ60.12-4")

    compact = get_choice(db, "ЛВ60.12-4")
    spaced = get_choice(db, "ЛВ 60.12 - 4")
    assert compact is not None
    assert spaced is not None
    assert compact.chosen_guid == GUID_A
    assert spaced.chosen_guid == GUID_A
    assert compact.plate_name_norm == spaced.plate_name_norm


def test_set_overwrites_previous_choice(db: sqlite3.Connection) -> None:
    set_choice(db, NAME_A, chosen_guid=GUID_A, chosen_name="A")
    set_choice(db, NAME_A, chosen_guid=GUID_B, chosen_name="B", decided_by="admin")

    row = get_choice(db, NAME_A)
    assert row is not None
    assert row.chosen_guid == GUID_B
    assert row.chosen_name == "B"
    assert row.decided_by == "admin"
    count = db.execute("SELECT COUNT(*) FROM plate_guid_choice").fetchone()[0]
    assert count == 1


def test_clear_choice_removes_row(db: sqlite3.Connection) -> None:
    set_choice(db, NAME_A, chosen_guid=GUID_A)
    clear_choice(db, "ЛВ60.12-4")
    assert get_choice(db, NAME_A) is None
