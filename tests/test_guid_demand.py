"""Очередь спроса GUID: таблица guid_demand, record, sweep, list."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.guid_demand import (
    FIELD_DUPLICATE,
    FIELD_GUID,
    FIELD_GUID_U,
    STATE_OPEN,
    STATE_RESOLVED,
    ensure_schema,
    list_open,
    record_guid_demand,
    sweep_guid_demand,
)
from core.guid_gate import (
    HINT_CHOOSE,
    HINT_CREATE,
    HINT_CREATE_U,
    MissingItem,
    OrderLine,
    REASON_AMBIGUOUS,
    REASON_DUP,
    REASON_MISSING,
    REASON_MISSING_U,
)
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.plate_guid_choice import ensure_schema as ensure_choice_schema

PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
OTHER_GUID = "22222222-2222-2222-2222-222222222222"

EXPECTED_COLUMNS = {
    "product_kind": (1, 1),
    "mark": (1, 2),
    "field": (1, 3),
    "reason": (1, 0),
    "hint": (1, 0),
    "state": (1, 0),
    "kp_ids_json": (1, 0),
    "opened_at": (1, 0),
    "resolved_at": (0, 0),
}


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    try:
        ensure_guid_schema(connection)
        ensure_choice_schema(connection)
        yield connection
    finally:
        connection.close()


def _table_info(conn: sqlite3.Connection) -> dict[str, sqlite3.Row]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute("PRAGMA table_info(guid_demand)").fetchall()
    conn.row_factory = None
    return {row["name"]: row for row in rows}


def _missing(
    kind: str,
    mark: str,
    *,
    reason: str = REASON_MISSING,
    hint: str = HINT_CREATE,
) -> MissingItem:
    return MissingItem(line=OrderLine(kind, mark), reason=reason, action_hint=hint)


def test_ensure_schema_is_idempotent(conn: sqlite3.Connection) -> None:
    ensure_schema(conn)
    ensure_schema(conn)

    info = _table_info(conn)
    assert set(info) == set(EXPECTED_COLUMNS)
    for name, (notnull, pk) in EXPECTED_COLUMNS.items():
        assert info[name]["type"] == "TEXT"
        assert info[name]["notnull"] == notnull
        assert info[name]["pk"] == pk


def test_record_two_kps_merges_kp_ids(conn: sqlite3.Connection) -> None:
    item = _missing("pile", "С70.35-9у", reason=REASON_MISSING_U, hint=HINT_CREATE_U)
    record_guid_demand(conn, 12, [item])
    record_guid_demand(conn, 18, [item])

    rows = list_open(conn)
    assert len(rows) == 1
    row = rows[0]
    assert row.product_kind == "pile"
    assert row.mark == "С70.35-9у"
    assert row.field == FIELD_GUID_U
    assert row.state == STATE_OPEN
    assert list(row.kp_ids) == [12, 18]
    assert row.resolved_at is None


def test_same_kp_does_not_duplicate_kp_id(conn: sqlite3.Connection) -> None:
    item = _missing("pile", "С30.30-3")
    record_guid_demand(conn, 7, [item])
    record_guid_demand(conn, 7, [item])

    rows = list_open(conn)
    assert len(rows) == 1
    assert list(rows[0].kp_ids) == [7]


def test_empty_missing_is_noop(conn: sqlite3.Connection) -> None:
    record_guid_demand(conn, 1, [])
    assert list_open(conn) == []


def test_unknown_mark_still_records_create(conn: sqlite3.Connection) -> None:
    item = _missing("pile", "С99.99-1")
    record_guid_demand(conn, 5, [item])
    rows = list_open(conn)
    assert len(rows) == 1
    assert rows[0].mark == "С99.99-1"
    assert rows[0].field == FIELD_GUID
    assert list(rows[0].kp_ids) == [5]


def test_reopen_after_resolved(conn: sqlite3.Connection) -> None:
    item = _missing("pile", "С30.30-3")
    record_guid_demand(conn, 12, [item])
    row = list_open(conn)[0]
    conn.execute(
        "UPDATE guid_demand SET state = ?, resolved_at = ? "
        "WHERE product_kind = ? AND mark = ? AND field = ?",
        (STATE_RESOLVED, "2026-09-01T00:00:00Z", row.product_kind, row.mark, row.field),
    )
    assert list_open(conn) == []

    record_guid_demand(conn, 12, [item])
    opened = list_open(conn)
    assert len(opened) == 1
    assert opened[0].state == STATE_OPEN
    assert opened[0].resolved_at is None
    assert list(opened[0].kp_ids) == [12]


def test_reason_maps_dup_and_ambiguous_to_duplicate(conn: sqlite3.Connection) -> None:
    record_guid_demand(
        conn,
        4,
        [
            _missing("pile", "С50.30-8", reason=REASON_AMBIGUOUS, hint=HINT_CHOOSE),
            _missing("plate", "ЛВ60.12-4", reason=REASON_DUP, hint=HINT_CHOOSE),
        ],
    )
    rows = {(row.product_kind, row.mark, row.field) for row in list_open(conn)}
    assert rows == {
        ("pile", "С50.30-8", FIELD_DUPLICATE),
        ("plate", "ЛВ60.12-4", FIELD_DUPLICATE),
    }


def test_sweep_closes_create_after_guid_upsert(conn: sqlite3.Connection) -> None:
    record_guid_demand(conn, 12, [_missing("pile", "С30.30-3")])
    upsert(
        conn,
        "pile",
        "С30.30-3",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )

    sweep_guid_demand(conn)

    assert list_open(conn) == []
    stored = conn.execute(
        "SELECT state, resolved_at FROM guid_demand "
        "WHERE product_kind = ? AND mark = ? AND field = ?",
        ("pile", "С30.30-3", FIELD_GUID),
    ).fetchone()
    assert stored is not None
    assert stored[0] == STATE_RESOLVED
    assert stored[1]


def test_sweep_shifts_create_to_duplicate_without_updating_pk(
    conn: sqlite3.Connection,
) -> None:
    record_guid_demand(conn, 12, [_missing("pile", "С50.30-8")])
    record_guid_demand(conn, 18, [_missing("pile", "С50.30-8")])
    upsert(conn, "pile", "С50.30-8", match_status="ambiguous")

    sweep_guid_demand(conn)

    opened = list_open(conn)
    assert len(opened) == 1
    assert opened[0].field == FIELD_DUPLICATE
    assert opened[0].state == STATE_OPEN
    assert list(opened[0].kp_ids) == [12, 18]

    old = conn.execute(
        "SELECT state FROM guid_demand "
        "WHERE product_kind = ? AND mark = ? AND field = ?",
        ("pile", "С50.30-8", FIELD_GUID),
    ).fetchone()
    assert old is not None
    assert old[0] == STATE_RESOLVED


def test_sweep_leaves_unrelated_open_mark(conn: sqlite3.Connection) -> None:
    upsert(
        conn,
        "pile",
        "С70.35-9",
        guid_1c=OTHER_GUID,
        match_status="auto",
    )
    record_guid_demand(conn, 1, [_missing("pile", "С30.30-3")])
    record_guid_demand(
        conn,
        2,
        [_missing("pile", "С70.35-9у", reason=REASON_MISSING_U, hint=HINT_CREATE_U)],
    )
    upsert(
        conn,
        "pile",
        "С30.30-3",
        guid_1c=PLAIN_GUID,
        match_status="auto",
    )

    sweep_guid_demand(conn)

    opened = list_open(conn)
    assert len(opened) == 1
    assert opened[0].mark == "С70.35-9у"
    assert opened[0].field == FIELD_GUID_U
    assert opened[0].state == STATE_OPEN
