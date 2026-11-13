"""D-plate-load: запись карточки плиты в prays_plity по GUID."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass

from core.plate_card_import import apply_plate_cards

PLATE = "Плиты ПБ 13-2,6-8п"
GUID = "77f52518-df3a-11f1-900b-04d9f5387845"
OTHER = "Плиты ПБ 17-2,6-10п"
OTHER_GUID = "11111111-1111-1111-1111-111111111111"
SECOND_GUID = "22222222-2222-2222-2222-222222222222"


@dataclass(frozen=True)
class _Row:
    name: str
    guid: str


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(":memory:")
    conn.execute(
        'CREATE TABLE prays_plity ('
        '"Уникальный идентификатор (Номенклатура)" TEXT, '
        '"Товар" TEXT)'
    )
    return conn


def _rows(conn: sqlite3.Connection) -> list[tuple[str, str]]:
    return conn.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        "FROM prays_plity ORDER BY 1"
    ).fetchall()


def test_new_guid_inserts_one_row_and_keeps_neighbor() -> None:
    conn = _connect()
    conn.execute(
        'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
        "VALUES (?, ?)",
        (OTHER_GUID, OTHER),
    )

    report = apply_plate_cards(conn, [_Row(PLATE, GUID)])

    assert _rows(conn) == [(OTHER_GUID, OTHER), (GUID, PLATE)]
    assert [(item.guid, item.name) for item in report.new_guids] == [(GUID, PLATE)]
    assert report.updated_guids == ()
    assert report.unchanged == 0
    conn.close()


def test_same_guid_and_name_is_unchanged() -> None:
    conn = _connect()
    conn.execute(
        'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
        "VALUES (?, ?)",
        (GUID, PLATE),
    )

    report = apply_plate_cards(conn, [_Row(PLATE, "{" + GUID.upper() + "}")])

    assert _rows(conn) == [(GUID, PLATE)]
    assert report.new_guids == ()
    assert report.updated_guids == ()
    assert report.unchanged == 1
    conn.close()


def test_same_guid_new_name_updates_row() -> None:
    conn = _connect()
    conn.execute(
        'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
        "VALUES (?, ?)",
        (GUID, "Плиты ПБ 13-2,6-8"),
    )

    report = apply_plate_cards(conn, [_Row(PLATE, GUID)])

    assert _rows(conn) == [(GUID, PLATE)]
    assert [(item.guid, item.name) for item in report.updated_guids] == [(GUID, PLATE)]
    assert report.new_guids == ()
    assert report.unchanged == 0
    conn.close()


def test_second_guid_on_same_name_inserts_second_row() -> None:
    conn = _connect()
    conn.execute(
        'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
        "VALUES (?, ?)",
        (GUID, PLATE),
    )

    report = apply_plate_cards(conn, [_Row(PLATE, SECOND_GUID)])

    assert _rows(conn) == [(SECOND_GUID, PLATE), (GUID, PLATE)]
    assert [(item.guid, item.name) for item in report.new_guids] == [(SECOND_GUID, PLATE)]
    conn.close()


def test_blank_guid_or_name_is_skipped() -> None:
    conn = _connect()

    report = apply_plate_cards(
        conn,
        [_Row("", GUID), _Row(PLATE, "  "), _Row("  ", "")],
    )

    assert _rows(conn) == []
    assert report.new_guids == ()
    assert report.updated_guids == ()
    assert report.unchanged == 0
    conn.close()


def test_repeated_guid_in_one_file_keeps_last_name() -> None:
    conn = _connect()

    report = apply_plate_cards(
        conn,
        [_Row("Плиты ПБ 13-2,6-6п", GUID), _Row(PLATE, GUID)],
    )

    assert _rows(conn) == [(GUID, PLATE)]
    assert [(item.guid, item.name) for item in report.new_guids] == [(GUID, PLATE)]
    assert report.updated_guids == ()
    conn.close()
