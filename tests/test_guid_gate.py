"""GPS-014 / GPS-028: резолвер GUID для счёта по всем 6 видам."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.guid_gate import (
    InvoiceGuidBlockError,
    OrderLine,
    assert_invoice_guids,
    check_invoice_guids,
)
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.plate_guid_choice import ensure_schema as ensure_choice_schema
from core.plate_guid_choice import set_choice

PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
U_GUID = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
OTHER_GUID = "22222222-2222-2222-2222-222222222222"
PLATE_GUID_A = "aaaaaaaa-1111-1111-1111-111111111111"
PLATE_GUID_B = "bbbbbbbb-1111-1111-1111-111111111111"


def _prays_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS prays_plity (
            "Уникальный идентификатор (Номенклатура)" TEXT,
            "Товар" TEXT
        )
        """
    )


@pytest.fixture
def conn(tmp_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(tmp_path / "pb.db")
    try:
        ensure_guid_schema(connection)
        ensure_choice_schema(connection)
        _prays_schema(connection)
        yield connection
    finally:
        connection.close()


def _seed_unique_plate(conn: sqlite3.Connection, name: str, guid: str) -> None:
    conn.execute(
        'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар")'
        " VALUES (?, ?)",
        (guid, name),
    )


def _seed_duplicate_plate(conn: sqlite3.Connection, name: str) -> None:
    _seed_unique_plate(conn, name, PLATE_GUID_A)
    _seed_unique_plate(conn, name, PLATE_GUID_B)


def test_six_kinds_ready_when_guid_present(conn: sqlite3.Connection) -> None:
    upsert(conn, "pile", "С30.30-3", guid_1c=PLAIN_GUID, match_status="auto")
    upsert(conn, "bridge_pile", "С13-40T7", guid_1c=PLAIN_GUID, match_status="manual")
    upsert(conn, "fbs", "ФБС 9.3.6-Т", guid_1c=PLAIN_GUID, match_status="auto")
    upsert(conn, "stair_flight", "1ЛМ 30-11-15-4", guid_1c=PLAIN_GUID, match_status="auto")
    upsert(conn, "stair_step", "ЛС11", guid_1c=PLAIN_GUID, match_status="auto")
    _seed_unique_plate(conn, "ПБ 51-12-8", OTHER_GUID)

    report = check_invoice_guids(
        [
            OrderLine("pile", "С30.30-3"),
            OrderLine("bridge_pile", "С13-40T7"),
            OrderLine("fbs", "ФБС 9.3.6-Т"),
            OrderLine("stair_flight", "1ЛМ 30-11-15-4"),
            OrderLine("stair_step", "ЛС11"),
            OrderLine("plate", "ПБ 51-12-8"),
        ],
        conn,
    )

    assert report.missing == ()
    assert len(report.ready) == 6
    by_kind = {item.line.product_kind: item.guid for item in report.ready}
    assert by_kind["pile"] == PLAIN_GUID
    assert by_kind["plate"] == OTHER_GUID


def test_u_suffix_without_guid_1c_u_is_missing(conn: sqlite3.Connection) -> None:
    upsert(conn, "pile", "С110.30-9", guid_1c=PLAIN_GUID, match_status="auto")

    report = check_invoice_guids([OrderLine("pile", "С110.30-9у")], conn)

    assert report.ready == ()
    assert len(report.missing) == 1
    item = report.missing[0]
    assert item.line.mark == "С110.30-9у"
    assert "guid_1c_u" in item.reason.lower() or "у" in item.reason.lower()
    assert item.action_hint


def test_u_suffix_with_guid_1c_u_is_ready(conn: sqlite3.Connection) -> None:
    upsert(
        conn,
        "pile",
        "С110.30-9",
        guid_1c=PLAIN_GUID,
        guid_1c_u=U_GUID,
        match_status="auto",
    )

    report = check_invoice_guids([OrderLine("pile", "С110.30-9у")], conn)

    assert report.missing == ()
    assert len(report.ready) == 1
    assert report.ready[0].guid == U_GUID


def test_plate_duplicate_without_choice_is_missing(conn: sqlite3.Connection) -> None:
    _seed_duplicate_plate(conn, "ЛВ60.12-4")

    report = check_invoice_guids([OrderLine("plate", "ЛВ 60.12-4")], conn)

    assert report.ready == ()
    assert len(report.missing) == 1
    assert "дубль" in report.missing[0].reason.lower()
    assert report.missing[0].action_hint


def test_plate_duplicate_with_choice_is_ready(conn: sqlite3.Connection) -> None:
    _seed_duplicate_plate(conn, "ЛВ60.12-4")
    set_choice(
        conn,
        "ЛВ 60.12-4",
        chosen_guid=PLATE_GUID_B,
        chosen_name="ЛВ60.12-4",
    )

    report = check_invoice_guids([OrderLine("plate", "ЛВ60.12-4")], conn)

    assert report.missing == ()
    assert len(report.ready) == 1
    assert report.ready[0].guid == PLATE_GUID_B


def test_custom_mark_without_guid_is_missing(conn: sqlite3.Connection) -> None:
    report = check_invoice_guids([OrderLine("pile", "СГС 80.35")], conn)

    assert report.ready == ()
    assert len(report.missing) == 1
    assert report.missing[0].line.mark == "СГС 80.35"
    assert report.missing[0].action_hint


def test_ambiguous_status_is_missing(conn: sqlite3.Connection) -> None:
    upsert(
        conn,
        "pile",
        "С50.30-8",
        guid_1c=PLAIN_GUID,
        match_status="ambiguous",
        match_note="в 1С несколько GUID",
    )

    report = check_invoice_guids([OrderLine("pile", "С50.30-8")], conn)

    assert report.ready == ()
    assert "дубль" in report.missing[0].reason.lower() or "неоднознач" in report.missing[0].reason.lower()


def test_assert_invoice_guids_raises_with_human_list(conn: sqlite3.Connection) -> None:
    upsert(conn, "fbs", "ФБС 9.3.6-Т", guid_1c=PLAIN_GUID, match_status="auto")

    with pytest.raises(InvoiceGuidBlockError) as caught:
        assert_invoice_guids(
            [
                OrderLine("fbs", "ФБС 9.3.6-Т"),
                OrderLine("pile", "С30.30-3"),
            ],
            conn,
        )

    error = caught.value
    assert error.report.missing
    text = str(error)
    assert "С30.30-3" in text
    assert "завести" in text.lower() or "GUID" in text


@pytest.mark.parametrize(
    ("kind", "mark"),
    [
        ("pile", "С30.30-3"),
        ("bridge_pile", "С13-40T7"),
        ("fbs", "ФБС 9.3.6-Т"),
        ("stair_flight", "1ЛМ 30-11-15-4"),
        ("stair_step", "ЛС11"),
        ("plate", "ПБ 51-12-8"),
    ],
)
def test_assert_blocks_each_of_six_kinds_without_guid(
    conn: sqlite3.Connection, kind: str, mark: str
) -> None:
    with pytest.raises(InvoiceGuidBlockError) as caught:
        assert_invoice_guids([OrderLine(kind, mark)], conn)

    text = str(caught.value)
    assert mark in text
    assert caught.value.report.missing
    assert caught.value.report.ready == ()


def test_assert_blocks_u_suffix_without_guid_1c_u(conn: sqlite3.Connection) -> None:
    upsert(conn, "pile", "С110.30-9", guid_1c=PLAIN_GUID, match_status="auto")

    with pytest.raises(InvoiceGuidBlockError) as caught:
        assert_invoice_guids([OrderLine("pile", "С110.30-9у")], conn)

    text = str(caught.value)
    assert "С110.30-9у" in text
    assert "guid_1c_u" in caught.value.report.missing[0].reason.lower()


def test_order_lines_from_draft_payload() -> None:
    from core.guid_gate import order_lines_from_order_data

    lines = order_lines_from_order_data(
        [
            {"product_type": "piles", "mark": "С110.30-9у", "qty": 2},
            {"product_type": "plates", "name": "ПБ 51-12-8", "qty": 4},
        ]
    )
    assert [(item.product_kind, item.mark, item.qty) for item in lines] == [
        ("pile", "С110.30-9у", 2),
        ("plate", "ПБ 51-12-8", 4),
    ]
