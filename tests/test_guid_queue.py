"""GPS-018: живые очереди 🏭/💰/⚠️ из pb.db."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core.duplicate_candidates import ensure_schema as ensure_dup_schema
from core.duplicate_candidates import upsert_candidates
from core.guid_queue import collect_guid_queue
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.price_queue_db import ensure_schema as ensure_price_schema
from core.price_queue_db import upsert_unmatched

GUID_A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
GUID_B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"


def _db(tmp_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(tmp_path / "pb.db")
    ensure_guid_schema(conn)
    ensure_dup_schema(conn)
    ensure_price_schema(conn)
    return conn


def test_collects_create_price_and_duplicates(tmp_path: Path) -> None:
    conn = _db(tmp_path)
    try:
        upsert(conn, "pile", "С30.30-3", match_status="missing")
        upsert(
            conn,
            "pile",
            "С40.30-6",
            guid_1c=GUID_A,
            match_status="auto",
        )
        upsert_unmatched(
            conn,
            [(GUID_B, "С999.30-1", "pile", "Сваи С 999.30-1")],
        )
        upsert_candidates(
            conn,
            "pile",
            "С50.30-8",
            [(GUID_A, "Сваи С 50.30-8", 1000.0), (GUID_B, "Сваи С 50.30-8", 1000.0)],
        )

        snap = collect_guid_queue(conn)
    finally:
        conn.close()

    create_marks = {(item.mark, item.field) for item in snap.to_create_1c}
    assert ("С30.30-3", "guid_1c") in create_marks
    assert ("С40.30-6", "guid_1c_u") in create_marks
    assert [item.guid for item in snap.to_price] == [GUID_B]
    assert snap.duplicates
    assert snap.duplicates[0].scope == "nonplate"
    assert snap.duplicates[0].key == "С50.30-8"


def test_collect_queue_without_plate_guid_choice_table(tmp_path: Path) -> None:
    """Живой pb.db: prays_plity есть, plate_guid_choice ещё нет — не 500."""
    conn = sqlite3.connect(tmp_path / "pb.db")
    try:
        conn.execute(
            """
            CREATE TABLE prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        conn.execute(
            'INSERT INTO prays_plity VALUES (?, ?)',
            (GUID_A, "ПБ 60-12-8"),
        )
        conn.execute(
            'INSERT INTO prays_plity VALUES (?, ?)',
            (GUID_B, "ПБ 60-12-8"),
        )
        snap = collect_guid_queue(conn)
    finally:
        conn.close()

    plate_dups = [item for item in snap.duplicates if item.scope == "plate"]
    assert plate_dups
    assert plate_dups[0].key == "ПБ 60-12-8"
    assert len(plate_dups[0].candidates) == 2


def test_excel_tasks_sheet_matches_collect_guid_queue(tmp_path: Path) -> None:
    """GPS-018: лист «Задачи» берёт 🏭/💰/⚠️ из того же collect_guid_queue."""
    from openpyxl import Workbook

    from scripts.build_guid_queue import add_live_tasks_sheet

    conn = _db(tmp_path)
    try:
        upsert(conn, "pile", "С30.30-3", match_status="missing")
        upsert(conn, "pile", "С40.30-6", guid_1c=GUID_A, match_status="auto")
        upsert_unmatched(conn, [(GUID_B, "С999.30-1", "pile", "Сваи С 999.30-1")])
        upsert_candidates(
            conn,
            "pile",
            "С50.30-8",
            [(GUID_A, "Сваи С 50.30-8", 1000.0), (GUID_B, "Сваи С 50.30-8", 1000.0)],
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        conn.execute('INSERT INTO prays_plity VALUES (?, ?)', (GUID_A, "ПБ 60-12-8"))
        conn.execute('INSERT INTO prays_plity VALUES (?, ?)', (GUID_B, "ПБ 60-12-8"))
        snap = collect_guid_queue(conn)
        wb = Workbook()
        wb.remove(wb.active)
        create_n, price_n, dup_n = add_live_tasks_sheet(wb, conn)
    finally:
        conn.close()

    assert create_n == len(snap.to_create_1c)
    assert price_n == len(snap.to_price)
    assert dup_n == len(snap.duplicates)
    ws = wb["Задачи"]
    cells = [
        str(cell.value)
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    ]
    blob = " ".join(cells)
    for item in snap.to_create_1c:
        assert item.mark in blob
    for item in snap.to_price:
        assert item.guid in blob
    for item in snap.duplicates:
        assert item.key in blob
    assert "🏭" in blob
    assert "💰" in blob
    assert "⚠️" in blob or "дубл" in blob.lower()
