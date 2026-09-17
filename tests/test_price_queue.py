"""GPS-013: персистентная очередь «ввести цену» из unmatched_1c."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.services.nomenclature_import_service import NomenclatureImportService
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.pricelist_1c_parser import PricelistRow
from core.price_queue_db import (
    ensure_schema,
    list_open,
    set_resolved,
)

GUID_NEW = "99999999-9999-9999-9999-999999999999"
GUID_KNOWN = "11111111-1111-1111-1111-111111111111"
MARK_KNOWN = "С30.30-3"


@pytest.fixture
def db_path(tmp_path: Path) -> Path:
    path = tmp_path / "pb.db"
    conn = sqlite3.connect(path)
    try:
        ensure_guid_schema(conn)
        ensure_schema(conn)
        upsert(conn, "pile", MARK_KNOWN, match_status="missing")
        conn.commit()
    finally:
        conn.close()
    return path


def _row(guid: str, name: str, *, idx: int = 3) -> PricelistRow:
    return PricelistRow(
        guid=guid,
        name=name,
        price=None,
        unit="шт",
        source_file="Прайс сваи.xls",
        row_index=idx,
    )


def _import(db_path: Path, rows: list[PricelistRow]):
    service = NomenclatureImportService(db_path)
    return service._sync_and_commit(rows, "pile")


def test_ensure_schema_is_idempotent(tmp_path: Path) -> None:
    conn = sqlite3.connect(tmp_path / "q.db")
    try:
        ensure_schema(conn)
        ensure_schema(conn)
        conn.row_factory = sqlite3.Row
        cols = {row["name"] for row in conn.execute("PRAGMA table_info(price_queue)")}
        conn.row_factory = None
        assert cols == {
            "guid",
            "mark_norm",
            "kind",
            "name",
            "state",
            "resolved_by",
            "resolved_price",
            "resolved_at",
            "first_seen",
        }
    finally:
        conn.close()


def test_unmatched_1c_opens_price_task(db_path: Path) -> None:
    report = _import(
        db_path,
        [
            _row(GUID_KNOWN, "Сваи С 30.30-3"),
            _row(GUID_NEW, "Сваи С 999.30-1", idx=4),
        ],
    )
    assert [item.guid for item in report.unmatched_1c] == [GUID_NEW]

    conn = sqlite3.connect(db_path)
    try:
        ensure_schema(conn)
        open_items = list_open(conn)
    finally:
        conn.close()

    assert len(open_items) == 1
    item = open_items[0]
    assert item.guid == GUID_NEW
    assert item.kind == "pile"
    assert item.name == "Сваи С 999.30-1"
    assert item.state == "open"
    assert item.mark_norm


def test_repeat_import_does_not_duplicate(db_path: Path) -> None:
    rows = [_row(GUID_NEW, "Сваи С 999.30-1")]
    _import(db_path, rows)
    _import(db_path, rows)

    conn = sqlite3.connect(db_path)
    try:
        ensure_schema(conn)
        open_items = list_open(conn)
        total = conn.execute("SELECT COUNT(*) FROM price_queue").fetchone()[0]
    finally:
        conn.close()

    assert total == 1
    assert len(open_items) == 1
    assert open_items[0].guid == GUID_NEW


def test_resolved_does_not_reopen(db_path: Path) -> None:
    _import(db_path, [_row(GUID_NEW, "Сваи С 999.30-1")])
    conn = sqlite3.connect(db_path)
    try:
        ensure_schema(conn)
        set_resolved(conn, GUID_NEW, resolved_by="economist", resolved_price=1500.0)
        conn.commit()
        first_seen = conn.execute(
            "SELECT first_seen FROM price_queue WHERE guid = ?",
            (GUID_NEW,),
        ).fetchone()[0]
    finally:
        conn.close()

    _import(db_path, [_row(GUID_NEW, "Сваи С 999.30-1 обновлённая")])

    conn = sqlite3.connect(db_path)
    try:
        ensure_schema(conn)
        assert list_open(conn) == []
        row = conn.execute(
            "SELECT state, resolved_by, resolved_price, first_seen, name FROM price_queue"
            " WHERE guid = ?",
            (GUID_NEW,),
        ).fetchone()
    finally:
        conn.close()

    assert row is not None
    state, resolved_by, resolved_price, seen, name = row
    assert state == "resolved"
    assert resolved_by == "economist"
    assert resolved_price == 1500.0
    assert seen == first_seen
    assert name == "Сваи С 999.30-1 обновлённая"


def test_resolve_price_writes_pricelist_and_guid(db_path: Path) -> None:
    from app.services.nomenclature_queue_service import NomenclatureQueueService
    from core.nomenclature_guid import get_by_mark, get_guid_for_invoice
    from core.pile_price_db import get_pile_price

    _import(db_path, [_row(GUID_NEW, "Сваи С 999.30-1")])
    service = NomenclatureQueueService(db_path)
    result = service.resolve_price(GUID_NEW, 2500.0, resolved_by="economist")
    assert result.state == "resolved"
    assert get_pile_price("С999.30-1", "B25", str(db_path)) == 2500.0
    conn = sqlite3.connect(db_path)
    try:
        assert get_guid_for_invoice(conn, "pile", "С999.30-1") == GUID_NEW
        stored = get_by_mark(conn, "pile", "С999.30-1")
        assert list_open(conn) == []
    finally:
        conn.close()
    assert stored is not None
    assert stored.match_status == "auto"
    assert stored.guid_1c == GUID_NEW

    again = service.resolve_price(GUID_NEW, 2500.0, resolved_by="economist")
    assert again.state == "resolved"


def test_resolve_price_rejects_non_positive(db_path: Path) -> None:
    from app.services.nomenclature_queue_service import NomenclatureQueueError, NomenclatureQueueService

    _import(db_path, [_row(GUID_NEW, "Сваи С 999.30-1")])
    service = NomenclatureQueueService(db_path)
    try:
        service.resolve_price(GUID_NEW, 0, resolved_by="economist")
        raised = False
    except NomenclatureQueueError as exc:
        raised = True
        assert exc.status_code == 400
        assert "цена" in exc.message
    assert raised
