"""GPS-020: резолв дублей GUID не-плит и плит."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from app.services.nomenclature_queue_service import (
    NomenclatureQueueError,
    NomenclatureQueueService,
)
from core.duplicate_candidates import (
    ensure_schema as ensure_dup,
    list_candidates,
    upsert_candidates,
)
from core.guid_demand import list_open, record_guid_demand
from core.guid_gate import HINT_CHOOSE, MissingItem, OrderLine, REASON_AMBIGUOUS, REASON_DUP
from core.guid_gate import check_invoice_guids
from core.nomenclature_guid import ensure_schema as ensure_guid
from core.nomenclature_guid import get_guid_for_invoice, upsert
from core.plate_guid_choice import ensure_schema as ensure_choice
from core.plate_guid_choice import get_choice

GUID_A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
GUID_B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
GONE = "cccccccc-cccc-cccc-cccc-cccccccccccc"


def _service(tmp_path: Path) -> tuple[NomenclatureQueueService, Path]:
    db = tmp_path / "pb.db"
    conn = sqlite3.connect(db)
    try:
        ensure_guid(conn)
        ensure_dup(conn)
        ensure_choice(conn)
        conn.execute(
            """
            CREATE TABLE prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        conn.execute(
            'INSERT INTO prays_plity VALUES (?, ?), (?, ?)',
            (GUID_A, "ЛВ60.12-4", GUID_B, "ЛВ60.12-4"),
        )
        upsert(conn, "pile", "С50.30-8", match_status="ambiguous")
        upsert_candidates(
            conn,
            "pile",
            "С50.30-8",
            [(GUID_A, "Сваи С 50.30-8", 1000.0), (GUID_B, "Сваи С 50.30-8", 1100.0)],
        )
        record_guid_demand(
            conn,
            12,
            [
                MissingItem(
                    line=OrderLine("pile", "С50.30-8"),
                    reason=REASON_AMBIGUOUS,
                    action_hint=HINT_CHOOSE,
                ),
                MissingItem(
                    line=OrderLine("plate", "ЛВ60.12-4"),
                    reason=REASON_DUP,
                    action_hint=HINT_CHOOSE,
                ),
            ],
        )
        conn.commit()
    finally:
        conn.close()
    return NomenclatureQueueService(db), db


def test_resolve_nonplate_sets_manual_and_clears_candidates(tmp_path: Path) -> None:
    service, db = _service(tmp_path)
    result = service.resolve_duplicate(
        scope="nonplate",
        key="С50.30-8",
        chosen_guid=GUID_B,
        note="по решению бухгалтерии",
        decided_by="economist",
        product_kind="pile",
    )
    assert result.match_status == "manual"
    conn = sqlite3.connect(db)
    try:
        assert get_guid_for_invoice(conn, "pile", "С50.30-8") == GUID_B
        assert list_candidates(conn, kind="pile", mark="С50.30-8") == []
    finally:
        conn.close()


def test_resolve_plate_writes_choice(tmp_path: Path) -> None:
    service, db = _service(tmp_path)
    result = service.resolve_duplicate(
        scope="plate",
        key="ЛВ60.12-4",
        chosen_guid=GUID_A,
        note="бухгалтерия",
        decided_by="economist",
    )
    assert result.scope == "plate"
    conn = sqlite3.connect(db)
    try:
        choice = get_choice(conn, "ЛВ 60.12-4")
        assert choice is not None
        assert choice.chosen_guid == GUID_A
        report = check_invoice_guids([OrderLine("plate", "ЛВ60.12-4")], conn)
    finally:
        conn.close()
    assert report.missing == ()
    assert len(report.ready) == 1
    assert report.ready[0].guid == GUID_A


def test_missing_candidate_returns_409(tmp_path: Path) -> None:
    service, _db = _service(tmp_path)
    try:
        service.resolve_duplicate(
            scope="nonplate",
            key="С50.30-8",
            chosen_guid=GONE,
            note="x",
            decided_by="economist",
            product_kind="pile",
        )
        raised = False
    except NomenclatureQueueError as exc:
        raised = True
        assert exc.status_code == 409
        assert "исчез" in exc.message
    assert raised

    still_open = service.list_tasks()
    keys = {item.key for item in still_open.duplicates if item.scope == "nonplate"}
    assert "С50.30-8" in keys
    conn = sqlite3.connect(_db)
    try:
        demand_marks = {row.mark for row in list_open(conn)}
    finally:
        conn.close()
    assert "С50.30-8" in demand_marks
    assert "ЛВ60.12-4" in demand_marks


def test_resolve_nonplate_closes_demand_keeps_neighbor(tmp_path: Path) -> None:
    service, db = _service(tmp_path)
    service.resolve_duplicate(
        scope="nonplate",
        key="С50.30-8",
        chosen_guid=GUID_B,
        note="по решению бухгалтерии",
        decided_by="economist",
        product_kind="pile",
    )
    conn = sqlite3.connect(db)
    try:
        opened = {(row.product_kind, row.mark) for row in list_open(conn)}
    finally:
        conn.close()
    assert ("pile", "С50.30-8") not in opened
    assert ("plate", "ЛВ60.12-4") in opened


def test_resolve_plate_closes_demand(tmp_path: Path) -> None:
    service, db = _service(tmp_path)
    service.resolve_duplicate(
        scope="plate",
        key="ЛВ60.12-4",
        chosen_guid=GUID_A,
        note="бухгалтерия",
        decided_by="economist",
    )
    conn = sqlite3.connect(db)
    try:
        opened = {(row.product_kind, row.mark) for row in list_open(conn)}
    finally:
        conn.close()
    assert ("plate", "ЛВ60.12-4") not in opened
    assert ("pile", "С50.30-8") in opened
