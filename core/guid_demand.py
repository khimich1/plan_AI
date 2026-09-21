#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Спрос GUID из архивного КП (таблица guid_demand в pb.db)."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional, Sequence

from core.guid_gate import (
    MissingItem,
    OrderLine,
    REASON_AMBIGUOUS,
    REASON_DUP,
    REASON_MISSING_U,
    check_invoice_guids,
)

FIELD_GUID = "guid_1c"
FIELD_GUID_U = "guid_1c_u"
FIELD_DUPLICATE = "duplicate"
FIELDS = frozenset({FIELD_GUID, FIELD_GUID_U, FIELD_DUPLICATE})

STATE_OPEN = "open"
STATE_RESOLVED = "resolved"
STATES = frozenset({STATE_OPEN, STATE_RESOLVED})

_ROW_COLUMNS = (
    "product_kind",
    "mark",
    "field",
    "reason",
    "hint",
    "state",
    "kp_ids_json",
    "opened_at",
    "resolved_at",
)
_ROW_SELECT = ", ".join(_ROW_COLUMNS)


@dataclass(frozen=True)
class GuidDemandRow:
    product_kind: str
    mark: str
    field: str
    reason: str
    hint: str
    state: str
    kp_ids: tuple[int, ...]
    opened_at: str
    resolved_at: Optional[str]


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _normalize_text(value: object, *, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"{field} обязателен")
    return text


def _parse_kp_ids(raw: object) -> list[int]:
    if isinstance(raw, (list, tuple)):
        values = raw
    else:
        try:
            values = json.loads(str(raw or "[]"))
        except json.JSONDecodeError:
            return []
        if not isinstance(values, list):
            return []
    ids: list[int] = []
    seen: set[int] = set()
    for item in values:
        try:
            kp_id = int(item)
        except (TypeError, ValueError):
            continue
        if kp_id in seen:
            continue
        seen.add(kp_id)
        ids.append(kp_id)
    return ids


def _merge_kp_ids(existing: Sequence[int], extra: Sequence[int]) -> list[int]:
    merged = list(existing)
    seen = set(merged)
    for kp_id in extra:
        value = int(kp_id)
        if value in seen:
            continue
        seen.add(value)
        merged.append(value)
    return merged


def _row_from_tuple(values: tuple) -> GuidDemandRow:
    return GuidDemandRow(
        product_kind=values[0],
        mark=values[1],
        field=values[2],
        reason=values[3],
        hint=values[4],
        state=values[5],
        kp_ids=tuple(_parse_kp_ids(values[6])),
        opened_at=values[7],
        resolved_at=values[8],
    )


def field_for_reason(reason: str) -> str:
    text = str(reason or "").strip()
    if text == REASON_MISSING_U:
        return FIELD_GUID_U
    if text in {REASON_DUP, REASON_AMBIGUOUS}:
        return FIELD_DUPLICATE
    return FIELD_GUID


def ensure_schema(conn: sqlite3.Connection) -> None:
    """CREATE TABLE IF NOT EXISTS guid_demand."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS guid_demand (
            product_kind TEXT NOT NULL,
            mark TEXT NOT NULL,
            field TEXT NOT NULL,
            reason TEXT NOT NULL,
            hint TEXT NOT NULL,
            state TEXT NOT NULL,
            kp_ids_json TEXT NOT NULL,
            opened_at TEXT NOT NULL,
            resolved_at TEXT,
            PRIMARY KEY (product_kind, mark, field)
        )
        """
    )


def _get(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
    field: str,
) -> Optional[GuidDemandRow]:
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM guid_demand "
        "WHERE product_kind = ? AND mark = ? AND field = ?",
        (product_kind, mark, field),
    )
    values = cur.fetchone()
    if values is None:
        return None
    return _row_from_tuple(values)


def _upsert_open(
    conn: sqlite3.Connection,
    *,
    product_kind: str,
    mark: str,
    field: str,
    reason: str,
    hint: str,
    extra_kp_ids: Sequence[int],
) -> GuidDemandRow:
    now = _now_iso()
    existing = _get(conn, product_kind, mark, field)
    kp_ids = _merge_kp_ids(existing.kp_ids if existing else (), extra_kp_ids)
    opened_at = existing.opened_at if existing is not None else now
    conn.execute(
        """
        INSERT INTO guid_demand (
            product_kind, mark, field, reason, hint, state,
            kp_ids_json, opened_at, resolved_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
        ON CONFLICT(product_kind, mark, field) DO UPDATE SET
            reason = excluded.reason,
            hint = excluded.hint,
            state = excluded.state,
            kp_ids_json = excluded.kp_ids_json,
            resolved_at = NULL
        """,
        (
            product_kind,
            mark,
            field,
            reason,
            hint,
            STATE_OPEN,
            json.dumps(kp_ids),
            opened_at,
        ),
    )
    row = _get(conn, product_kind, mark, field)
    if row is None:
        raise RuntimeError("guid_demand upsert не сохранил строку")
    return row


def _set_resolved(conn: sqlite3.Connection, row: GuidDemandRow) -> None:
    conn.execute(
        """
        UPDATE guid_demand
        SET state = ?, resolved_at = ?
        WHERE product_kind = ? AND mark = ? AND field = ?
        """,
        (STATE_RESOLVED, _now_iso(), row.product_kind, row.mark, row.field),
    )


def record_guid_demand(
    conn: sqlite3.Connection,
    kp_id: int,
    missing: Sequence[MissingItem],
) -> None:
    """Записать открытый спрос по дыркам резолвера. Пустой missing — no-op."""
    ensure_schema(conn)
    if not missing:
        return
    extra = [int(kp_id)]
    seen: set[tuple[str, str, str]] = set()
    for item in missing:
        kind = _normalize_text(item.line.product_kind, field="product_kind").lower()
        mark = _normalize_text(item.line.mark, field="mark")
        field = field_for_reason(item.reason)
        key = (kind, mark, field)
        if key in seen:
            continue
        seen.add(key)
        _upsert_open(
            conn,
            product_kind=kind,
            mark=mark,
            field=field,
            reason=str(item.reason or "").strip() or "нет GUID 1С",
            hint=str(item.action_hint or "").strip(),
            extra_kp_ids=extra,
        )


def list_open(conn: sqlite3.Connection) -> list[GuidDemandRow]:
    ensure_schema(conn)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM guid_demand"
        " WHERE state = ? ORDER BY product_kind, mark, field",
        (STATE_OPEN,),
    )
    return [_row_from_tuple(values) for values in cur]


def sweep_guid_demand(conn: sqlite3.Connection) -> None:
    """Повторно прогнать резолвер по open-строкам: закрыть или сменить категорию."""
    ensure_schema(conn)
    for row in list_open(conn):
        report = check_invoice_guids(
            [OrderLine(row.product_kind, row.mark)],
            conn,
        )
        if not report.missing:
            _set_resolved(conn, row)
            continue
        item = report.missing[0]
        new_field = field_for_reason(item.reason)
        if new_field == row.field:
            continue
        _set_resolved(conn, row)
        _upsert_open(
            conn,
            product_kind=row.product_kind,
            mark=row.mark,
            field=new_field,
            reason=str(item.reason or "").strip(),
            hint=str(item.action_hint or "").strip(),
            extra_kp_ids=row.kp_ids,
        )


__all__ = [
    "FIELD_DUPLICATE",
    "FIELD_GUID",
    "FIELD_GUID_U",
    "FIELDS",
    "GuidDemandRow",
    "STATE_OPEN",
    "STATE_RESOLVED",
    "STATES",
    "ensure_schema",
    "field_for_reason",
    "list_open",
    "record_guid_demand",
    "sweep_guid_demand",
]
