#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Персистентная очередь «💰 ввести цену» (изделия 1С без нашей цены)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional, Sequence

from core.nomenclature_guid import PRODUCT_KINDS

STATES = frozenset({"open", "resolved"})
STATE_OPEN = "open"
STATE_RESOLVED = "resolved"

_ROW_COLUMNS = (
    "guid",
    "mark_norm",
    "kind",
    "name",
    "state",
    "resolved_by",
    "resolved_price",
    "resolved_at",
    "first_seen",
)
_ROW_SELECT = ", ".join(_ROW_COLUMNS)


@dataclass(frozen=True)
class PriceQueueItem:
    guid: str
    mark_norm: str
    kind: str
    name: str
    state: str
    resolved_by: Optional[str]
    resolved_price: Optional[float]
    resolved_at: Optional[str]
    first_seen: str


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _normalize_kind(kind: str) -> str:
    value = str(kind or "").strip().lower()
    if value not in PRODUCT_KINDS:
        raise ValueError(f"Unknown product_kind: {kind!r}")
    return value


def _normalize_guid(guid: str) -> str:
    text = str(guid or "").strip()
    if text.startswith("{") and text.endswith("}"):
        text = text[1:-1].strip()
    text = text.lower()
    if not text:
        raise ValueError("GUID обязателен")
    return text


def _normalize_text(value: object, *, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"{field} обязателен")
    return text


def _optional_text(value: object) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _row_from_tuple(values: tuple) -> PriceQueueItem:
    return PriceQueueItem(
        guid=values[0],
        mark_norm=values[1],
        kind=values[2],
        name=values[3],
        state=values[4],
        resolved_by=values[5],
        resolved_price=values[6],
        resolved_at=values[7],
        first_seen=values[8],
    )


def ensure_schema(conn: sqlite3.Connection) -> None:
    """CREATE TABLE IF NOT EXISTS price_queue."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS price_queue (
            guid TEXT NOT NULL PRIMARY KEY,
            mark_norm TEXT NOT NULL,
            kind TEXT NOT NULL,
            name TEXT NOT NULL,
            state TEXT NOT NULL,
            resolved_by TEXT,
            resolved_price REAL,
            resolved_at TEXT,
            first_seen TEXT NOT NULL
        )
        """
    )


def upsert_unmatched(
    conn: sqlite3.Connection,
    items: Sequence[tuple[str, str, str, str]],
) -> None:
    """Добавить unmatched_1c в очередь.

    Каждый элемент: (guid, mark_norm, kind, name).
    Повтор не дублирует; state=resolved не сбрасывается.
    """
    ensure_schema(conn)
    now = _now_iso()
    for guid, mark_norm, kind, name in items:
        conn.execute(
            """
            INSERT INTO price_queue (
                guid, mark_norm, kind, name, state,
                resolved_by, resolved_price, resolved_at, first_seen
            ) VALUES (?, ?, ?, ?, ?, NULL, NULL, NULL, ?)
            ON CONFLICT(guid) DO UPDATE SET
                mark_norm = excluded.mark_norm,
                kind = excluded.kind,
                name = excluded.name
            """,
            (
                _normalize_guid(guid),
                _normalize_text(mark_norm, field="mark_norm"),
                _normalize_kind(kind),
                _normalize_text(name, field="name"),
                STATE_OPEN,
                now,
            ),
        )


def list_open(conn: sqlite3.Connection) -> list[PriceQueueItem]:
    ensure_schema(conn)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM price_queue"
        " WHERE state = ? ORDER BY kind, mark_norm, guid",
        (STATE_OPEN,),
    )
    return [_row_from_tuple(values) for values in cur]


def set_resolved(
    conn: sqlite3.Connection,
    guid: str,
    *,
    resolved_by: str,
    resolved_price: float,
) -> PriceQueueItem:
    ensure_schema(conn)
    key = _normalize_guid(guid)
    existing = conn.execute(
        f"SELECT {_ROW_SELECT} FROM price_queue WHERE guid = ?",
        (key,),
    ).fetchone()
    if existing is None:
        raise LookupError(f"Нет задачи price_queue для GUID {key}")
    conn.execute(
        """
        UPDATE price_queue
        SET state = ?, resolved_by = ?, resolved_price = ?, resolved_at = ?
        WHERE guid = ?
        """,
        (
            STATE_RESOLVED,
            _optional_text(resolved_by),
            float(resolved_price),
            _now_iso(),
            key,
        ),
    )
    row = conn.execute(
        f"SELECT {_ROW_SELECT} FROM price_queue WHERE guid = ?",
        (key,),
    ).fetchone()
    if row is None:
        raise RuntimeError("price_queue set_resolved потерял строку")
    return _row_from_tuple(row)


__all__ = [
    "PriceQueueItem",
    "STATE_OPEN",
    "STATE_RESOLVED",
    "STATES",
    "ensure_schema",
    "list_open",
    "set_resolved",
    "upsert_unmatched",
]
