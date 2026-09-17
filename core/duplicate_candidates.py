#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Персистентные кандидаты дублей GUID не-плит (таблица duplicate_candidates)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional, Sequence

from core.nomenclature_guid import PRODUCT_KINDS

_ROW_COLUMNS = (
    "kind",
    "mark_norm",
    "guid",
    "name",
    "price",
    "first_seen",
    "active",
)
_ROW_SELECT = ", ".join(_ROW_COLUMNS)


@dataclass(frozen=True)
class DuplicateCandidate:
    kind: str
    mark_norm: str
    guid: str
    name: str
    price: Optional[float]
    first_seen: str
    active: bool


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _normalize_kind(kind: str) -> str:
    value = str(kind or "").strip().lower()
    if value not in PRODUCT_KINDS:
        raise ValueError(f"Unknown product_kind: {kind!r}")
    return value


def _normalize_mark(mark: str) -> str:
    text = str(mark or "").strip()
    if not text:
        raise ValueError("mark is required")
    return text


def _normalize_guid(guid: str) -> str:
    text = str(guid or "").strip()
    if text.startswith("{") and text.endswith("}"):
        text = text[1:-1].strip()
    text = text.lower()
    if not text:
        raise ValueError("GUID обязателен")
    return text


def _optional_name(value: object) -> str:
    text = str(value or "").strip()
    return text


def _optional_price(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if number > 0 else None
    return None


def _row_from_tuple(values: tuple) -> DuplicateCandidate:
    return DuplicateCandidate(
        kind=values[0],
        mark_norm=values[1],
        guid=values[2],
        name=values[3],
        price=values[4],
        first_seen=values[5],
        active=bool(values[6]),
    )


def ensure_schema(conn: sqlite3.Connection) -> None:
    """CREATE TABLE IF NOT EXISTS duplicate_candidates."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS duplicate_candidates (
            kind TEXT NOT NULL,
            mark_norm TEXT NOT NULL,
            guid TEXT NOT NULL,
            name TEXT NOT NULL,
            price REAL,
            first_seen TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            PRIMARY KEY (kind, mark_norm, guid)
        )
        """
    )


def upsert_candidates(
    conn: sqlite3.Connection,
    kind: str,
    mark: str,
    items: Sequence[tuple[str, str, Optional[float]]],
) -> None:
    """Записать кандидатов дубля. Повтор не плодит строки, first_seen сохраняется."""
    ensure_schema(conn)
    kind_key = _normalize_kind(kind)
    mark_key = _normalize_mark(mark)
    now = _now_iso()
    for guid, name, price in items:
        conn.execute(
            """
            INSERT INTO duplicate_candidates (
                kind, mark_norm, guid, name, price, first_seen, active
            ) VALUES (?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(kind, mark_norm, guid) DO UPDATE SET
                name = excluded.name,
                price = excluded.price,
                active = 1
            """,
            (
                kind_key,
                mark_key,
                _normalize_guid(guid),
                _optional_name(name) or mark_key,
                _optional_price(price),
                now,
            ),
        )


def list_candidates(
    conn: sqlite3.Connection,
    *,
    kind: Optional[str] = None,
    mark: Optional[str] = None,
) -> list[DuplicateCandidate]:
    ensure_schema(conn)
    clauses = ["active = 1"]
    params: list[object] = []
    if kind is not None:
        clauses.append("kind = ?")
        params.append(_normalize_kind(kind))
    if mark is not None:
        clauses.append("mark_norm = ?")
        params.append(_normalize_mark(mark))
    where = " AND ".join(clauses)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM duplicate_candidates"
        f" WHERE {where} ORDER BY kind, mark_norm, guid",
        params,
    )
    return [_row_from_tuple(values) for values in cur]


def clear_candidates(
    conn: sqlite3.Connection,
    kind: str,
    mark: str,
) -> None:
    ensure_schema(conn)
    conn.execute(
        """
        UPDATE duplicate_candidates
        SET active = 0
        WHERE kind = ? AND mark_norm = ?
        """,
        (_normalize_kind(kind), _normalize_mark(mark)),
    )


__all__ = [
    "DuplicateCandidate",
    "clear_candidates",
    "ensure_schema",
    "list_candidates",
    "upsert_candidates",
]
