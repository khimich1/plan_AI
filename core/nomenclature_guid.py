#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Маппинг GUID 1С для не-плитных групп в pb.db (таблица nomenclature_guid)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Iterator, Optional

PRODUCT_KINDS = frozenset(
    {"pile", "bridge_pile", "fbs", "stair_flight", "stair_step"}
)
MATCH_STATUSES = frozenset({"auto", "manual", "ambiguous", "missing"})

MATCH_STATUS_MISSING = "missing"

_UPSERT_FIELDS = frozenset({"guid_1c", "guid_1c_u", "match_status", "match_note"})
_ROW_COLUMNS = (
    "product_kind",
    "mark",
    "guid_1c",
    "guid_1c_u",
    "match_status",
    "match_note",
    "updated_at",
)
_ROW_SELECT = ", ".join(_ROW_COLUMNS)


@dataclass(frozen=True)
class NomenclatureGuidRow:
    product_kind: str
    mark: str
    guid_1c: Optional[str]
    guid_1c_u: Optional[str]
    match_status: str
    match_note: Optional[str]
    updated_at: str


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _optional_text(value: object) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _normalize_kind(product_kind: str) -> str:
    kind = str(product_kind or "").strip().lower()
    if kind not in PRODUCT_KINDS:
        raise ValueError(f"Unknown product_kind: {product_kind!r}")
    return kind


def _normalize_mark(mark: str) -> str:
    text = str(mark or "").strip()
    if not text:
        raise ValueError("mark is required")
    return text


def _normalize_status(status: str) -> str:
    value = str(status or "").strip().lower()
    if value not in MATCH_STATUSES:
        raise ValueError(f"Unknown match_status: {status!r}")
    return value


def _row_from_tuple(values: tuple) -> NomenclatureGuidRow:
    return NomenclatureGuidRow(
        product_kind=values[0],
        mark=values[1],
        guid_1c=values[2],
        guid_1c_u=values[3],
        match_status=values[4],
        match_note=values[5],
        updated_at=values[6],
    )


def ensure_schema(conn: sqlite3.Connection) -> None:
    """CREATE TABLE IF NOT EXISTS nomenclature_guid."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS nomenclature_guid (
            product_kind TEXT NOT NULL,
            mark TEXT NOT NULL,
            guid_1c TEXT,
            guid_1c_u TEXT,
            match_status TEXT NOT NULL,
            match_note TEXT,
            updated_at TEXT NOT NULL,
            PRIMARY KEY (product_kind, mark)
        )
        """
    )


def get_by_mark(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
) -> Optional[NomenclatureGuidRow]:
    kind = _normalize_kind(product_kind)
    mark_key = _normalize_mark(mark)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM nomenclature_guid "
        "WHERE product_kind = ? AND mark = ?",
        (kind, mark_key),
    )
    values = cur.fetchone()
    if values is None:
        return None
    return _row_from_tuple(values)


def upsert(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
    **fields: object,
) -> NomenclatureGuidRow:
    """Insert or update by (product_kind, mark). Unspecified fields stay as-is."""
    extra = set(fields) - _UPSERT_FIELDS
    if extra:
        raise TypeError(f"Unexpected fields: {sorted(extra)}")

    kind = _normalize_kind(product_kind)
    mark_key = _normalize_mark(mark)
    existing = get_by_mark(conn, kind, mark_key)

    guid_1c = existing.guid_1c if existing is not None else None
    guid_1c_u = existing.guid_1c_u if existing is not None else None
    match_status = existing.match_status if existing is not None else MATCH_STATUS_MISSING
    match_note = existing.match_note if existing is not None else None

    if "guid_1c" in fields:
        guid_1c = _optional_text(fields["guid_1c"])
    if "guid_1c_u" in fields:
        guid_1c_u = _optional_text(fields["guid_1c_u"])
    if "match_status" in fields:
        match_status = _normalize_status(str(fields["match_status"] or ""))
    if "match_note" in fields:
        match_note = _optional_text(fields["match_note"])

    updated_at = _now_iso()
    conn.execute(
        """
        INSERT INTO nomenclature_guid (
            product_kind, mark, guid_1c, guid_1c_u,
            match_status, match_note, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(product_kind, mark) DO UPDATE SET
            guid_1c = excluded.guid_1c,
            guid_1c_u = excluded.guid_1c_u,
            match_status = excluded.match_status,
            match_note = excluded.match_note,
            updated_at = excluded.updated_at
        """,
        (kind, mark_key, guid_1c, guid_1c_u, match_status, match_note, updated_at),
    )
    row = get_by_mark(conn, kind, mark_key)
    if row is None:
        raise RuntimeError("nomenclature_guid upsert did not persist the row")
    return row


def set_status(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
    status: str,
    note: Optional[str] = None,
) -> NomenclatureGuidRow:
    kind = _normalize_kind(product_kind)
    mark_key = _normalize_mark(mark)
    match_status = _normalize_status(status)
    existing = get_by_mark(conn, kind, mark_key)
    if existing is None:
        raise LookupError(
            f"No nomenclature_guid row for {kind!r} {mark_key!r}"
        )
    match_note = existing.match_note if note is None else _optional_text(note)
    conn.execute(
        """
        UPDATE nomenclature_guid
        SET match_status = ?, match_note = ?, updated_at = ?
        WHERE product_kind = ? AND mark = ?
        """,
        (match_status, match_note, _now_iso(), kind, mark_key),
    )
    row = get_by_mark(conn, kind, mark_key)
    if row is None:
        raise RuntimeError("nomenclature_guid set_status lost the row")
    return row


def iter_by_state(
    conn: sqlite3.Connection,
    status: str,
) -> Iterator[NomenclatureGuidRow]:
    match_status = _normalize_status(status)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM nomenclature_guid "
        "WHERE match_status = ? ORDER BY product_kind, mark",
        (match_status,),
    )
    for values in cur:
        yield _row_from_tuple(values)


def set_manual_choice(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
    *,
    guid_1c: str,
    decided_by: Optional[str] = None,
    note: Optional[str] = None,
) -> NomenclatureGuidRow:
    """Зафиксировать выбранный GUID дубля (status=manual + аудит в match_note)."""
    parts = [p for p in (decided_by, note) if p]
    match_note = "; ".join(parts) if parts else note
    return upsert(
        conn,
        product_kind,
        mark,
        guid_1c=guid_1c,
        match_status="manual",
        match_note=match_note,
    )


def get_guid_for_invoice(
    conn: sqlite3.Connection,
    product_kind: str,
    mark: str,
    reinforced: bool = False,
) -> Optional[str]:
    """Return guid_1c, or guid_1c_u when reinforced=True (pile «у» variant)."""
    row = get_by_mark(conn, product_kind, mark)
    if row is None:
        return None
    guid = row.guid_1c_u if reinforced else row.guid_1c
    return guid or None


__all__ = [
    "MATCH_STATUSES",
    "MATCH_STATUS_MISSING",
    "NomenclatureGuidRow",
    "PRODUCT_KINDS",
    "ensure_schema",
    "get_by_mark",
    "get_guid_for_invoice",
    "iter_by_state",
    "set_manual_choice",
    "set_status",
    "upsert",
]
