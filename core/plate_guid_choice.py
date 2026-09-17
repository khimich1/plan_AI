#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Override выбранного GUID плиты при дубле в prays_plity (pb.db)."""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

_ROW_COLUMNS = (
    "plate_name_norm",
    "chosen_guid",
    "chosen_name",
    "decided_by",
    "decided_at",
    "note",
)
_ROW_SELECT = ", ".join(_ROW_COLUMNS)
_SPACE_AROUND_PUNCT_RE = re.compile(r"\s*([.\-–—])\s*")


@dataclass(frozen=True)
class PlateGuidChoice:
    plate_name_norm: str
    chosen_guid: str
    chosen_name: Optional[str]
    decided_by: Optional[str]
    decided_at: str
    note: Optional[str]


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _optional_text(value: object) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def normalize_plate_name(plate_name: str) -> str:
    """Ключ lookup: strip, пробелы вокруг точек/тире схлопываются.

    «ЛВ 60.12-4» ≡ «ЛВ60.12-4».
    """
    text = str(plate_name or "").strip().replace("\u00a0", " ")
    if not text:
        raise ValueError("имя плиты обязательно")
    text = _SPACE_AROUND_PUNCT_RE.sub(r"\1", text)
    return re.sub(r"\s+", "", text)


def _normalize_guid(guid: str) -> str:
    text = str(guid or "").strip()
    if text.startswith("{") and text.endswith("}"):
        text = text[1:-1].strip()
    text = text.lower()
    if not text:
        raise ValueError("GUID обязателен")
    return text


def _row_from_tuple(values: tuple) -> PlateGuidChoice:
    return PlateGuidChoice(
        plate_name_norm=values[0],
        chosen_guid=values[1],
        chosen_name=values[2],
        decided_by=values[3],
        decided_at=values[4],
        note=values[5],
    )


def ensure_schema(conn: sqlite3.Connection) -> None:
    """CREATE TABLE IF NOT EXISTS plate_guid_choice."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS plate_guid_choice (
            plate_name_norm TEXT NOT NULL PRIMARY KEY,
            chosen_guid TEXT NOT NULL,
            chosen_name TEXT,
            decided_by TEXT,
            decided_at TEXT NOT NULL,
            note TEXT
        )
        """
    )


def get_choice(
    conn: sqlite3.Connection,
    plate_name: str,
) -> Optional[PlateGuidChoice]:
    ensure_schema(conn)
    key = normalize_plate_name(plate_name)
    cur = conn.execute(
        f"SELECT {_ROW_SELECT} FROM plate_guid_choice WHERE plate_name_norm = ?",
        (key,),
    )
    values = cur.fetchone()
    if values is None:
        return None
    return _row_from_tuple(values)


def set_choice(
    conn: sqlite3.Connection,
    plate_name: str,
    *,
    chosen_guid: str,
    chosen_name: Optional[str] = None,
    decided_by: Optional[str] = None,
    note: Optional[str] = None,
) -> PlateGuidChoice:
    ensure_schema(conn)
    key = normalize_plate_name(plate_name)
    guid = _normalize_guid(chosen_guid)
    decided_at = _now_iso()
    conn.execute(
        """
        INSERT INTO plate_guid_choice (
            plate_name_norm, chosen_guid, chosen_name,
            decided_by, decided_at, note
        ) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(plate_name_norm) DO UPDATE SET
            chosen_guid = excluded.chosen_guid,
            chosen_name = excluded.chosen_name,
            decided_by = excluded.decided_by,
            decided_at = excluded.decided_at,
            note = excluded.note
        """,
        (
            key,
            guid,
            _optional_text(chosen_name),
            _optional_text(decided_by),
            decided_at,
            _optional_text(note),
        ),
    )
    row = get_choice(conn, plate_name)
    if row is None:
        raise RuntimeError("plate_guid_choice set_choice не сохранил строку")
    return row


def clear_choice(conn: sqlite3.Connection, plate_name: str) -> None:
    ensure_schema(conn)
    key = normalize_plate_name(plate_name)
    conn.execute(
        "DELETE FROM plate_guid_choice WHERE plate_name_norm = ?",
        (key,),
    )


__all__ = [
    "PlateGuidChoice",
    "clear_choice",
    "ensure_schema",
    "get_choice",
    "normalize_plate_name",
    "set_choice",
]
