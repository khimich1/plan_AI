#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Запись карточек плит из выгрузки 1С в prays_plity (pb.db).

Строка добавляется или обновляется по GUID. Строк вне файла эта функция
не удаляет и сама не коммитит.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Sequence

_GUID_COLUMN = '"Уникальный идентификатор (Номенклатура)"'
_NAME_COLUMN = '"Товар"'


@dataclass(frozen=True, slots=True)
class PlateCardWrite:
    name: str
    guid: str


@dataclass(frozen=True, slots=True)
class PlateCardReport:
    new_guids: tuple[PlateCardWrite, ...] = ()
    updated_guids: tuple[PlateCardWrite, ...] = ()
    unchanged: int = 0


def apply_plate_cards(conn: sqlite3.Connection, rows: Sequence[object]) -> PlateCardReport:
    """Добавить или обновить карточки по GUID. Caller коммитит сам."""
    _ensure_table(conn)
    existing = _load_by_guid(conn)
    final_name: dict[str, str] = {}
    order: list[str] = []
    for row in rows:
        name = str(getattr(row, "name", "") or "").strip()
        guid = _normalize_guid(str(getattr(row, "guid", "") or ""))
        if not name or not guid:
            continue
        if guid not in final_name:
            order.append(guid)
        final_name[guid] = name

    new_guids: list[PlateCardWrite] = []
    updated_guids: list[PlateCardWrite] = []
    unchanged = 0
    for guid in order:
        name = final_name[guid]
        found = existing.get(guid, [])
        if not found:
            conn.execute(
                f"INSERT INTO prays_plity ({_GUID_COLUMN}, {_NAME_COLUMN}) VALUES (?, ?)",
                (guid, name),
            )
            new_guids.append(PlateCardWrite(name=name, guid=guid))
            continue
        if {item_name for _rowid, item_name in found} == {name}:
            unchanged += 1
            continue
        for rowid, _item_name in found:
            conn.execute(
                f"UPDATE prays_plity SET {_NAME_COLUMN} = ? WHERE rowid = ?",
                (name, rowid),
            )
        updated_guids.append(PlateCardWrite(name=name, guid=guid))
    return PlateCardReport(
        new_guids=tuple(new_guids),
        updated_guids=tuple(updated_guids),
        unchanged=unchanged,
    )


def _ensure_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS prays_plity (
            {_GUID_COLUMN} TEXT,
            {_NAME_COLUMN} TEXT
        )
        """
    )


def _load_by_guid(conn: sqlite3.Connection) -> dict[str, list[tuple[int, str]]]:
    found: dict[str, list[tuple[int, str]]] = {}
    cur = conn.execute(f"SELECT rowid, {_GUID_COLUMN}, {_NAME_COLUMN} FROM prays_plity")
    for rowid, raw_guid, raw_name in cur:
        guid = _normalize_guid(str(raw_guid or ""))
        if not guid:
            continue
        found.setdefault(guid, []).append((int(rowid), str(raw_name or "")))
    return found


def _normalize_guid(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text.startswith("{") and text.endswith("}"):
        text = text[1:-1].strip()
    return text.lower()


__all__ = [
    "PlateCardReport",
    "PlateCardWrite",
    "apply_plate_cards",
]
