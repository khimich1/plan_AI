#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Мини-справочник свай ``pile_catalog`` (SHIP-100): парсинг прайса и upsert.

Источник — лист «Вес и объем» прайса цельных свай: заголовок в строке 2,
данные со строки 3 (марка, объём м³, вес кг, шт. на а/м 20тн).
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Optional

from openpyxl import load_workbook

PILE_WEIGHT_SHEET = "Вес и объем"
PILE_HEADER_ROW = 2


@dataclass(frozen=True)
class PileCatalogEntry:
    mark: str
    length_m: Optional[float]
    section_mm: Optional[int]
    volume_m3: Optional[float]
    weight_kg: float
    pcs_per_20t: Optional[int]


def parse_pile_mark(mark: str) -> tuple[Optional[float], Optional[int]]:
    """«С137,5.40» → (13.75 м, 400 мм); «С60.30» → (6.0 м, 300 мм).

    Длина в марке — в дециметрах (допустима запятая как разделитель),
    сечение — в сантиметрах.
    """
    text = mark.strip().lstrip("СсCc").strip()
    length_part, sep, section_part = text.rpartition(".")
    if not sep:
        return None, None
    try:
        length_dm = float(length_part.replace(",", "."))
        section_cm = float(section_part.replace(",", "."))
    except ValueError:
        return None, None
    return length_dm / 10.0, int(round(section_cm * 10))


def _as_float(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text or text == "-":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _as_int(value: object) -> Optional[int]:
    number = _as_float(value)
    if number is None:
        return None
    return int(round(number))


def parse_pile_catalog_from_xlsx(
    xlsx_path: str,
    *,
    sheet: str = PILE_WEIGHT_SHEET,
) -> list[PileCatalogEntry]:
    """Разобрать лист «Вес и объем» в записи каталога (марки без веса пропускаются)."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Лист «{sheet}» не найден в файле {xlsx_path}")
        ws = wb[sheet]
        entries: list[PileCatalogEntry] = []
        seen: set[str] = set()
        for row in ws.iter_rows(min_row=PILE_HEADER_ROW + 1, values_only=True):
            if not row:
                continue
            mark = str(row[0]).strip() if row[0] is not None else ""
            if not mark or mark in seen:
                continue
            weight = _as_float(row[2] if len(row) > 2 else None)
            if weight is None or weight <= 0:
                continue
            length_m, section_mm = parse_pile_mark(mark)
            entries.append(
                PileCatalogEntry(
                    mark=mark,
                    length_m=length_m,
                    section_mm=section_mm,
                    volume_m3=_as_float(row[1] if len(row) > 1 else None),
                    weight_kg=weight,
                    pcs_per_20t=_as_int(row[3] if len(row) > 3 else None),
                )
            )
            seen.add(mark)
        return entries
    finally:
        wb.close()


def upsert_pile_catalog(db_path: str, entries: list[PileCatalogEntry]) -> tuple[int, int]:
    """Upsert по ``mark`` (безопасен для повторного запуска). Возвращает (inserted, updated)."""
    from core.kp_db_schema import ensure_schema

    ensure_schema(db_path)
    inserted = 0
    updated = 0
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        for entry in entries:
            cur.execute("SELECT id FROM pile_catalog WHERE mark = ?", (entry.mark,))
            exists = cur.fetchone() is not None
            cur.execute(
                """
                INSERT INTO pile_catalog (mark, length_m, section_mm, volume_m3, weight_kg, pcs_per_20t)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(mark) DO UPDATE SET
                    length_m = excluded.length_m,
                    section_mm = excluded.section_mm,
                    volume_m3 = excluded.volume_m3,
                    weight_kg = excluded.weight_kg,
                    pcs_per_20t = excluded.pcs_per_20t
                """,
                (
                    entry.mark,
                    entry.length_m,
                    entry.section_mm,
                    entry.volume_m3,
                    entry.weight_kg,
                    entry.pcs_per_20t,
                ),
            )
            if exists:
                updated += 1
            else:
                inserted += 1
        conn.commit()
    finally:
        conn.close()
    return inserted, updated
