#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс составных свай в pb.db (таблица composite_pile_prices)."""

from __future__ import annotations

import os
import re
import sqlite3
from datetime import datetime
from typing import List, Optional, Tuple

try:
    import pandas as pd
except Exception:
    pd = None

from core.composite_pile_line_parser import is_composite_section_mark
from core.composite_pile_text_normalizer import canonicalize_composite_pile_mark
from core.pile_price_db import GRADE_CODES, grade_code_from_value
from core.price_db import DEFAULT_DB, _connect

CompositePilePriceRow = Tuple[str, str, float]  # mark, grade, price

_DEFAULT_SHEET_NAMES = ("прайс", "price")


class CompositePilePriceImportError(ValueError):
    """Кривое имя или структура прайса — импорт валится, не silent skip."""


def init_composite_pile_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS composite_pile_prices (
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                price REAL NOT NULL,
                price_list_date TEXT,
                imported_at TEXT DEFAULT (datetime('now')),
                PRIMARY KEY (mark, concrete_grade)
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _find_price_sheet_name(xlsx_path: str, preferred_sheet: Optional[str]) -> Optional[str]:
    if pd is None:
        return None
    sheets = pd.ExcelFile(xlsx_path).sheet_names
    if preferred_sheet:
        for sheet in sheets:
            if sheet.lower() == preferred_sheet.lower():
                return sheet
    for sheet in sheets:
        if sheet.strip().lower() in _DEFAULT_SHEET_NAMES:
            return sheet
    return None


def _find_header_row(raw_df) -> Optional[int]:
    for idx, row in raw_df.iterrows():
        for value in row.tolist():
            if str(value).strip().lower() == "наименование":
                return int(idx)
    return None


def _parse_price_list_date(xlsx_path: str, explicit_date: Optional[str] = None) -> Optional[str]:
    if explicit_date:
        return explicit_date
    match = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{2,4})", os.path.basename(xlsx_path))
    if not match:
        return None
    day, month, year = match.groups()
    if len(year) == 2:
        year = f"20{year}"
    return f"{year}-{month}-{day}"


def _is_section_header_row(name: str) -> bool:
    upper = str(name or "").strip().upper()
    return "СЕЧЕНИЕ" in upper or upper in {"", "-", "—", "–"}


def parse_composite_pile_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[CompositePilePriceRow]:
    """Читает лист «Прайс»; имена → канон секции. Кривое имя → ошибка."""
    if pd is None:
        raise CompositePilePriceImportError("pandas недоступен")
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(xlsx_path)

    sheet_name = _find_price_sheet_name(xlsx_path, preferred_sheet)
    if sheet_name is None:
        raise CompositePilePriceImportError("лист «Прайс» не найден")

    raw_df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)
    header_row_idx = _find_header_row(raw_df)
    if header_row_idx is None:
        raise CompositePilePriceImportError("нет строки «Наименование»")

    header = raw_df.iloc[header_row_idx]
    name_col_idx = next(
        (
            idx
            for idx, value in enumerate(header.tolist())
            if str(value).strip().lower() == "наименование"
        ),
        1,
    )

    grade_cols: List[Tuple[int, str]] = []
    for idx, value in enumerate(header.tolist()):
        grade = grade_code_from_value(value)
        if grade is not None and grade in GRADE_CODES:
            grade_cols.append((idx, grade))
    if not grade_cols:
        raise CompositePilePriceImportError("нет колонок классов бетона")

    rows: List[CompositePilePriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        raw_name = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not raw_name or raw_name.lower() == "nan" or _is_section_header_row(raw_name):
            continue
        canon = canonicalize_composite_pile_mark(raw_name)
        if not is_composite_section_mark(canon):
            raise CompositePilePriceImportError(
                f"кривое имя в прайсе составных: {raw_name!r} → {canon!r}"
            )
        for col_idx, grade in grade_cols:
            if col_idx >= len(row):
                continue
            val = row.iloc[col_idx]
            if pd.isna(val):
                continue
            try:
                price = float(str(val).replace(" ", "").replace(",", "."))
            except (TypeError, ValueError):
                continue
            if price <= 0:
                continue
            rows.append((canon, grade, price))
    return rows


def import_composite_pile_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_composite_pile_price_rows_from_xlsx(
        xlsx_path, preferred_sheet=preferred_sheet
    )
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [(mark, grade, price, list_date, imported_at) for mark, grade, price in rows]

    init_composite_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO composite_pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            db_rows,
        )
        conn.commit()
        return len(db_rows)
    finally:
        conn.close()


def get_composite_pile_price(
    mark: str,
    concrete_grade: str,
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Strict lookup канон + класс. None если нет цены."""
    canon = canonicalize_composite_pile_mark(mark)
    grade = str(concrete_grade or "").strip()
    if not canon or not grade:
        return None
    init_composite_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.execute(
            """
            SELECT price FROM composite_pile_prices
            WHERE mark = ? AND concrete_grade = ?
            """,
            (canon, grade),
        )
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def list_composite_pile_price_marks(db_path: str = DEFAULT_DB) -> set[str]:
    init_composite_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.execute("SELECT DISTINCT mark FROM composite_pile_prices")
        return {str(r[0]) for r in cur.fetchall()}
    finally:
        conn.close()
