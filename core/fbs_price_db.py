#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс ФБС в pb.db (таблица fbs_prices).

Grades: B7_5 / B20 / B22_5 / B25 (колонки 7.5 | 20 | 22.5 | 25).
Плотная матрица; нулевые ячейки не импортируются.
Lookup нормализует пробелы/регистр и Т↔T; display mark — как ввёл менеджер.
"""

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

from core.price_db import DEFAULT_DB, _connect

# Grades present in the FBS price list.
FBS_GRADE_CODES = ("B7_5", "B20", "B22_5", "B25")

FbsPriceRow = Tuple[str, str, float]  # mark, grade, price

_DEFAULT_SHEET_NAMES = ("прайс", "price")


def normalize_fbs_mark_for_lookup(mark: str) -> str:
    """Normalize mark for price lookup only (does not rewrite display text)."""
    text = str(mark or "").strip().upper()
    text = text.replace("Т", "T")
    text = re.sub(r"\s+", "", text)
    return text


def grade_code_from_fbs_header(value) -> Optional[str]:
    """Map price-list header 7.5/20/22.5/25 to FBS grade codes."""
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        if abs(numeric - 7.5) < 0.001:
            return "B7_5"
        if numeric == 20:
            return "B20"
        if abs(numeric - 22.5) < 0.001:
            return "B22_5"
        if numeric == 25:
            return "B25"
        return None
    text = str(value).strip().lower().replace(",", ".")
    if text in {"7.5", "7,5", "b7_5", "b7.5"}:
        return "B7_5"
    if text in {"20", "b20"}:
        return "B20"
    if text in {"22.5", "22,5", "b22_5", "b22.5"}:
        return "B22_5"
    if text in {"25", "b25"}:
        return "B25"
    return None


def grade_code_from_fbs_value(value) -> Optional[str]:
    """Map order-line grade token to FBS grade codes."""
    return grade_code_from_fbs_header(value)


def _is_fbs_data_row(name: str) -> bool:
    mark = str(name or "").strip()
    if not mark or mark in {"-", "—", "–"}:
        return False
    upper = mark.upper()
    if upper == "НАИМЕНОВАНИЕ":
        return False
    if "СЕЧЕНИЕ" in upper or "НАГРУЗК" in upper:
        return False
    return bool(re.search(r"ФБС\s*\d", mark, re.IGNORECASE))


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


def parse_fbs_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[FbsPriceRow]:
    """Читает прайс ФБС только с листа «Прайс»; skip zero cells."""
    if pd is None or not os.path.exists(xlsx_path):
        return []

    sheet_name = _find_price_sheet_name(xlsx_path, preferred_sheet)
    if sheet_name is None:
        return []

    raw_df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)
    header_row_idx = _find_header_row(raw_df)
    if header_row_idx is None:
        return []

    header = raw_df.iloc[header_row_idx]
    name_col_idx = next(
        (idx for idx, value in enumerate(header.tolist()) if str(value).strip().lower() == "наименование"),
        1,
    )

    grade_cols: List[Tuple[int, str]] = []
    for idx, value in enumerate(header.tolist()):
        grade = grade_code_from_fbs_header(value)
        if grade is not None:
            grade_cols.append((idx, grade))

    if not grade_cols:
        return []

    rows: List[FbsPriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        raw_mark = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not _is_fbs_data_row(raw_mark):
            continue

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
            rows.append((raw_mark, grade, price))
    return rows


def init_fbs_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS fbs_prices (
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                price REAL NOT NULL,
                display_name TEXT,
                price_list_date TEXT,
                imported_at TEXT DEFAULT (datetime('now')),
                PRIMARY KEY (mark, concrete_grade)
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_fbs_lookup ON fbs_prices(concrete_grade)"
        )
        conn.commit()
    finally:
        conn.close()


def import_fbs_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_fbs_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [
        (mark, grade, price, mark, list_date, imported_at)
        for mark, grade, price in rows
    ]

    init_fbs_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO fbs_prices
                (mark, concrete_grade, price, display_name,
                 price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            db_rows,
        )
        conn.commit()
        return len(db_rows)
    finally:
        conn.close()


def _find_matching_marks(
    conn: sqlite3.Connection,
    mark: str,
) -> List[str]:
    key = normalize_fbs_mark_for_lookup(mark)
    if not key:
        return []
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT mark FROM fbs_prices")
    return [
        row[0]
        for row in cur.fetchall()
        if normalize_fbs_mark_for_lookup(row[0]) == key
    ]


def get_fbs_price(
    mark: str,
    concrete_grade: str = "B25",
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Return price for mark+grade via normalized lookup."""
    if not mark:
        return None

    grade = (concrete_grade or "").strip()
    if grade not in FBS_GRADE_CODES:
        return None

    init_fbs_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        matches = _find_matching_marks(conn, mark)
        if not matches:
            return None
        cur = conn.cursor()
        for stored_mark in matches:
            cur.execute(
                "SELECT price FROM fbs_prices WHERE mark = ? AND concrete_grade = ?",
                (stored_mark, grade),
            )
            row = cur.fetchone()
            if row:
                return float(row[0])
        return None
    finally:
        conn.close()


def list_available_grades(
    mark: str,
    db_path: str = DEFAULT_DB,
) -> List[str]:
    """Return priced grades for mark (subset of FBS_GRADE_CODES), order preserved."""
    if not mark:
        return []

    init_fbs_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        matches = _find_matching_marks(conn, mark)
        if not matches:
            return []
        cur = conn.cursor()
        found: set[str] = set()
        for stored_mark in matches:
            cur.execute(
                "SELECT concrete_grade FROM fbs_prices WHERE mark = ?",
                (stored_mark,),
            )
            for row in cur.fetchall():
                grade = str(row[0])
                if grade in FBS_GRADE_CODES:
                    found.add(grade)
        return [g for g in FBS_GRADE_CODES if g in found]
    finally:
        conn.close()


def resolve_default_fbs_grade(
    mark: str,
    *,
    preferred: Optional[str] = None,
    db_path: str = DEFAULT_DB,
) -> Optional[str]:
    """Pick grade: preferred if priced; else B25 if priced; else single/first available."""
    available = list_available_grades(mark, db_path=db_path)
    if not available:
        return None
    if preferred and preferred in available:
        return preferred
    if "B25" in available:
        return "B25"
    if len(available) == 1:
        return available[0]
    return available[0]


__all__ = [
    "FBS_GRADE_CODES",
    "FbsPriceRow",
    "get_fbs_price",
    "grade_code_from_fbs_header",
    "grade_code_from_fbs_value",
    "import_fbs_prices_from_xlsx",
    "init_fbs_prices_schema",
    "list_available_grades",
    "normalize_fbs_mark_for_lookup",
    "parse_fbs_price_rows_from_xlsx",
    "resolve_default_fbs_grade",
]
