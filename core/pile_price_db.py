#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс цельных свай в pb.db (таблица pile_prices)."""

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

from core.pile_catalog import normalize_pile_mark_key
from core.price_db import DEFAULT_DB, _connect

_TRAILING_U_RE = re.compile(r"[уУ]$")
_FRACTIONAL_LOAD_RE = re.compile(r"(-\d+)\.\d+([иИуУ]?)$")


def strip_trailing_u_suffix(mark: str) -> str:
    """Срез хвостовой «у»/«У» для lookup цены и GUID (D11)."""
    return _TRAILING_U_RE.sub("", (mark or "").strip())

PilePriceRow = Tuple[str, str, float]

GRADE_CODES = ("B15", "B20", "B22_5", "B25", "B30_granite")

_DEFAULT_SHEET_NAMES = ("прайс", "price")


def grade_code_from_value(value) -> Optional[str]:
    """Map a price-list header cell or order-line token to a GRADE_CODES entry."""
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        if numeric == 15:
            return "B15"
        if numeric == 20:
            return "B20"
        if abs(numeric - 22.5) < 0.001:
            return "B22_5"
        if numeric == 25:
            return "B25"
        if numeric == 30:
            return "B30_granite"
    text = str(value).strip().lower().replace(",", ".")
    if text in {"15", "b15"}:
        return "B15"
    if text in {"20", "b20"}:
        return "B20"
    if text in {"22.5", "22,5", "b22_5", "b22.5"}:
        return "B22_5"
    if text in {"25", "b25"}:
        return "B25"
    if text in {"30", "b30", "b30_granite"}:
        return "B30_granite"
    if "30" in text and "гранит" in text:
        return "B30_granite"
    return None


def _grade_code_from_header(value) -> Optional[str]:
    return grade_code_from_value(value)


def _is_pile_data_row(name: str) -> bool:
    mark = str(name or "").strip()
    if not mark or mark in {"-", "—", "–"}:
        return False
    upper = mark.upper()
    if upper == "НАИМЕНОВАНИЕ":
        return False
    if "СЕЧЕНИЕ" in upper or "НАГРУЗК" in upper:
        return False
    return bool(re.search(r"[СC]", mark, re.IGNORECASE)) or "ТИП" in upper


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


def parse_pile_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[PilePriceRow]:
    """Читает прайс свай с листа «Прайс»."""
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
        grade = _grade_code_from_header(value)
        if grade is not None:
            grade_cols.append((idx, grade))

    if not grade_cols:
        return []

    rows: List[PilePriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        mark = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not _is_pile_data_row(mark):
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
            rows.append((mark, grade, price))
    return rows


def init_pile_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS pile_prices (
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


def import_pile_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_pile_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [
        (mark, grade, price, list_date, imported_at)
        for mark, grade, price in rows
    ]

    init_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
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
    key = normalize_pile_mark_key(mark)
    if not key:
        return []
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT mark FROM pile_prices")
    return [
        row[0]
        for row in cur.fetchall()
        if normalize_pile_mark_key(row[0]) == key
    ]


def _pile_price_lookup_candidates(mark: str) -> List[str]:
    """Exact mark, then strip trailing у, then collapse load ``.\\d``."""
    ordered: List[str] = []

    def add(value: str) -> None:
        text = (value or "").strip()
        if text and text not in ordered:
            ordered.append(text)

    add(mark)
    without_u = strip_trailing_u_suffix(mark)
    add(without_u)
    add(_FRACTIONAL_LOAD_RE.sub(r"\1\2", without_u))
    add(_FRACTIONAL_LOAD_RE.sub(r"\1\2", mark))
    return ordered


def _price_for_mark(
    conn: sqlite3.Connection,
    mark: str,
    concrete_grade: str,
) -> Optional[float]:
    cur = conn.cursor()
    cur.execute(
        "SELECT price FROM pile_prices WHERE mark = ? AND concrete_grade = ?",
        (mark, concrete_grade),
    )
    row = cur.fetchone()
    if row:
        return float(row[0])

    prices = []
    for stored_mark in _find_matching_marks(conn, mark):
        cur.execute(
            "SELECT price FROM pile_prices WHERE mark = ? AND concrete_grade = ?",
            (stored_mark, concrete_grade),
        )
        found = cur.fetchone()
        if found:
            prices.append(float(found[0]))
    unique = set(prices)
    if len(unique) == 1:
        return next(iter(unique))
    return None


def get_pile_price(
    mark: str,
    concrete_grade: str = "B25",
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Возвращает цену сваи по марке и классу бетона.

    Порядок: точное совпадение → C↔С/пробелы → срез хвостовой ``у`` →
    схлопнуть ``.\\d`` только у числа нагрузки. Суффикс ``и`` не срезается.
    Несколько совпадений с разными ценами не угадываются.
    """
    queried = (mark or "").strip()
    if not queried:
        return None

    init_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        for candidate in _pile_price_lookup_candidates(queried):
            price = _price_for_mark(conn, candidate, concrete_grade)
            if price is not None:
                return price
        return None
    finally:
        conn.close()
