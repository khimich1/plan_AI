#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс лестничных маршей (ЛМ) в pb.db (таблица march_prices)."""

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

from core.pile_price_db import GRADE_CODES, grade_code_from_value
from core.price_db import DEFAULT_DB, _connect

MarchPriceRow = Tuple[str, str, float]

_DEFAULT_SHEET_NAMES = ("прайс", "price")
_FULL_NAME_PREFIX_RE = re.compile(
    r"^лестничн(?:ые|ая|ый)?\s+марш[иае]?\s+",
    re.IGNORECASE | re.UNICODE,
)


def normalize_march_mark(raw: str) -> str:
    """Strip full-name prefix; canon ЛМ 2,8 (comma); collapse whitespace."""
    mark = str(raw or "").strip()
    mark = _FULL_NAME_PREFIX_RE.sub("", mark).strip()
    mark = re.sub(r"\s{2,}", " ", mark)
    # ЛМ 2.8 / ЛМ2.8 → ЛМ 2,8 (price-list canon)
    mark = re.sub(
        r"^(ЛМ)\s*(\d+)[.,](\d+)\s*$",
        lambda m: f"{m.group(1)} {m.group(2)},{m.group(3)}",
        mark,
        flags=re.IGNORECASE,
    )
    # Normalize 1ЛМ spacing: «1ЛМ27-…» → «1ЛМ 27-…»
    mark = re.sub(
        r"^(1ЛМ)\s*(\d)",
        r"\1 \2",
        mark,
        flags=re.IGNORECASE,
    )
    # Preserve «закладные справа» as part of mark; normalize case of Latin digits only
    if mark.upper().startswith("1ЛМ"):
        # Keep Cyrillic ЛМ; normalize leading digit block casing lightly
        rest = mark[3:].lstrip()
        mark = f"1ЛМ {rest}" if rest else "1ЛМ"
    elif mark.upper().startswith("ЛМ"):
        rest = mark[2:].lstrip()
        mark = f"ЛМ {rest}" if rest else "ЛМ"
    return mark.strip()


def _is_march_data_row(name: str) -> bool:
    mark = str(name or "").strip()
    if not mark or mark in {"-", "—", "–"}:
        return False
    upper = mark.upper()
    if upper == "НАИМЕНОВАНИЕ":
        return False
    if "СЕЧЕНИЕ" in upper or "НАГРУЗК" in upper:
        return False
    stripped = _FULL_NAME_PREFIX_RE.sub("", mark).strip().upper()
    return stripped.startswith("1ЛМ") or stripped.startswith("ЛМ")


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


def parse_march_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[MarchPriceRow]:
    """Читает прайс ЛМ с листа «Прайс»; strips «Лестничные марши»; canon marks."""
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
        grade = grade_code_from_value(value)
        if grade is not None:
            grade_cols.append((idx, grade))

    if not grade_cols:
        return []

    rows: List[MarchPriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        raw_mark = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not _is_march_data_row(raw_mark):
            continue
        mark = normalize_march_mark(raw_mark)
        if not mark:
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


def init_march_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS march_prices (
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
        conn.commit()
    finally:
        conn.close()


def import_march_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_march_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [
        (mark, grade, price, f"Лестничные марши {mark}", list_date, imported_at)
        for mark, grade, price in rows
    ]

    init_march_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO march_prices
                (mark, concrete_grade, price, display_name, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            db_rows,
        )
        conn.commit()
        return len(db_rows)
    finally:
        conn.close()


def get_march_price(
    mark: str,
    concrete_grade: str = "B25",
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Возвращает цену марша по точной марке и классу бетона."""
    if not mark:
        return None

    canon = normalize_march_mark(mark)
    init_march_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT price FROM march_prices WHERE mark = ? AND concrete_grade = ?",
            (canon, concrete_grade),
        )
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


__all__ = [
    "GRADE_CODES",
    "MarchPriceRow",
    "get_march_price",
    "grade_code_from_value",
    "import_march_prices_from_xlsx",
    "init_march_prices_schema",
    "normalize_march_mark",
    "parse_march_price_rows_from_xlsx",
]
