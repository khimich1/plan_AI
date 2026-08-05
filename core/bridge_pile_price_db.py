#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс мостовых свай в pb.db (таблица bridge_pile_prices).

Grades: только B25 / B30 (колонки 25 и 30). Нулевые ячейки не импортируются.
Алиасы «C8-35T4; C8-35В4» → несколько mark с общим variant_group.
Lookup нормализует C↔С, B↔В, T↔Т; display mark остаётся как ввёл менеджер.
"""

from __future__ import annotations

import hashlib
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

# Only grades present in the bridge-pile price list.
BRIDGE_PILE_GRADE_CODES = ("B25", "B30")

BridgePilePriceRow = Tuple[str, str, float, str]  # mark, grade, price, variant_group

_DEFAULT_SHEET_NAMES = ("прайс", "price")


def normalize_bridge_pile_mark_for_lookup(mark: str) -> str:
    """Normalize mark for price lookup only (does not rewrite display text)."""
    text = str(mark or "").strip().upper()
    text = (
        text.replace("С", "C")
        .replace("В", "B")
        .replace("Т", "T")
    )
    text = re.sub(r"\s+", "", text)
    return text


def grade_code_from_bridge_header(value) -> Optional[str]:
    """Map price-list header 25/30 (or B25/B30) to bridge grade codes."""
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        if numeric == 25:
            return "B25"
        if numeric == 30:
            return "B30"
        return None
    text = str(value).strip().lower().replace(",", ".")
    if text in {"25", "b25"}:
        return "B25"
    if text in {"30", "b30"}:
        return "B30"
    return None


def grade_code_from_bridge_value(value) -> Optional[str]:
    """Map order-line grade token to B25/B30 (bridge catalog only)."""
    return grade_code_from_bridge_header(value)


def split_alias_marks(raw_name: str) -> List[str]:
    """Split «C8-35T4; C8-35В4» into trimmed synonym marks."""
    parts = [p.strip() for p in str(raw_name or "").split(";")]
    return [p for p in parts if p and p not in {"-", "—", "–"}]


def _variant_group_id(raw_name: str) -> str:
    key = normalize_bridge_pile_mark_for_lookup(str(raw_name or "").split(";")[0])
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]


def _is_bridge_pile_data_row(name: str) -> bool:
    mark = str(name or "").strip()
    if not mark or mark in {"-", "—", "–"}:
        return False
    upper = mark.upper()
    if upper == "НАИМЕНОВАНИЕ":
        return False
    if "СЕЧЕНИЕ" in upper or "НАГРУЗК" in upper:
        return False
    # C8-35T1 / С7-35Т5 / alias groups
    return bool(re.search(r"[СC]\s*\d+", mark, re.IGNORECASE))


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


def parse_bridge_pile_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[BridgePilePriceRow]:
    """Читает прайс мостовых свай только с листа «Прайс»; skip zero cells."""
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
        grade = grade_code_from_bridge_header(value)
        if grade is not None:
            grade_cols.append((idx, grade))

    if not grade_cols:
        return []

    rows: List[BridgePilePriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        raw_mark = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not _is_bridge_pile_data_row(raw_mark):
            continue
        aliases = split_alias_marks(raw_mark)
        if not aliases:
            continue
        variant_group = _variant_group_id(raw_mark)

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
            for mark in aliases:
                rows.append((mark, grade, price, variant_group))
    return rows


def init_bridge_pile_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS bridge_pile_prices (
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                price REAL NOT NULL,
                variant_group TEXT,
                display_name TEXT,
                price_list_date TEXT,
                imported_at TEXT DEFAULT (datetime('now')),
                PRIMARY KEY (mark, concrete_grade)
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_bridge_pile_lookup "
            "ON bridge_pile_prices(concrete_grade)"
        )
        conn.commit()
    finally:
        conn.close()


def import_bridge_pile_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_bridge_pile_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [
        (mark, grade, price, variant_group, mark, list_date, imported_at)
        for mark, grade, price, variant_group in rows
    ]

    init_bridge_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO bridge_pile_prices
                (mark, concrete_grade, price, variant_group, display_name,
                 price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
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
    key = normalize_bridge_pile_mark_for_lookup(mark)
    if not key:
        return []
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT mark FROM bridge_pile_prices")
    return [
        row[0]
        for row in cur.fetchall()
        if normalize_bridge_pile_mark_for_lookup(row[0]) == key
    ]


def get_bridge_pile_price(
    mark: str,
    concrete_grade: str = "B25",
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Return price for mark+grade via synonym-aware lookup."""
    if not mark:
        return None

    grade = (concrete_grade or "").strip()
    if grade not in BRIDGE_PILE_GRADE_CODES:
        return None

    init_bridge_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        matches = _find_matching_marks(conn, mark)
        if not matches:
            return None
        cur = conn.cursor()
        for stored_mark in matches:
            cur.execute(
                "SELECT price FROM bridge_pile_prices "
                "WHERE mark = ? AND concrete_grade = ?",
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
    """Return priced grades for mark (subset of B25/B30), order preserved."""
    if not mark:
        return []

    init_bridge_pile_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        matches = _find_matching_marks(conn, mark)
        if not matches:
            return []
        cur = conn.cursor()
        found: set[str] = set()
        for stored_mark in matches:
            cur.execute(
                "SELECT concrete_grade FROM bridge_pile_prices WHERE mark = ?",
                (stored_mark,),
            )
            for row in cur.fetchall():
                grade = str(row[0])
                if grade in BRIDGE_PILE_GRADE_CODES:
                    found.add(grade)
        return [g for g in BRIDGE_PILE_GRADE_CODES if g in found]
    finally:
        conn.close()


def resolve_default_bridge_pile_grade(
    mark: str,
    *,
    preferred: Optional[str] = None,
    db_path: str = DEFAULT_DB,
) -> Optional[str]:
    """Pick grade: single available → that; else preferred if priced; else B25; else first."""
    available = list_available_grades(mark, db_path=db_path)
    if not available:
        return None
    if len(available) == 1:
        return available[0]
    if preferred and preferred in available:
        return preferred
    if "B25" in available:
        return "B25"
    return available[0]


__all__ = [
    "BRIDGE_PILE_GRADE_CODES",
    "BridgePilePriceRow",
    "get_bridge_pile_price",
    "grade_code_from_bridge_header",
    "grade_code_from_bridge_value",
    "import_bridge_pile_prices_from_xlsx",
    "init_bridge_pile_prices_schema",
    "list_available_grades",
    "normalize_bridge_pile_mark_for_lookup",
    "parse_bridge_pile_price_rows_from_xlsx",
    "resolve_default_bridge_pile_grade",
    "split_alias_marks",
]
