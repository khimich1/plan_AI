#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прайс лестничных ступеней (ЛС) в pb.db (таблица step_prices)."""

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

# (mark, price, display_name)
StepPriceRow = Tuple[str, float, str]

_DEFAULT_SHEET_NAMES = ("прайс", "price")

# Короткая марка: ЛС11, ЛС14-1лев, ЛС11-Б-1, … (лишние пробелы схлопываются в normalize)
_STEP_MARK_RE = re.compile(r"ЛС\s*\S+", re.IGNORECASE | re.UNICODE)


def normalize_step_mark(raw: str) -> str:
    """Normalize short mark: upper-case, collapse spaces around dashes."""
    mark = (raw or "").strip().upper().replace("Ё", "Е")
    mark = re.sub(r"\s+", "", mark)
    # Ensure Cyrillic ЛС prefix (OCR may use Latin LS / LС)
    if mark.startswith("LS") or mark.startswith("LС") or mark.startswith("ЛS"):
        mark = "ЛС" + mark[2:]
    elif not mark.startswith("ЛС") and re.match(r"^[LЛ][SС]", mark):
        mark = "ЛС" + mark[2:]
    return mark


def extract_step_mark(text: str) -> Optional[str]:
    """Extract short ЛС… mark from display name or free text."""
    if not text:
        return None
    cleaned = str(text).replace("\u00a0", " ")
    match = _STEP_MARK_RE.search(cleaned)
    if not match:
        return None
    return normalize_step_mark(match.group(0))


def _is_step_data_row(name: str) -> bool:
    mark = str(name or "").strip()
    if not mark or mark in {"-", "—", "–"}:
        return False
    upper = mark.upper()
    if upper == "НАИМЕНОВАНИЕ":
        return False
    return extract_step_mark(mark) is not None


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


def _first_numeric_price(row, name_col_idx: int) -> Optional[float]:
    """Take the first numeric cell after the name column (ignore grade header semantics)."""
    for col_idx in range(name_col_idx + 1, len(row)):
        val = row.iloc[col_idx]
        if pd is not None and pd.isna(val):
            continue
        if val is None:
            continue
        try:
            return float(str(val).replace(" ", "").replace(",", "."))
        except (TypeError, ValueError):
            continue
    return None


def parse_step_price_rows_from_xlsx(
    xlsx_path: str,
    preferred_sheet: Optional[str] = "Прайс",
) -> List[StepPriceRow]:
    """Читает прайс ступеней с листа «Прайс» → (mark, price, display_name)."""
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

    rows: List[StepPriceRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        display_name = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        if not _is_step_data_row(display_name):
            continue
        mark = extract_step_mark(display_name)
        if not mark:
            continue
        price = _first_numeric_price(row, name_col_idx)
        if price is None:
            continue
        rows.append((mark, price, display_name))
    return rows


def init_step_prices_schema(db_path: str = DEFAULT_DB) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS step_prices (
                mark TEXT NOT NULL PRIMARY KEY,
                price REAL NOT NULL,
                display_name TEXT,
                price_list_date TEXT,
                imported_at TEXT DEFAULT (datetime('now'))
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def import_step_prices_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB,
    preferred_sheet: Optional[str] = "Прайс",
    price_list_date: Optional[str] = None,
) -> int:
    rows = parse_step_price_rows_from_xlsx(xlsx_path, preferred_sheet=preferred_sheet)
    if not rows:
        return 0

    list_date = _parse_price_list_date(xlsx_path, price_list_date)
    imported_at = datetime.now().isoformat(timespec="seconds")
    db_rows = [
        (mark, price, display_name, list_date, imported_at)
        for mark, price, display_name in rows
    ]

    init_step_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.executemany(
            """
            INSERT OR REPLACE INTO step_prices
                (mark, price, display_name, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            db_rows,
        )
        conn.commit()
        return len(db_rows)
    finally:
        conn.close()


def get_step_price(
    mark: str,
    db_path: str = DEFAULT_DB,
) -> Optional[float]:
    """Возвращает цену ступени по короткой марке."""
    if not mark:
        return None

    key = normalize_step_mark(mark)
    init_step_prices_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT price FROM step_prices WHERE mark = ?", (key,))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()
