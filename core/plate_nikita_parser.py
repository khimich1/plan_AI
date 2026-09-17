"""Plate prices from sheet «Прайс», Nikita column only (not M400/M500, not loads 16/21)."""

from __future__ import annotations

import os
import re
from typing import List, Optional, Tuple

try:
    import pandas as pd
except Exception:  # pragma: no cover - pandas is a project dependency
    pd = None

PlateNikitaRow = Tuple[int, int, float]

NIKITA_HEADER = "цены по прайсу, который прислал никита"
ALLOWED_LOADS = {6, 8, 10, 12}
SKIP_LOADS = {16, 21}

_DEFAULT_SHEET_NAMES = ("прайс", "price")
_PLATE_NAME_RE = re.compile(
    r"П[БBKК]\s*(\d+(?:[.,]\d+)?)\s*[-–]\s*\d+(?:[.,]\d+)?\s*[-–]\s*(\d+(?:[.,]\d+)?)",
    re.IGNORECASE,
)


def _normalize_header(value) -> str:
    text = str(value or "").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def find_price_sheet_name(xlsx_path: str, preferred_sheet: Optional[str] = "Прайс") -> Optional[str]:
    if pd is None or not os.path.exists(xlsx_path):
        return None
    sheets = pd.ExcelFile(xlsx_path).sheet_names
    if preferred_sheet:
        for sheet in sheets:
            if sheet.strip().lower() == preferred_sheet.strip().lower():
                return sheet
    for sheet in sheets:
        if sheet.strip().lower() in _DEFAULT_SHEET_NAMES:
            return sheet
    return None


def has_nikita_column(xlsx_path: str) -> bool:
    return _find_nikita_layout(xlsx_path) is not None


def _find_nikita_layout(
    xlsx_path: str,
) -> Optional[tuple[object, int, int, int]]:
    """Return (dataframe, name_col_idx, nikita_col_idx, header_row_idx) or None."""
    if pd is None or not os.path.exists(xlsx_path):
        return None
    sheet_name = find_price_sheet_name(xlsx_path)
    if sheet_name is None:
        return None
    raw_df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)
    header_row_idx = None
    nikita_col_idx = None
    for idx, row in raw_df.iterrows():
        for col_idx, value in enumerate(row.tolist()):
            if NIKITA_HEADER in _normalize_header(value):
                header_row_idx = int(idx)
                nikita_col_idx = int(col_idx)
                break
        if header_row_idx is not None:
            break
    if header_row_idx is None or nikita_col_idx is None:
        return None

    header = raw_df.iloc[header_row_idx]
    name_col_idx = next(
        (
            i
            for i, value in enumerate(header.tolist())
            if "наимен" in _normalize_header(value)
        ),
        None,
    )
    if name_col_idx is None:
        name_col_idx = 1 if len(header) > 1 else 0
    return raw_df, name_col_idx, nikita_col_idx, header_row_idx


def _map_load_code(raw: float) -> Optional[int]:
    if abs(raw - 12.5) < 0.05:
        return 12
    if raw in SKIP_LOADS or abs(raw - 16) < 0.05 or abs(raw - 21) < 0.05:
        return None
    code = int(round(raw))
    if abs(raw - code) > 0.05:
        return None
    if code in ALLOWED_LOADS:
        return code
    return None


def _parse_plate_name(name: str) -> Optional[tuple[int, int]]:
    match = _PLATE_NAME_RE.search(str(name or ""))
    if not match:
        return None
    length_val = float(match.group(1).replace(",", "."))
    load_val = float(match.group(2).replace(",", "."))
    load_code = _map_load_code(load_val)
    if load_code is None:
        return None
    return int(round(length_val)), load_code


def _cell_price(value) -> Optional[float]:
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        price = float(text.replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None
    if price <= 0:
        return None
    return price


def parse_plate_nikita_rows(xlsx_path: str) -> List[PlateNikitaRow]:
    """Лист «Прайс». Имя ПБ {length}-12-{load}. Цена — колонка Никиты. price > 0."""
    found = _find_nikita_layout(xlsx_path)
    if found is None:
        return []
    raw_df, name_col_idx, nikita_col_idx, header_row_idx = found
    rows: List[PlateNikitaRow] = []
    for _, row in raw_df.iloc[header_row_idx + 1 :].iterrows():
        name = str(row.iloc[name_col_idx] if name_col_idx < len(row) else "").strip()
        parsed = _parse_plate_name(name)
        if parsed is None:
            continue
        if nikita_col_idx >= len(row):
            continue
        price = _cell_price(row.iloc[nikita_col_idx])
        if price is None:
            continue
        length_dm, load_code = parsed
        rows.append((length_dm, load_code, price))
    return rows
