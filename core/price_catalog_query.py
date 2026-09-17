"""Read-only listing of factory price tables for the KP catalog drawer.

SELECT only: never INSERT/UPDATE/CREATE. Filter ``q`` in Python after C↔С /
space / case normalization. Cap is applied after filtering.
"""

from __future__ import annotations

import re
import sqlite3
from typing import Any

from core.price_db import DEFAULT_DB, _connect

CATALOG_CAP = 2000

_GRADED_TABLES = {
    "piles": "pile_prices",
    "bridge_piles": "bridge_pile_prices",
    "fbs": "fbs_prices",
    "marches": "march_prices",
}
_STEPS_TABLE = "step_prices"
_ALLOWED_TYPES = frozenset({*_GRADED_TABLES, "steps"})

_UNKNOWN_TYPE = "Неизвестный тип продукции."
_CAP_EXCEEDED = "Слишком много позиций. Уточните поиск."


def normalize_catalog_search_mark(mark: str) -> str:
    """C↔С, T↔Т, B↔В, collapse spaces, case-fold for substring search."""
    text = str(mark or "").strip().upper().replace("Ё", "Е")
    text = text.replace("С", "C").replace("В", "B").replace("Т", "T")
    return re.sub(r"\s+", "", text)


def list_price_catalog(
    product_type: str,
    q: str = "",
    db_path: str | None = None,
) -> list[dict[str, Any]]:
    """Return priced catalog rows for ``product_type`` matching optional ``q``.

    Raises:
        ValueError: unknown type, or more than ``CATALOG_CAP`` matches.
    """
    kind = str(product_type or "").strip().lower()
    if kind not in _ALLOWED_TYPES:
        raise ValueError(_UNKNOWN_TYPE)

    path = db_path if db_path is not None else DEFAULT_DB
    needle = normalize_catalog_search_mark(q)
    rows = _select_priced_rows(kind, path)
    if needle:
        rows = [
            row
            for row in rows
            if needle in normalize_catalog_search_mark(str(row["mark"]))
        ]
    if len(rows) > CATALOG_CAP:
        raise ValueError(_CAP_EXCEEDED)
    return rows


def _select_priced_rows(product_type: str, db_path: str) -> list[dict[str, Any]]:
    if product_type == "steps":
        sql = "SELECT mark, price FROM step_prices WHERE price > 0 ORDER BY mark"
        has_grade = False
    else:
        table = _GRADED_TABLES[product_type]
        sql = (
            f"SELECT mark, concrete_grade, price FROM {table} "
            "WHERE price > 0 ORDER BY mark, concrete_grade"
        )
        has_grade = True

    conn = _connect(db_path)
    try:
        try:
            fetched = conn.execute(sql).fetchall()
        except sqlite3.OperationalError:
            return []
    finally:
        conn.close()

    items: list[dict[str, Any]] = []
    for row in fetched:
        mark = str(row[0] or "").strip()
        if not mark:
            continue
        if has_grade:
            grade = str(row[1] or "").strip() or None
            price = float(row[2])
        else:
            grade = None
            price = float(row[1])
        if price <= 0:
            continue
        items.append({"mark": mark, "concrete_grade": grade, "price": price})
    return items
