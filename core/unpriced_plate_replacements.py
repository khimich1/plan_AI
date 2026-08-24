"""Helpers for unpriced plate load-class replacements in commercial offers."""

from __future__ import annotations

import math
import re
import sqlite3
from typing import Any

from core.config_and_data import (
    format_reinforcement_from_load_code,
    parse_load_code_from_name,
    parse_name_to_sizes,
)
from core.price_db import length_m_to_price_length_dm

# Standard commercial load codes in the price DB (descending).
STANDARD_LOAD_CODES: tuple[int, ...] = (12, 10, 8, 6)

_LOAD_SUFFIX_RE = re.compile(r"-\s*\d+(?:[,.]\d+)?\s*п\b", re.IGNORECASE)


def _resolve_length_dm(*, length_m: float | None, length_dm: int | None) -> int:
    if length_dm is not None:
        return int(length_dm)
    if length_m is None:
        raise ValueError("Нужно указать length_m или length_dm")
    return length_m_to_price_length_dm(float(length_m))


def _normalize_current_load_code(load_code: float | int) -> int:
    """Floor load for price-DB matching (12.5 → 12), same idea as get_price."""
    return int(math.floor(float(load_code)))


def list_lower_load_replacements(
    *,
    load_code: float | int,
    db_path: str,
    length_m: float | None = None,
    length_dm: int | None = None,
) -> list[dict[str, float | int]]:
    """Return cheaper/lower load options with price > 0 for the same length.

    Candidates are STANDARD_LOAD_CODES strictly below the current load_code,
    sorted by load_code descending (nearest lower first).
    """
    length_key = _resolve_length_dm(length_m=length_m, length_dm=length_dm)
    current = _normalize_current_load_code(load_code)
    candidates = [code for code in STANDARD_LOAD_CODES if code < current]
    if not candidates:
        return []

    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        results: list[dict[str, float | int]] = []
        for code in candidates:
            row = cur.execute(
                "SELECT price FROM prices WHERE length_dm = ? AND load_code = ?",
                (length_key, code),
            ).fetchone()
            if row is None:
                continue
            price = float(row[0])
            if price <= 0:
                continue
            results.append({"load_code": code, "price": price})
        results.sort(key=lambda item: int(item["load_code"]), reverse=True)
        return results
    finally:
        con.close()


def rewrite_plate_line_load(line: str, new_load_code: int | float) -> str:
    """Replace the load suffix in a plate order line (handles 12п / 12,5п).

    Uses the last ``-Nп`` / ``-N,Mп`` match so width tokens like ``-12-`` are kept.
    """
    text = str(line or "")
    matches = list(_LOAD_SUFFIX_RE.finditer(text))
    if not matches:
        raise ValueError(f"Не удалось найти суффикс нагрузки в строке: {line!r}")
    match = matches[-1]
    reinforcement = format_reinforcement_from_load_code(new_load_code)
    return f"{text[: match.start()]}-{reinforcement}{text[match.end() :]}"


def _is_plate_order_item(item: dict[str, Any]) -> bool:
    product_kind = str(item.get("product_kind", "") or "").strip().lower()
    if product_kind and product_kind != "plate":
        return False
    product_type = str(item.get("product_type", "") or "").strip().lower()
    if product_type and product_type != "plates":
        return False
    return "length_m" in item and "width_m" in item


def _load_code_from_item(item: dict[str, Any]) -> int:
    if item.get("load_class") is not None:
        try:
            return int(float(item["load_class"])) // 100
        except (TypeError, ValueError):
            pass
    return parse_load_code_from_name(str(item.get("name", "") or ""), default=8)


def _dims_match(
    *,
    length_m: float,
    width_m: float,
    load_code: int,
    line: str,
) -> bool:
    parsed_length, parsed_width = parse_name_to_sizes(line)
    if parsed_length is None or parsed_width is None:
        return False
    if abs(float(parsed_length) - float(length_m)) >= 0.01:
        return False
    if abs(float(parsed_width) - float(width_m)) >= 0.01:
        return False
    line_load = parse_load_code_from_name(line, default=-1)
    # 12.5п → 13 via parse_load_code_from_name; price floor is 12
    line_load_for_match = 12 if line_load in (12, 13) else line_load
    item_load_for_match = 12 if load_code in (12, 13) else load_code
    return line_load_for_match == item_load_for_match


def _find_source_line(
    item: dict[str, Any],
    normalized_lines: list[str],
) -> str:
    length_m = float(item["length_m"])
    width_m = float(item["width_m"])
    load_code = _load_code_from_item(item)
    for line in normalized_lines:
        if _dims_match(length_m=length_m, width_m=width_m, load_code=load_code, line=line):
            return line
    name = str(item.get("name", "") or "").strip()
    return name


def build_unpriced_plate_lines(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
    normalized_lines: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Serialize plate positions with ``unit_price is None`` plus replacement options."""
    lines_src = [str(line).strip() for line in (normalized_lines or []) if str(line).strip()]
    result: list[dict[str, Any]] = []
    for item in order_data or []:
        if not isinstance(item, dict) or not _is_plate_order_item(item):
            continue
        if item.get("unit_price") is not None:
            continue
        length_m = float(item["length_m"])
        width_m = float(item["width_m"])
        load_code = _load_code_from_item(item)
        load_class = int(item.get("load_class") or load_code * 100)
        name = str(item.get("name", "") or "").strip()
        source_line = _find_source_line(item, lines_src)
        replacements = list_lower_load_replacements(
            length_m=length_m,
            load_code=load_code,
            db_path=db_path,
        )
        result.append(
            {
                "id": f"unpriced-{len(result) + 1}",
                "name": name,
                "line": source_line or name,
                "qty": int(item.get("qty", 1) or 1),
                "length_m": length_m,
                "width_m": width_m,
                "load_class": load_class,
                "replacements": replacements,
            }
        )
    return result
