"""Build invalid-width plate lines for the commercial-offer wizard gate."""

from __future__ import annotations

from typing import Any, Iterable

from core.commercial_pricing import lookup_plate_price
from core.exceptions import PriceNotFoundError
from core.factory_width import (
    format_factory_width_label,
    is_factory_width_mm,
    suggest_factory_width_mm,
    width_m_to_mm,
)
from core.unpriced_plate_replacements import (
    _find_source_line,
    _is_plate_order_item,
    _load_code_from_item,
)

WIDE_WIDTH_MM = 1200


def _lookup_replacement_price(
    *,
    length_m: float,
    width_mm: int,
    load_class: int,
    db_path: str,
) -> float | None:
    try:
        price = lookup_plate_price(
            length_m,
            width_mm / 1000.0,
            load_class,
            db_path=db_path,
        )
    except PriceNotFoundError:
        return None
    if price <= 0:
        return None
    return float(price)


def _skip_wide_keys(skip_wide_lines: Iterable[str] | None) -> set[str]:
    return {str(line).strip() for line in (skip_wide_lines or []) if str(line).strip()}


def build_invalid_width_lines(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
    normalized_lines: list[str] | None = None,
    skip_wide_lines: Iterable[str] | None = None,
) -> list[dict[str, Any]]:
    """Serialize plate positions whose parsed width is ≤12 dm and outside factory cuts."""
    lines_src = [str(line).strip() for line in (normalized_lines or []) if str(line).strip()]
    wide_skip = _skip_wide_keys(skip_wide_lines)
    result: list[dict[str, Any]] = []
    for item in order_data or []:
        if not isinstance(item, dict) or not _is_plate_order_item(item):
            continue
        try:
            width_m = float(item["width_m"])
        except (TypeError, ValueError, KeyError):
            continue
        width_mm = width_m_to_mm(width_m)
        if width_mm > WIDE_WIDTH_MM:
            continue
        name = str(item.get("name", "") or "").strip()
        source_line = _find_source_line(item, lines_src)
        if source_line in wide_skip or name in wide_skip:
            continue
        if is_factory_width_mm(width_mm):
            continue
        length_m = float(item["length_m"])
        load_code = _load_code_from_item(item)
        load_class = int(item.get("load_class") or load_code * 100)
        replacements: list[dict[str, Any]] = []
        for repl_mm in suggest_factory_width_mm(width_mm):
            replacement: dict[str, Any] = {
                "width_mm": repl_mm,
                "width_label": format_factory_width_label(repl_mm),
            }
            price = _lookup_replacement_price(
                length_m=length_m,
                width_mm=repl_mm,
                load_class=load_class,
                db_path=db_path,
            )
            if price is not None:
                replacement["price"] = price
            replacements.append(replacement)
        result.append(
            {
                "id": f"invalid-width-{len(result) + 1}",
                "name": name,
                "line": source_line or name,
                "qty": int(item.get("qty", 1) or 1),
                "length_m": length_m,
                "width_m": width_m,
                "width_mm": width_mm,
                "load_class": load_class,
                "replacements": replacements,
            }
        )
    return result
