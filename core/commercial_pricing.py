from __future__ import annotations

import logging
import sqlite3
from typing import Any

from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.price_db import length_m_to_price_length_dm

_log = logging.getLogger(__name__)


def lookup_plate_price(
    length_m: float,
    width_m: float,
    load_class: int = 800,
    *,
    db_path: str,
) -> float:
    """Return plate unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    _ = width_m  # width is part of the public contract for callers / future pricing rules
    try:
        length_dm = length_m_to_price_length_dm(length_m)
        load_code = int(load_class) // 100
        con = sqlite3.connect(db_path)
        try:
            cur = con.cursor()
            row = cur.execute(
                "SELECT price FROM prices WHERE length_dm = ? AND load_code = ?",
                (length_dm, load_code),
            ).fetchone()
        finally:
            con.close()
        if row:
            return float(row[0])
        raise PriceNotFoundError(
            f"Цена не найдена: length_dm={length_dm}, load_code={load_code}"
        )
    except PriceNotFoundError:
        raise
    except Exception as exc:
        _log.warning(
            "Ошибка получения цены (length_m=%s, load_class=%s): %s",
            length_m,
            load_class,
            exc,
            exc_info=True,
        )
        raise PriceNotFoundError(
            f"Ошибка получения цены для length_m={length_m}, load_class={load_class}"
        ) from exc


def position_label(item: dict[str, Any]) -> str:
    name = str(item.get("name", "") or "").strip()
    if name:
        return name
    length_m = item.get("length_m", 0)
    width_m = item.get("width_m", 0)
    load_class = int(item.get("load_class", 800) or 800)
    return f"ПБ {length_m}-{width_m}-{load_class // 100}п"


def collect_unpriced_positions(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
) -> list[str]:
    unpriced: list[str] = []
    seen: set[str] = set()
    for item in order_data:
        if item.get("unit_price") is not None:
            continue
        length_m = float(item.get("length_m", 0) or 0)
        width_m = float(item.get("width_m", 0) or 0)
        load_class = int(item.get("load_class", 800) or 800)
        try:
            lookup_plate_price(length_m, width_m, load_class, db_path=db_path)
        except PriceNotFoundError:
            label = position_label(item)
            if label not in seen:
                seen.add(label)
                unpriced.append(label)
    return unpriced


def ensure_order_priced(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
) -> None:
    positions = collect_unpriced_positions(order_data, db_path=db_path)
    if positions:
        _log.warning("Непрорасценённые позиции: %s", positions)
        raise UnpricedPlatesError(positions)
