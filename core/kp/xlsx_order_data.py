"""Извлечение цен из сохранённого XLSX коммерческого предложения."""

from __future__ import annotations

import io
from typing import Any

import pandas as pd


def parse_discounted_prices_from_kp_xlsx(xlsx_bytes: bytes) -> dict[str, float]:
    """Возвращает {наименование: цена со скидкой} из листа «КП»."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name="КП", header=None)
    header_row: int | None = None
    name_col = qty_col = price_col = None

    for row_idx in range(len(df)):
        row = [str(value).strip() if pd.notna(value) else "" for value in df.iloc[row_idx]]
        if "Наименование" in row:
            header_row = row_idx
            name_col = row.index("Наименование")
            qty_col = row.index("Кол-во") if "Кол-во" in row else None
            price_col = row.index("Цена") if "Цена" in row else None
            break

    if header_row is None or name_col is None or price_col is None:
        return {}

    prices: dict[str, float] = {}
    for row_idx in range(header_row + 1, len(df)):
        name_raw = df.iat[row_idx, name_col]
        if pd.isna(name_raw):
            continue
        name = str(name_raw).strip()
        if not name or name.lower().startswith("услуга по доставке"):
            continue
        if qty_col is not None:
            qty_raw = df.iat[row_idx, qty_col]
            if pd.isna(qty_raw) or float(qty_raw or 0) <= 0:
                continue
        price_raw = df.iat[row_idx, price_col]
        price = 0.0 if pd.isna(price_raw) else float(price_raw)
        prices[name] = price
    return prices


def enrich_order_data_prices_from_xlsx(
    order_data: list[dict[str, Any]],
    xlsx_bytes: bytes,
    discount_percent: float = 0,
) -> list[dict[str, Any]]:
    """Дополняет order_data ценами из сохранённого XLSX (для выполненных КП)."""
    discounted_by_name = parse_discounted_prices_from_kp_xlsx(xlsx_bytes)
    if not discounted_by_name:
        return order_data

    factor = 1.0 - (float(discount_percent or 0) / 100.0)
    if factor <= 0:
        factor = 1.0

    enriched: list[dict[str, Any]] = []
    for item in order_data:
        row = dict(item)
        if row.get("unit_price") is not None:
            enriched.append(row)
            continue
        name = str(row.get("name") or "").strip()
        discounted_price = discounted_by_name.get(name)
        if discounted_price is None:
            enriched.append(row)
            continue
        row["unit_price"] = discounted_price / factor if discounted_price > 0 else 0.0
        enriched.append(row)
    return enriched
