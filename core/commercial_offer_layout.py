"""Shared commercial-offer table layout: mono (R3) vs unified multi/append."""

from __future__ import annotations

from typing import Any, Mapping, Optional, Sequence

from core.commercial_pricing import (
    is_bridge_pile_order,
    is_fbs_order,
    is_march_order,
    is_pile_order,
    is_step_order,
)

UNIFIED_HEADERS: list[str] = ["№", "Тип", "Наименование", "Кол-во", "Цена", "Сумма"]
MONO_PLATES_HEADERS: list[str] = [
    "№",
    "Наименование",
    "Кол-во",
    "Ед.",
    "Вес(кг)",
    "Цена",
    "Сумма",
]
MONO_PILES_HEADERS: list[str] = [
    "№",
    "Наименование",
    "Класс бетона",
    "Кол-во",
    "Цена",
    "Сумма",
]
MONO_STEPS_HEADERS: list[str] = ["№", "Наименование", "Кол-во", "Цена", "Сумма"]

_PRODUCT_TYPE_LABELS: dict[str, str] = {
    "plates": "Плиты",
    "piles": "Сваи",
    "steps": "Ступени",
    "marches": "Марши",
    "bridge_piles": "Мостовые сваи",
    "fbs": "ФБС",
}

_PRODUCT_KIND_TO_TYPE: dict[str, str] = {
    "pile": "piles",
    "step": "steps",
    "march": "marches",
    "bridge_pile": "bridge_piles",
    "fbs": "fbs",
}


def line_product_type(item: Mapping[str, Any]) -> str:
    """Resolve product_type for a line (explicit → product_kind → plates)."""
    explicit = str(item.get("product_type") or "").strip().lower()
    if explicit in _PRODUCT_TYPE_LABELS:
        return explicit
    kind = str(item.get("product_kind") or "").strip().lower()
    mapped = _PRODUCT_KIND_TO_TYPE.get(kind)
    if mapped:
        return mapped
    return "plates"


def product_type_label(product_type: str) -> str:
    """Human-readable type for the «Тип» column."""
    key = str(product_type or "").strip().lower()
    return _PRODUCT_TYPE_LABELS.get(key, "Плиты")


def is_unified_commercial_document(
    order_data: Sequence[Mapping[str, Any]],
    append_batches: Optional[Sequence[Mapping[str, Any]]] = None,
) -> bool:
    """True for multi-type or multi-append documents (unified columns).

    Mono one-shot (0–1 append batch, single product_type) stays classic (R3).
    """
    if not order_data:
        return False

    distinct_types = {
        str(item.get("product_type") or "").strip().lower()
        for item in order_data
        if str(item.get("product_type") or "").strip()
    }
    if len(distinct_types) >= 2:
        return True

    if append_batches is not None and len(append_batches) > 1:
        return True

    batch_ids = {
        str(item.get("append_batch_id") or "").strip()
        for item in order_data
        if str(item.get("append_batch_id") or "").strip()
    }
    if len(batch_ids) >= 2:
        return True

    return False


def commercial_offer_table_headers(
    order_data: Sequence[Mapping[str, Any]],
    append_batches: Optional[Sequence[Mapping[str, Any]]] = None,
) -> list[str]:
    """Column headers for PDF/XLSX: unified multi or classic mono layout."""
    if is_unified_commercial_document(order_data, append_batches=append_batches):
        return list(UNIFIED_HEADERS)

    order_list = list(order_data)
    if (
        is_pile_order(order_list)
        or is_bridge_pile_order(order_list)
        or is_fbs_order(order_list)
        or is_march_order(order_list)
    ):
        return list(MONO_PILES_HEADERS)
    if is_step_order(order_list):
        return list(MONO_STEPS_HEADERS)
    return list(MONO_PLATES_HEADERS)
