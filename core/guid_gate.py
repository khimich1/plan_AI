#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Резолвер GUID для счёта: все 6 видов, ворота assert_invoice_guids.

Точка подключения будущей кнопки «В 1С»: вызвать
``assert_invoice_guids(order_lines, conn)`` до выгрузки счёта.
Одна позиция без GUID блокирует весь счёт.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Optional, Sequence

from core.kp_db_nomenclature import (
    count_nomenclature_by_plate_name,
    lookup_nomenclature_by_plate_name,
)
from core.nomenclature_guid import (
    PRODUCT_KINDS,
    get_by_mark,
    get_guid_for_invoice,
)
from core.pile_line_parser import is_reinforced_pile_mark
from core.pile_price_db import strip_trailing_u_suffix
from core.plate_guid_choice import (
    get_choice,
    normalize_plate_name,
)

INVOICE_STATUSES = frozenset({"auto", "manual"})
_PLATE_KINDS = frozenset({"plate", "plates", "plity", "плита", "плиты"})
_KIND_ALIASES = {
    "piles": "pile",
    "bridge_piles": "bridge_pile",
    "composite_piles": "composite_pile",
    "steps": "stair_step",
    "marches": "stair_flight",
    "stair_flights": "stair_flight",
    "stair_steps": "stair_step",
}

REASON_MISSING = "нет GUID 1С"
REASON_MISSING_U = "нет guid_1c_u для сваи «у»"
REASON_DUP = "выберите GUID (дубль 1С)"
REASON_AMBIGUOUS = "неоднозначный GUID (дубль 1С)"
HINT_CREATE = "заведите карточку в 1С и загрузите отчёт"
HINT_CREATE_U = "заведите карточку «у» в 1С и загрузите отчёт"
HINT_CHOOSE = "откройте «Прайсы и 1С» → Дубли GUID и запомните выбор"


@dataclass(frozen=True)
class OrderLine:
    product_kind: str
    mark: str
    qty: int = 1


@dataclass(frozen=True)
class ReadyItem:
    line: OrderLine
    guid: str


@dataclass(frozen=True)
class MissingItem:
    line: OrderLine
    reason: str
    action_hint: str


@dataclass(frozen=True)
class GateReport:
    ready: tuple[ReadyItem, ...]
    missing: tuple[MissingItem, ...]


class InvoiceGuidBlockError(ValueError):
    """Счёт заблокирован: хотя бы у одной позиции нет GUID."""

    def __init__(self, report: GateReport) -> None:
        self.report = report
        details = "; ".join(
            f"{item.line.mark}: {item.reason}. {item.action_hint}"
            for item in report.missing
        )
        super().__init__(f"Счёт невозможен без GUID. Что завести: {details}")


def _normalize_kind(product_kind: str) -> str:
    raw = str(product_kind or "").strip().lower()
    if raw in _PLATE_KINDS:
        return "plate"
    return _KIND_ALIASES.get(raw, raw)


def _plate_name_candidates(plate_name: str) -> list[str]:
    names: list[str] = []
    raw = str(plate_name or "").strip()
    if raw:
        names.append(raw)
    try:
        compact = normalize_plate_name(plate_name)
    except ValueError:
        compact = ""
    if compact and compact not in names:
        names.append(compact)
    return names


def _lookup_plate(conn: sqlite3.Connection, plate_name: str) -> tuple[Optional[str], Optional[str], int]:
    cur = conn.cursor()
    count = 0
    matched_name = plate_name
    for candidate in _plate_name_candidates(plate_name):
        n = count_nomenclature_by_plate_name(candidate, cur)
        if n:
            count = n
            matched_name = candidate
            break
    if count > 1:
        return None, None, count

    canonical = None
    guid = None
    for candidate in _plate_name_candidates(matched_name):
        canonical, guid, _match = lookup_nomenclature_by_plate_name(candidate, cur)
        if guid:
            break
    if guid and canonical:
        named_count = count_nomenclature_by_plate_name(canonical, cur)
        if named_count > 1:
            return None, canonical, named_count
        count = max(count, named_count or 1)
    return guid, canonical, count


def _resolve_plate(line: OrderLine, conn: sqlite3.Connection) -> ReadyItem | MissingItem:
    choice = get_choice(conn, line.mark)
    if choice is not None:
        return ReadyItem(line=line, guid=choice.chosen_guid)

    guid, _canonical, count = _lookup_plate(conn, line.mark)
    if count > 1:
        return MissingItem(line=line, reason=REASON_DUP, action_hint=HINT_CHOOSE)
    if not guid:
        return MissingItem(line=line, reason=REASON_MISSING, action_hint=HINT_CREATE)
    return ReadyItem(line=line, guid=str(guid))


def _resolve_non_plate(line: OrderLine, conn: sqlite3.Connection) -> ReadyItem | MissingItem:
    kind = _normalize_kind(line.product_kind)
    lookup_mark = line.mark
    reinforced = False
    if kind == "pile":
        reinforced = is_reinforced_pile_mark(line.mark)
        if reinforced:
            lookup_mark = strip_trailing_u_suffix(line.mark) or line.mark

    row = get_by_mark(conn, kind, lookup_mark)
    if row is None:
        return MissingItem(line=line, reason=REASON_MISSING, action_hint=HINT_CREATE)
    if row.match_status == "ambiguous":
        return MissingItem(line=line, reason=REASON_AMBIGUOUS, action_hint=HINT_CHOOSE)
    if row.match_status not in INVOICE_STATUSES:
        return MissingItem(line=line, reason=REASON_MISSING, action_hint=HINT_CREATE)

    guid = get_guid_for_invoice(conn, kind, lookup_mark, reinforced=reinforced)
    if not guid:
        if reinforced:
            return MissingItem(line=line, reason=REASON_MISSING_U, action_hint=HINT_CREATE_U)
        return MissingItem(line=line, reason=REASON_MISSING, action_hint=HINT_CREATE)
    return ReadyItem(line=line, guid=guid)


def check_invoice_guids(
    order_lines: Sequence[OrderLine],
    conn: sqlite3.Connection,
) -> GateReport:
    """Возвращает GateReport(ready, missing); missing блокирует счёт целиком."""
    from core.nomenclature_guid import ensure_schema as ensure_guid_schema
    from core.plate_guid_choice import ensure_schema as ensure_choice_schema

    ensure_guid_schema(conn)
    ensure_choice_schema(conn)

    ready: list[ReadyItem] = []
    missing: list[MissingItem] = []
    for line in order_lines:
        kind = _normalize_kind(line.product_kind)
        if kind == "plate":
            result = _resolve_plate(line, conn)
        elif kind in PRODUCT_KINDS:
            result = _resolve_non_plate(line, conn)
        else:
            result = MissingItem(line=line, reason=REASON_MISSING, action_hint=HINT_CREATE)
        if isinstance(result, ReadyItem):
            ready.append(result)
        else:
            missing.append(result)
    return GateReport(ready=tuple(ready), missing=tuple(missing))


def assert_invoice_guids(
    order_lines: Sequence[OrderLine],
    conn: sqlite3.Connection,
) -> GateReport:
    """Ворота счёта. Вставить вызов перед кнопкой «В 1С».

    Если хоть у одной позиции нет GUID — бросает InvoiceGuidBlockError
    со списком «что завести». Custom-позиции без исключений.
    """
    report = check_invoice_guids(order_lines, conn)
    if report.missing:
        raise InvoiceGuidBlockError(report)
    return report


_PRODUCT_TYPE_TO_KIND = {
    "plates": "plate",
    "piles": "pile",
    "bridge_piles": "bridge_pile",
    "composite_piles": "composite_pile",
    "fbs": "fbs",
    "marches": "stair_flight",
    "steps": "stair_step",
}


def order_lines_from_order_data(order_data: Sequence[object]) -> list[OrderLine]:
    """Строки КП/черновика → OrderLine для резолвера."""
    lines: list[OrderLine] = []
    for raw in order_data:
        if isinstance(raw, dict):
            item = raw
        elif hasattr(raw, "model_dump"):
            item = raw.model_dump()
        else:
            continue
        ptype = str(item.get("product_type") or "").strip()
        kind = _PRODUCT_TYPE_TO_KIND.get(ptype)
        if kind is None:
            continue
        mark = str(item.get("mark") or item.get("name") or "").strip()
        if not mark:
            continue
        try:
            qty = int(item.get("qty") or 1)
        except (TypeError, ValueError):
            qty = 1
        lines.append(OrderLine(kind, mark, qty))
    return lines


__all__ = [
    "GateReport",
    "InvoiceGuidBlockError",
    "MissingItem",
    "OrderLine",
    "ReadyItem",
    "assert_invoice_guids",
    "check_invoice_guids",
    "order_lines_from_order_data",
]
