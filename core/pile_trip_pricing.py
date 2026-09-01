#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Гибридный расчёт авторейсов свай и мостовых свай для КП.

Полные машины — floor(qty / pcs_per_20t) по марке со справочником.
Остатки известных марок — ceil(Σ кг / 19800). Марки без нормы — ручное N
(override); пока N нет, доставку свай не считаем (ready=False, total_trips=0).
"""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Callable, Mapping, Optional, Sequence

from core.pile_catalog import PileCatalogEntry, normalize_pile_mark_key

PILE_REMAINDER_TRUCK_CAPACITY_KG: float = 19800.0
PILE_TRIP_PRODUCT_TYPES = frozenset({"piles", "bridge_piles"})

CatalogLookup = Callable[[str], Optional[PileCatalogEntry]]


@dataclass(frozen=True)
class PileTripBreakdown:
    full_trips: int
    remainder_kg: float
    remainder_trips: int
    override_trips: int
    pending_marks: tuple[str, ...]
    total_trips: int  # 0 если pending_marks не пуст (доставку не считаем)

    @property
    def ready(self) -> bool:
        return not self.pending_marks


def _line_product_type(item: Mapping[str, Any]) -> str:
    return str(item.get("product_type") or "").strip().lower()


def _line_mark(item: Mapping[str, Any]) -> str:
    return str(item.get("mark") or item.get("name") or "").strip()


def _line_qty(item: Mapping[str, Any]) -> int:
    try:
        qty = int(item.get("qty") or 0)
    except (TypeError, ValueError):
        return 0
    return max(0, qty)


def _has_pcs_norm(entry: Optional[PileCatalogEntry]) -> bool:
    if entry is None:
        return False
    pcs = entry.pcs_per_20t
    return pcs is not None and int(pcs) >= 1


def override_trips_for_mark(
    mark: str,
    overrides: Mapping[str, int] | None,
) -> int | None:
    """N машин для марки без нормы. None = ответ не введён (пусто ≠ 0)."""
    coerced = coerce_pile_trip_overrides(overrides)
    if not coerced:
        return None
    key = normalize_pile_mark_key(mark)
    if not key:
        return None
    for raw_key, raw_n in coerced.items():
        if normalize_pile_mark_key(str(raw_key)) == key:
            return raw_n
    return None


def coerce_pile_trip_overrides(raw: Any) -> dict[str, int]:
    """Нормализовать overrides из dict/JSON. Пустой ввод → {}. Явный 0 сохраняется."""
    if raw is None:
        return {}
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return {}
        import json

        try:
            raw = json.loads(text)
        except json.JSONDecodeError:
            return {}
    if not isinstance(raw, Mapping):
        return {}
    out: dict[str, int] = {}
    for key, value in raw.items():
        mark = str(key).strip()
        if not mark:
            continue
        try:
            n = int(value)
        except (TypeError, ValueError):
            continue
        if n < 0:
            continue
        out[mark] = n
    return out


def dumps_pile_trip_overrides(raw: Any) -> str | None:
    data = coerce_pile_trip_overrides(raw)
    if not data:
        return None
    import json

    return json.dumps(data, ensure_ascii=False)


def compute_pile_trips(
    lines: Sequence[Mapping[str, Any]],
    overrides: Mapping[str, int] | None,
    catalog_lookup: CatalogLookup,
) -> PileTripBreakdown:
    """Рейсы свай по строкам КП. Плиты/ФБС/ЛС/ЛМ игнорируются."""
    known_qty: dict[str, int] = defaultdict(int)
    known_entry: dict[str, PileCatalogEntry] = {}
    pending_display: dict[str, str] = {}
    override_by_key: dict[str, int] = {}

    for item in lines:
        if _line_product_type(item) not in PILE_TRIP_PRODUCT_TYPES:
            continue
        mark = _line_mark(item)
        qty = _line_qty(item)
        if not mark or qty <= 0:
            continue
        entry = catalog_lookup(mark)
        if _has_pcs_norm(entry):
            assert entry is not None
            catalog_key = entry.mark
            known_qty[catalog_key] += qty
            known_entry[catalog_key] = entry
            continue

        override_key = normalize_pile_mark_key(mark)
        n_manual = override_trips_for_mark(mark, overrides)
        if n_manual is None:
            pending_display.setdefault(override_key, mark)
            continue
        override_by_key[override_key] = n_manual

    full_trips = 0
    remainder_kg = 0.0
    for catalog_key, qty in known_qty.items():
        entry = known_entry[catalog_key]
        pcs = int(entry.pcs_per_20t or 0)
        full = qty // pcs
        rem_pcs = qty % pcs
        full_trips += full
        if rem_pcs:
            remainder_kg += rem_pcs * float(entry.weight_kg)

    if remainder_kg <= 0:
        remainder_trips = 0
        remainder_kg = 0.0
    else:
        remainder_trips = int(math.ceil(remainder_kg / PILE_REMAINDER_TRUCK_CAPACITY_KG))

    override_trips = sum(override_by_key.values())
    pending_marks = tuple(pending_display.values())
    if pending_marks:
        total_trips = 0
    else:
        total_trips = full_trips + remainder_trips + override_trips

    return PileTripBreakdown(
        full_trips=full_trips,
        remainder_kg=remainder_kg,
        remainder_trips=remainder_trips,
        override_trips=override_trips,
        pending_marks=pending_marks,
        total_trips=total_trips,
    )
