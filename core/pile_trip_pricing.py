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

from core.pile_catalog import (
    PileCatalogEntry,
    normalize_pile_mark_key,
    parse_bridge_pile_geometry,
    parse_pile_mark,
)

PILE_REMAINDER_TRUCK_CAPACITY_KG: float = 19800.0
PILE_TRIP_PRODUCT_TYPES = frozenset({"piles", "bridge_piles"})
LONG_PILE_LENGTH_M_MIN: float = 13.0

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


@dataclass(frozen=True)
class LongPileLengthGroup:
    """Рейсы одной длины > 13 м. Pending не затирает trips соседних длин."""

    length_key: int
    full_trips: int
    remainder_kg: float
    remainder_trips: int
    override_trips: int
    pending_marks: tuple[str, ...]
    trips: int
    qty: int

    @property
    def ready(self) -> bool:
        return not self.pending_marks


@dataclass(frozen=True)
class LongPileSplit:
    short: PileTripBreakdown
    lengths: tuple[LongPileLengthGroup, ...]


def _length_m_for_mark(mark: str, entry: Optional[PileCatalogEntry]) -> float | None:
    if entry is not None and entry.length_m is not None:
        return float(entry.length_m)
    length_m, _section = parse_bridge_pile_geometry(mark)
    if length_m is None:
        length_m, _section = parse_pile_mark(mark)
    return length_m


def _is_long_pile_length(length_m: float | None) -> bool:
    return length_m is not None and length_m > LONG_PILE_LENGTH_M_MIN


def _length_key(length_m: float) -> int:
    return int(round(length_m * 10))


def compute_long_pile_groups(
    lines: Sequence[Mapping[str, Any]],
    overrides: Mapping[str, int] | None,
    catalog_lookup: CatalogLookup,
) -> LongPileSplit:
    """Короткий котёл (≤ 13 м) и группы длиннее 13 м. Остатки групп не смешиваются."""
    short_lines: list[Mapping[str, Any]] = []
    grouped: dict[int, list[Mapping[str, Any]]] = defaultdict(list)

    for item in lines:
        if _line_product_type(item) not in PILE_TRIP_PRODUCT_TYPES:
            continue
        mark = _line_mark(item)
        qty = _line_qty(item)
        if not mark or qty <= 0:
            continue
        entry = catalog_lookup(mark)
        length_m = _length_m_for_mark(mark, entry)
        if _is_long_pile_length(length_m):
            assert length_m is not None
            grouped[_length_key(length_m)].append(item)
        else:
            short_lines.append(item)

    short = compute_pile_trips(short_lines, overrides, catalog_lookup)
    lengths: list[LongPileLengthGroup] = []
    for length_key in sorted(grouped):
        bucket = grouped[length_key]
        breakdown = compute_pile_trips(bucket, overrides, catalog_lookup)
        qty = sum(_line_qty(item) for item in bucket)
        lengths.append(
            LongPileLengthGroup(
                length_key=length_key,
                full_trips=breakdown.full_trips,
                remainder_kg=breakdown.remainder_kg,
                remainder_trips=breakdown.remainder_trips,
                override_trips=breakdown.override_trips,
                pending_marks=breakdown.pending_marks,
                trips=breakdown.total_trips,
                qty=qty,
            )
        )
    return LongPileSplit(short=short, lengths=tuple(lengths))


@dataclass(frozen=True)
class LongPileLengthQuote:
    """Деньги одной длины. Пустой тариф — None, явный 0 — цена введена."""

    length_key: int
    trips: int
    trip_cost: float | None
    pending_marks: tuple[str, ...]
    amount: float
    qty: int

    @property
    def ready(self) -> bool:
        return self.trip_cost is not None and not self.pending_marks


def _trip_cost_from_value(value: Any) -> float | None:
    if isinstance(value, Mapping):
        if "trip_cost" not in value:
            return None
        value = value.get("trip_cost")
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        cost = float(value)
    except (TypeError, ValueError):
        return None
    if cost < 0 or math.isnan(cost) or math.isinf(cost):
        return None
    return cost


def coerce_long_pile_delivery(raw: Any) -> dict[int, float]:
    """Тарифы длин из dict/JSON. Пустой ввод не пишется. Явный 0 сохраняется."""
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
    out: dict[int, float] = {}
    for key, value in raw.items():
        try:
            length_key = int(str(key).strip())
        except (TypeError, ValueError):
            continue
        if length_key <= 0:
            continue
        cost = _trip_cost_from_value(value)
        if cost is None:
            continue
        out[length_key] = cost
    return out


def dumps_long_pile_delivery(raw: Any) -> str | None:
    data = coerce_long_pile_delivery(raw)
    if not data:
        return None
    import json

    return json.dumps(long_pile_delivery_for_metadata(data), ensure_ascii=False)


def long_pile_delivery_for_metadata(raw: Any) -> dict[str, dict[str, float]]:
    """Тарифы для metadata/API: только введённые цены, ключ — строка length_key."""
    data = coerce_long_pile_delivery(raw)
    return {
        str(length_key): {"trip_cost": data[length_key]}
        for length_key in sorted(data)
    }


def quote_long_pile_lengths(
    split: LongPileSplit,
    tariffs: Any,
) -> tuple[LongPileLengthQuote, ...]:
    costs = coerce_long_pile_delivery(tariffs)
    quotes: list[LongPileLengthQuote] = []
    for group in split.lengths:
        trip_cost = costs[group.length_key] if group.length_key in costs else None
        amount = (
            round(float(trip_cost) * group.trips, 2)
            if trip_cost is not None and group.ready
            else 0.0
        )
        quotes.append(
            LongPileLengthQuote(
                length_key=group.length_key,
                trips=group.trips,
                trip_cost=trip_cost,
                pending_marks=group.pending_marks,
                amount=amount,
                qty=group.qty,
            )
        )
    return tuple(quotes)


def format_long_pile_delivery_label(length_key: int) -> str:
    """140 → «Доставка свай 14,0 м», 138 → «Доставка свай 13,8 м»."""
    meters = f"{int(length_key) / 10:.1f}".replace(".", ",")
    return f"Доставка свай {meters} м"
