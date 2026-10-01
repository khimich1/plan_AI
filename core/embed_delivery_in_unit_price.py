"""Bake delivery totals into per-piece prices after product discount.

Pure arithmetic: no I/O, no Excel, no ``calculate_total_cost``.
"""

from __future__ import annotations

from dataclasses import dataclass

PLATE_POOL = frozenset({"plates"})
PILE_POOL = frozenset({"piles", "bridge_piles"})
FBS_POOL = frozenset({"fbs", "steps", "marches"})

_KIND_ALIASES = {
    "pile": "piles",
    "bridge_pile": "bridge_piles",
    "step": "steps",
    "march": "marches",
}
_LINE_DELIVERY_ALIASES = {
    **_KIND_ALIASES,
    "plate": "plates",
    "stair_step": "steps",
    "stair_flight": "marches",
}


@dataclass(frozen=True)
class EmbeddedLine:
    index: int
    qty: int
    discounted_unit: float
    surcharge_unit: float
    unit_price: float
    line_sum: float


@dataclass(frozen=True)
class EmbedDeliveryResult:
    lines: list[EmbeddedLine]
    embedded_plate: bool
    embedded_pile: bool


def _normalize_product_type(raw: str) -> str:
    key = str(raw or "").strip().lower()
    if not key:
        return "plates"
    return _KIND_ALIASES.get(key, key)


def _to_kopecks(amount: float) -> int:
    return int(round(float(amount) * 100))


def _pool_line_indices(
    product_types: list[str],
    qtys: list[int],
    pool: frozenset[str],
) -> list[int]:
    return [
        i
        for i, (pt, qty) in enumerate(zip(product_types, qtys))
        if pt in pool and qty > 0
    ]


def _allocate_pool_delivery(
    *,
    indices: list[int],
    qtys: list[int],
    delivery_kopecks: int,
) -> dict[int, int]:
    """Return extra delivery kopecks per line index. Empty if pool is not embeddable."""
    total_qty = sum(qtys[i] for i in indices)
    if delivery_kopecks <= 0 or total_qty <= 0:
        return {}
    per_piece_k, extra_k = divmod(delivery_kopecks, total_qty)
    remaining_extra = extra_k
    allocated: dict[int, int] = {}
    for i in indices:
        qty = qtys[i]
        extras_on_line = min(qty, remaining_extra)
        remaining_extra -= extras_on_line
        allocated[i] = qty * per_piece_k + extras_on_line
    return allocated


def _line_delivery_type(raw: str) -> str:
    key = str(raw or "").strip().lower()
    if not key:
        return "plates"
    return _LINE_DELIVERY_ALIASES.get(key, key)


def _positive_length_key(raw: int | None) -> int | None:
    if raw is None:
        return None
    try:
        key = int(raw)
    except (TypeError, ValueError):
        return None
    return key if key > 0 else None


def _add_kopecks(allocated: list[int], pool: dict[int, int]) -> None:
    for index, kopecks in pool.items():
        allocated[index] += kopecks


def allocate_line_delivery_kopecks(
    *,
    qty_by_index: list[int],
    product_type_by_index: list[str],
    plate_delivery_total: float = 0.0,
    pile_delivery_total: float = 0.0,
    fbs_delivery_total: float = 0.0,
    long_pile_delivery_by_length: dict[int, float] | None = None,
    length_key_by_index: list[int | None] | None = None,
) -> list[int]:
    """Delivery kopecks per line. Discount is not an input.

    Pools do not mix: plates, short piles with bridge piles, FBS/steps/marches,
    and each long-pile length. The extra kopeck goes to the first pieces.
    """
    count = len(qty_by_index)
    if len(product_type_by_index) != count:
        raise ValueError("qty and product_type lists must be the same length")
    length_keys = list(length_key_by_index) if length_key_by_index is not None else [None] * count
    if len(length_keys) != count:
        raise ValueError("length_key list must match qty")

    qtys = [max(0, int(qty)) for qty in qty_by_index]
    types = [_line_delivery_type(raw) for raw in product_type_by_index]
    long_keys = [
        _positive_length_key(length_keys[index]) if types[index] in PILE_POOL else None
        for index in range(count)
    ]
    allocated = [0] * count
    _add_kopecks(
        allocated,
        _allocate_pool_delivery(
            indices=_pool_line_indices(types, qtys, PLATE_POOL),
            qtys=qtys,
            delivery_kopecks=_to_kopecks(plate_delivery_total),
        ),
    )
    short_pile_indices = [
        index
        for index in _pool_line_indices(types, qtys, PILE_POOL)
        if long_keys[index] is None
    ]
    _add_kopecks(
        allocated,
        _allocate_pool_delivery(
            indices=short_pile_indices,
            qtys=qtys,
            delivery_kopecks=_to_kopecks(pile_delivery_total),
        ),
    )
    _add_kopecks(
        allocated,
        _allocate_pool_delivery(
            indices=_pool_line_indices(types, qtys, FBS_POOL),
            qtys=qtys,
            delivery_kopecks=_to_kopecks(fbs_delivery_total),
        ),
    )
    _add_long_pile_kopecks(allocated, qtys, long_keys, long_pile_delivery_by_length or {})
    return allocated


def _add_long_pile_kopecks(
    allocated: list[int],
    qtys: list[int],
    long_keys: list[int | None],
    totals: dict[int, float],
) -> None:
    by_length: dict[int, list[int]] = {}
    for index, key in enumerate(long_keys):
        if key is None or qtys[index] <= 0:
            continue
        by_length.setdefault(key, []).append(index)
    for key, indices in by_length.items():
        _add_kopecks(
            allocated,
            _allocate_pool_delivery(
                indices=indices,
                qtys=qtys,
                delivery_kopecks=_to_kopecks(float(totals.get(key, 0.0) or 0.0)),
            ),
        )


def embed_delivery_in_unit_prices(
    *,
    qty_by_index: list[int],
    product_type_by_index: list[str],
    unit_price_by_index: list[float],
    discount_percent: float,
    plate_delivery_total: float,
    pile_delivery_total: float,
) -> EmbedDeliveryResult:
    """Discount the product, then add this pool's delivery share per piece."""
    n = len(qty_by_index)
    if len(product_type_by_index) != n or len(unit_price_by_index) != n:
        raise ValueError("qty, product_type and unit_price lists must be the same length")

    qtys = [max(0, int(q)) for q in qty_by_index]
    types = [_normalize_product_type(pt) for pt in product_type_by_index]
    discount = max(0.0, min(100.0, float(discount_percent)))
    factor = 1.0 - discount / 100.0

    discounted_k = [
        int(round(float(price) * factor * 100)) for price in unit_price_by_index
    ]

    plate_k = _to_kopecks(plate_delivery_total)
    pile_k = _to_kopecks(pile_delivery_total)
    plate_indices = _pool_line_indices(types, qtys, PLATE_POOL)
    pile_indices = _pool_line_indices(types, qtys, PILE_POOL)

    plate_alloc = _allocate_pool_delivery(
        indices=plate_indices, qtys=qtys, delivery_kopecks=plate_k
    )
    pile_alloc = _allocate_pool_delivery(
        indices=pile_indices, qtys=qtys, delivery_kopecks=pile_k
    )
    embedded_plate = bool(plate_alloc)
    embedded_pile = bool(pile_alloc)

    lines: list[EmbeddedLine] = []
    for i in range(n):
        qty = qtys[i]
        disc_k = discounted_k[i]
        delivery_k = plate_alloc.get(i, 0) + pile_alloc.get(i, 0)
        line_sum_k = disc_k * qty + delivery_k
        discounted_unit = disc_k / 100.0
        if qty > 0:
            surcharge_unit = delivery_k / qty / 100.0
            unit_price = round(line_sum_k / qty) / 100.0
        else:
            surcharge_unit = 0.0
            unit_price = discounted_unit
        lines.append(
            EmbeddedLine(
                index=i,
                qty=qty,
                discounted_unit=discounted_unit,
                surcharge_unit=surcharge_unit,
                unit_price=unit_price,
                line_sum=line_sum_k / 100.0,
            )
        )

    return EmbedDeliveryResult(
        lines=lines,
        embedded_plate=embedded_plate,
        embedded_pile=embedded_pile,
    )
