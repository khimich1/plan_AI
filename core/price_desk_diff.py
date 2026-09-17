"""Diff current price table vs parsed file rows (no DELETE of missing keys)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping, Sequence

PRICE_EPS = 0.005
EXAMPLES_LIMIT = 30

PriceKey = tuple


@dataclass(frozen=True)
class DiffExample:
    key: str
    old_price: float | None
    new_price: float | None


@dataclass
class PriceDeskDiffExamples:
    changed: list[DiffExample] = field(default_factory=list)
    new: list[DiffExample] = field(default_factory=list)
    missing: list[DiffExample] = field(default_factory=list)
    unchanged: list[DiffExample] = field(default_factory=list)


@dataclass
class PriceDeskDiff:
    changed: int
    new: int
    missing: int
    unchanged: int
    examples: PriceDeskDiffExamples


def format_price_key(key: PriceKey) -> str:
    if len(key) == 1:
        return str(key[0])
    if len(key) == 2 and isinstance(key[0], int):
        return f"{key[0]}-{key[1]}"
    return " / ".join(str(part) for part in key)


def prices_to_map(rows: Sequence[tuple], *, key_len: int) -> dict[PriceKey, float]:
    mapping: dict[PriceKey, float] = {}
    for row in rows:
        key = tuple(row[:key_len])
        price = float(row[key_len])
        mapping[key] = price
    return mapping


def diff_price_maps(
    current: Mapping[PriceKey, float],
    incoming: Mapping[PriceKey, float],
) -> PriceDeskDiff:
    changed: list[DiffExample] = []
    new: list[DiffExample] = []
    missing: list[DiffExample] = []
    unchanged: list[DiffExample] = []

    for key, new_price in incoming.items():
        old_price = current.get(key)
        example = DiffExample(
            key=format_price_key(key),
            old_price=old_price,
            new_price=new_price,
        )
        if old_price is None:
            new.append(example)
        elif abs(old_price - new_price) > PRICE_EPS:
            changed.append(example)
        else:
            unchanged.append(example)

    for key, old_price in current.items():
        if key not in incoming:
            missing.append(
                DiffExample(
                    key=format_price_key(key),
                    old_price=old_price,
                    new_price=None,
                )
            )

    examples = PriceDeskDiffExamples(
        changed=changed[:EXAMPLES_LIMIT],
        new=new[:EXAMPLES_LIMIT],
        missing=missing[:EXAMPLES_LIMIT],
        unchanged=unchanged[:EXAMPLES_LIMIT],
    )
    return PriceDeskDiff(
        changed=len(changed),
        new=len(new),
        missing=len(missing),
        unchanged=len(unchanged),
        examples=examples,
    )
