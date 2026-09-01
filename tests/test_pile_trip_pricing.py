"""PT-101: гибридный расчёт рейсов свай (полные + остатки 19800 + ручные N)."""

from __future__ import annotations

import math
from pathlib import Path

import pytest

from core.pile_catalog import (
    PileCatalogEntry,
    parse_pile_catalog_from_xlsx,
    resolve_catalog_for_mark,
)
from core.pile_trip_pricing import (
    PILE_REMAINDER_TRUCK_CAPACITY_KG,
    PileTripBreakdown,
    compute_pile_trips,
)

REAL_XLSX = Path(__file__).resolve().parents[1] / "банк знаний" / "сваи вес и объем.xlsx"

# Тендер без C18: полные 17+2+3+3+3+11 = 39; остаток 46950 кг → 3; итого 42.
TENDER_LINES_NO_C18 = [
    {"mark": "C14-40T4", "qty": 52, "product_type": "bridge_piles"},
    {"mark": "C9-35T6", "qty": 19, "product_type": "bridge_piles"},
    {"mark": "C10-35T6", "qty": 19, "product_type": "bridge_piles"},
    {"mark": "C13-35T7", "qty": 19, "product_type": "bridge_piles"},
    {"mark": "C11-35T6", "qty": 19, "product_type": "bridge_piles"},
    {"mark": "C15-35T6", "qty": 45, "product_type": "bridge_piles"},
]
C18_LINE = {"mark": "C18-40T8", "qty": 49, "product_type": "bridge_piles"}

# Канон Excel: С140.40 / С90.35 / … (если реального файла нет — те же цифры).
_FALLBACK_CATALOG = [
    PileCatalogEntry("С140.40", 14.0, 400, 2.26, 5650.0, 3),
    PileCatalogEntry("С90.35", 9.0, 350, 1.12, 2800.0, 7),
    PileCatalogEntry("С100.35", 10.0, 350, 1.24, 3100.0, 6),
    PileCatalogEntry("С130.35", 13.0, 350, 1.61, 4030.0, 5),
    PileCatalogEntry("С110.35", 11.0, 350, 1.37, 3430.0, 6),
    PileCatalogEntry("С150.35", 15.0, 350, 1.86, 4650.0, 4),
    PileCatalogEntry("С160.40", 16.0, 400, 2.58, 6450.0, None),
]


def _catalog_entries() -> list[PileCatalogEntry]:
    if REAL_XLSX.is_file():
        return parse_pile_catalog_from_xlsx(str(REAL_XLSX), sheet="Лист1")
    return list(_FALLBACK_CATALOG)


def _lookup(entries: list[PileCatalogEntry]):
    def lookup(mark: str) -> PileCatalogEntry | None:
        return resolve_catalog_for_mark(mark, entries)

    return lookup


def test_remainder_capacity_is_19800_not_18600() -> None:
    assert PILE_REMAINDER_TRUCK_CAPACITY_KG == 19800


def test_tender_without_c18_is_42_trips() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(TENDER_LINES_NO_C18, overrides=None, catalog_lookup=lookup)
    assert result.full_trips == 39
    assert result.remainder_kg == pytest.approx(46950.0)
    assert result.remainder_trips == 3
    assert result.override_trips == 0
    assert result.pending_marks == ()
    assert result.ready is True
    assert result.total_trips == 42
    assert math.ceil(46950.0 / PILE_REMAINDER_TRUCK_CAPACITY_KG) == 3


def test_c18_without_override_is_pending_and_not_counted() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [*TENDER_LINES_NO_C18, C18_LINE],
        overrides=None,
        catalog_lookup=lookup,
    )
    assert result.ready is False
    assert "C18-40T8" in result.pending_marks or any(
        "18-40" in m for m in result.pending_marks
    )
    assert result.total_trips == 0


def test_c18_override_n_adds_to_42() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [*TENDER_LINES_NO_C18, C18_LINE],
        overrides={"C18-40T8": 5},
        catalog_lookup=lookup,
    )
    assert result.ready is True
    assert result.pending_marks == ()
    assert result.override_trips == 5
    assert result.total_trips == 47  # 42 + 5


def test_c18_explicit_zero_is_ready_without_extra_trips() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [*TENDER_LINES_NO_C18, C18_LINE],
        overrides={"С18-40Т8": 0},
        catalog_lookup=lookup,
    )
    assert result.ready is True
    assert result.override_trips == 0
    assert result.total_trips == 42


def test_empty_override_is_not_zero() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [*TENDER_LINES_NO_C18, C18_LINE],
        overrides={},
        catalog_lookup=lookup,
    )
    assert result.ready is False
    assert result.total_trips == 0


def test_pcs_null_is_pending_unless_override() -> None:
    lookup = _lookup(_catalog_entries())
    line = {"mark": "С160.40", "qty": 4, "product_type": "piles"}
    pending = compute_pile_trips([line], overrides=None, catalog_lookup=lookup)
    assert pending.ready is False
    assert pending.total_trips == 0
    ready = compute_pile_trips([line], overrides={"С160.40": 2}, catalog_lookup=lookup)
    assert ready.ready is True
    assert ready.total_trips == 2
    assert ready.full_trips == 0
    assert ready.remainder_trips == 0


def test_unknown_mark_is_pending() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [{"mark": "XYZ-99", "qty": 3, "product_type": "piles"}],
        overrides=None,
        catalog_lookup=lookup,
    )
    assert result.ready is False
    assert result.total_trips == 0


def test_remainder_modulo_and_zero_remainder() -> None:
    entries = [
        PileCatalogEntry("С60.30", 6.0, 300, 0.55, 1380.0, 14),
    ]
    lookup = _lookup(entries)
    exact = compute_pile_trips(
        [{"mark": "С60.30", "qty": 28, "product_type": "piles"}],
        overrides=None,
        catalog_lookup=lookup,
    )
    assert exact.full_trips == 2
    assert exact.remainder_kg == pytest.approx(0.0)
    assert exact.remainder_trips == 0
    assert exact.total_trips == 2

    leftover = compute_pile_trips(
        [{"mark": "С60.30", "qty": 15, "product_type": "piles"}],
        overrides=None,
        catalog_lookup=lookup,
    )
    assert leftover.full_trips == 1
    assert leftover.remainder_kg == pytest.approx(1380.0)
    assert leftover.remainder_trips == 1
    assert leftover.total_trips == 2


def test_plates_and_fbs_are_ignored() -> None:
    lookup = _lookup(_catalog_entries())
    result = compute_pile_trips(
        [
            {"mark": "ПБ 60-12", "qty": 10, "product_type": "plates"},
            {"mark": "ФБС 24-4-6", "qty": 8, "product_type": "fbs"},
            *TENDER_LINES_NO_C18,
        ],
        overrides=None,
        catalog_lookup=lookup,
    )
    assert result.total_trips == 42
    assert result.ready is True


def test_breakdown_ready_property() -> None:
    pending = PileTripBreakdown(
        full_trips=39,
        remainder_kg=46950.0,
        remainder_trips=3,
        override_trips=0,
        pending_marks=("C18-40T8",),
        total_trips=0,
    )
    assert pending.ready is False
    ready = PileTripBreakdown(
        full_trips=39,
        remainder_kg=46950.0,
        remainder_trips=3,
        override_trips=0,
        pending_marks=(),
        total_trips=42,
    )
    assert ready.ready is True
