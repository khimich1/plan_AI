# -*- coding: utf-8 -*-
"""Read-only port for plate/order normalization and legacy 1D cost heuristics."""

from __future__ import annotations

from typing import Any, Protocol

from core.config.constants import LONG_CUT_PRICE_PER_M
from core.domain.plate_order import normalize_load_code

# Legacy 1D total_cost used a flat plate price not tied to PRICE_DB in this path.
_LEGACY_1D_PLATE_UNIT_PRICE_RUB = 12_000
# Matches historical literal in optimize_1d_widths (LONG_CUT_PRICE_PER_M was 460 → ×6).
_LEGACY_1D_LONG_CUT_FACTOR_M = 6


class PlateOrderDataPort(Protocol):
    """
    Minimal read-only surface for order/plate rules on the optimizer entry path.

    Implemented in-process by ``DefaultPlateOrderAdapter``.
    """

    def normalize_load_code(self, value: Any, default: int = 8) -> int | float:
        """Same contract as ``core.domain.plate_order.normalize_load_code``."""
        ...

    def one_d_plate_unit_price_rub(self) -> int:
        """Flat price per raw plate in the legacy 1D ``total_cost`` estimate."""
        ...

    def one_d_long_cut_cost_rub(self) -> int:
        """Longitudinal cut cost used per plate / secondary group in legacy 1D totals."""
        ...


class DefaultPlateOrderAdapter:
    """Thin adapter: explicit domain/constants imports; 1D costs match previous defaults."""

    def normalize_load_code(self, value: Any, default: int = 8) -> int | float:
        return normalize_load_code(value, default)

    def one_d_plate_unit_price_rub(self) -> int:
        return _LEGACY_1D_PLATE_UNIT_PRICE_RUB

    def one_d_long_cut_cost_rub(self) -> int:
        return int(round(LONG_CUT_PRICE_PER_M * _LEGACY_1D_LONG_CUT_FACTOR_M))


# Backward-compatible alias for tests/callers that referenced the cfg-module wrapper.
ConfigModulePlateOrderAdapter = DefaultPlateOrderAdapter


def default_plate_order_port() -> PlateOrderDataPort:
    """Composition-root default: explicit domain/constants (no PEP 562 module alias)."""
    return DefaultPlateOrderAdapter()


def resolve_plate_order_port(order_data: PlateOrderDataPort | None) -> PlateOrderDataPort:
    if order_data is None:
        return default_plate_order_port()
    return order_data
