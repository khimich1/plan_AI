# -*- coding: utf-8 -*-
"""Typed, immutable application defaults (track geometry, cut pricing, weight factor)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

_AppConfig_singleton: Optional["AppConfig"] = None


@dataclass(frozen=True, slots=True)
class AppConfig:
    """Tunable/static parameters mirrored from ``core.config.constants`` at first ``get_config()`` call."""

    track_length_m: float
    track_width_m: float
    long_cut_price_per_m: float
    transverse_cut_price: float
    weight_kg_per_dm2: float


def _build_app_config() -> AppConfig:
    from . import constants as c

    return AppConfig(
        track_length_m=c.TRACK_LENGTH_M,
        track_width_m=c.TRACK_WIDTH_M,
        long_cut_price_per_m=c.LONG_CUT_PRICE_PER_M,
        transverse_cut_price=c.TRANSVERSE_CUT_PRICE,
        weight_kg_per_dm2=c.WEIGHT_KG_PER_DM2,
    )


def get_config() -> AppConfig:
    """Lazily constructed module-level singleton; reads current ``constants`` values on first use."""
    global _AppConfig_singleton
    if _AppConfig_singleton is None:
        _AppConfig_singleton = _build_app_config()
    return _AppConfig_singleton
