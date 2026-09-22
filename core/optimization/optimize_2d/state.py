# -*- coding: utf-8 -*-
"""Cross-phase state for the 2D optimizer (split across phases A/B/C)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from core.config_and_data import canonical_plate_key
from core.optimization.geometry import GeometryConfig

if TYPE_CHECKING:
    from core.optimization.ilp_model import TwoDCuttingILPArtifacts


def norm_demand_key(k: tuple | list | None) -> tuple:
    """Normalize demand key identically to ``verify_coverage`` (canonical_plate_key)."""
    if not k or len(k) < 2:
        return (0, 0, 800)
    lc = k[2] if len(k) > 2 else 800
    return canonical_plate_key(k[0], k[1], lc)


@dataclass
class TwoDPhaseAState:
    """
    State after prep, ILP build, and solve (phase A).
    Phases B/C consume this for cut extraction and post-correction.
    """

    orders_2d: list
    demand_2d: dict
    order_info_list: list
    slot_lists: dict
    slot_cursors: dict
    geometry_config: GeometryConfig
    primary_options: list[dict]
    secondary_options: list[dict]
    solid_widths: Any
    ilp: TwoDCuttingILPArtifacts
    solver_status: str
