#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль оптимизации раскроя плит (тонкий фасад реэкспорта).
Терминология: продольный рез — по ширине; поперечный — по длине.
"""

from core.optimization.context import (
    LOAD_TO_REINFORCEMENT_MAP,
    OPT_CASCADING_PLAN,
    OPT_CASCADING_PLAN_BY_LOAD,
    OPT_PLAN,
    OPT_WIDTH_PRIORITY,
)
from core.optimization.coverage_verify import verify_coverage
from core.optimization.ffd_packing import (
    Piece,
    Track,
    first_fit_decreasing,
    optimize_tracks,
    pack_tracks,
)
from core.optimization.geometry import (
    GeometryConfig,
    KERF_WIDTH_MM,
    NARROWING_TABLE,
    filter_secondary_cut_options_2d,
    generate_primary_cut_options_1d,
    generate_primary_cut_options_2d,
    generate_raw_secondary_cut_options_2d,
    generate_secondary_cut_options_1d,
)
from core.optimization.ilp_model import (
    _build_residual_balance_constraints,
    _residual_phys_key,
    build_two_d_cutting_ilp,
)
from core.optimization.legacy_width_plan import (
    _append_actions,
    _group_plate_lengths,
    apply_width_optimization,
    optimize_cuts_pulp,
)
from core.optimization.optimize_1d_widths import _optimize_1d_widths_only
from core.optimization.optimize_2d.with_lengths import _optimize_2d_with_lengths
from core.optimization.optimization_config import (
    DEFAULT_CONFIG,
    OLD_CONFIG,
    OptimizationConfig,
)
from core.optimization.order_dispatch import (
    _build_proportional_slot_lists,
    _get_next_order_info,
    _next_slot_info,
    _peek_order_info,
    build_order_info_list,
)
from core.optimization.pulp_qty import _opt_1d_pulp_nonneg_qty

__all__ = (
    "DEFAULT_CONFIG",
    "GeometryConfig",
    "KERF_WIDTH_MM",
    "LOAD_TO_REINFORCEMENT_MAP",
    "NARROWING_TABLE",
    "OLD_CONFIG",
    "OPT_CASCADING_PLAN",
    "OPT_CASCADING_PLAN_BY_LOAD",
    "OPT_PLAN",
    "OPT_WIDTH_PRIORITY",
    "OptimizationConfig",
    "Piece",
    "Track",
    "apply_width_optimization",
    "build_order_info_list",
    "build_two_d_cutting_ilp",
    "filter_secondary_cut_options_2d",
    "first_fit_decreasing",
    "generate_primary_cut_options_1d",
    "generate_primary_cut_options_2d",
    "generate_raw_secondary_cut_options_2d",
    "generate_secondary_cut_options_1d",
    "optimize_cuts_pulp",
    "optimize_tracks",
    "pack_tracks",
    "verify_coverage",
    "_append_actions",
    "_build_proportional_slot_lists",
    "_build_residual_balance_constraints",
    "_get_next_order_info",
    "_group_plate_lengths",
    "_next_slot_info",
    "_optimize_1d_widths_only",
    "_optimize_2d_with_lengths",
    "_peek_order_info",
    "_residual_phys_key",
)
