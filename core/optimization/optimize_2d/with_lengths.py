# -*- coding: utf-8 -*-
"""Full 2D optimization: orchestrates prep/solve, extract, finalize.

Поток OPT_* (TLS): как и :func:`core.optimization.optimize_1d_widths._optimize_1d_widths_only`,
этот код не пишет в thread-local план — возвращается словарь результата, а проставление
``OPT_CASCADING_PLAN`` / смежных ключей делается composition root через присваивание
атрибутам ``core.optimization`` (делегирует в ``core.optimization.context``). Дублировать
TLS здесь не нужно.
"""

from __future__ import annotations

from core.optimization.optimization_config import DEFAULT_CONFIG, OptimizationConfig
from core.optimization.optimize_2d.extract_cuts import extract_two_d_phase_b
from core.optimization.optimize_2d.finalize import run_two_d_phase_finalize
from core.optimization.optimize_2d.prep_solve import run_two_d_phase_a
from core.optimization.ports.order_data import PlateOrderDataPort
from core.plate_audit import PlateAudit


def _optimize_2d_with_lengths(
    orders_2d: list,
    plate_width: int = 1200,
    min_useful_width: int = 200,
    opt_config: OptimizationConfig | None = None,
    order_data: PlateOrderDataPort | None = None,
) -> dict:
    """Полная 2D оптимизация: Phase A (prep+solve), Phase B (extract), Phase C (finalize)."""
    if opt_config is None:
        opt_config = DEFAULT_CONFIG

    phase_state, phase_err = run_two_d_phase_a(
        orders_2d=orders_2d,
        plate_width=plate_width,
        min_useful_width=min_useful_width,
        opt_config=opt_config,
        order_data=order_data,
    )
    if phase_err is not None:
        return phase_err

    audit = PlateAudit(orders_2d)
    audit.checkpoint("demand_2d", phase_state.demand_2d)

    result = {
        "primary_cuts": [],
        "secondary_cuts": [],
        "total_plates": 0,
        "plate_assignments": [],
        "orders_requested": [order.copy() for order in orders_2d],
        "rests_created": [],
        "rests_used": [],
    }

    phase_b = extract_two_d_phase_b(phase_state)
    result["primary_cuts"] = phase_b.primary_cuts
    result["secondary_cuts"] = phase_b.secondary_cuts
    result["total_plates"] = phase_b.total_plates
    result["rests_used"] = phase_b.rests_used

    return run_two_d_phase_finalize(
        demand_2d=phase_state.demand_2d,
        plate_width=plate_width,
        slot_lists=phase_state.slot_lists,
        slot_cursors=phase_state.slot_cursors,
        no_sources_keys=phase_state.ilp.no_sources_keys,
        solver_status=phase_state.solver_status,
        audit=audit,
        result=result,
        n_solid_primary_plates=phase_b.n_solid_primary_plates,
        n_cut_primary_plates=phase_b.n_cut_primary_plates,
        next_primary_instance_id=phase_b.next_primary_instance_id,
    )
