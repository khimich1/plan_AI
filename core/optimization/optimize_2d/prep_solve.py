# -*- coding: utf-8 -*-
"""Phase A: demand prep, cut-option generation, ILP build, solver run."""

from __future__ import annotations

from core.config_and_data import canonical_plate_key
from core.optimization.geometry import (
    GeometryConfig,
    NARROWING_TABLE,
    filter_secondary_cut_options_2d,
    generate_primary_cut_options_2d,
    generate_raw_secondary_cut_options_2d,
)
from core.optimization.ilp_model import build_two_d_cutting_ilp
from core.optimization.logging_utils import order_line_for_console
from core.optimization.optimization_config import DEFAULT_CONFIG, OptimizationConfig
from core.optimization.optimize_2d.state import TwoDPhaseAState
from core.optimization.ports.order_data import PlateOrderDataPort, resolve_plate_order_port
from core.optimization.order_dispatch import (
    _build_proportional_slot_lists,
    _peek_order_info,
    build_order_info_list,
)
from core.optimization.result_contract import (
    ERROR_EMPTY_ORDERS_2D,
    ERROR_PULP_MISSING,
    ERROR_SOLVER_INFEASIBLE,
    ERROR_SOLVER_UNDEFINED,
    opt_error,
)

def run_two_d_phase_a(
    *,
    orders_2d: list,
    plate_width: int = 1200,
    min_useful_width: int = 200,
    opt_config: OptimizationConfig | None = None,
    order_data: PlateOrderDataPort | None = None,
) -> tuple[TwoDPhaseAState | None, dict | None]:
    """
    Prepare demand and cut options, build the 2D ILP, and run the solver.

    Returns:
        (state, None) on success (including non-optimal but feasible statuses).
        (None, error_dict) on missing PuLP, empty orders, or infeasible/undefined solve.
    """
    if opt_config is None:
        opt_config = DEFAULT_CONFIG
    _order_view = resolve_plate_order_port(order_data)
    try:
        from pulp import PULP_CBC_CMD, LpStatus, value
    except ImportError:
        print("[OPT_2D] PuLP не установлен.")
        return None, opt_error(
            ERROR_PULP_MISSING,
            "PuLP не установлен — 2D ILP недоступен.",
        )

    if not orders_2d:
        return None, opt_error(ERROR_EMPTY_ORDERS_2D, "Пустой список заказов orders_2d.")


    print(f"\n[OPT_2D] === ПОЛНАЯ 2D ОПТИМИЗАЦИЯ ===")
    print(f"[OPT_2D] Заказ:")
    for order in orders_2d:
        print(order_line_for_console(order))

    # 1. ПОДГОТОВКА: Группируем спрос по (length, width, load_code)
    demand_2d = {}
    for order in orders_2d:
        key = canonical_plate_key(order["length"], order["width"], order.get("load_code", 800))
        demand_2d[key] = demand_2d.get(key, 0) + order["qty"]
    order_info_list = build_order_info_list(orders_2d, _order_view)

    slot_lists, slot_cursors = _build_proportional_slot_lists(orders_2d, demand_2d)

    tolerance_width = 20
    demand_tolerance_width = 10
    tolerance_length = 0

    geometry_config = GeometryConfig(
        plate_width=plate_width,
        min_useful_width=min_useful_width,
        tolerance_width=tolerance_width,
    )
    primary_result = generate_primary_cut_options_2d(
        demand_2d=demand_2d,
        order_info_list=order_info_list,
        order_info_getter=_peek_order_info,
        config=geometry_config,
    )
    primary_options = primary_result.raw_options
    solid_widths = primary_result.solid_widths

    print(f"[OPT_2D] Таблица narrowing создана: {len(NARROWING_TABLE)} правил")
    for target_w, sources in primary_result.target_to_sources.items():
        print(f"  {target_w}мм можно получить через: {sources}")

    print(f"[OPT_2D] Опций первичных резов (до фильтрации): {len(primary_options)}")
    primary_options = primary_result.options
    print(f"[OPT_2D] После фильтрации осталось: {len(primary_options)} первичных опций")
    secondary_options = generate_raw_secondary_cut_options_2d(
        primary_options=primary_options,
        demand_2d=demand_2d,
        config=geometry_config,
    )
    print(f"[OPT_2D] Опций вторичных резов (до фильтрации): {len(secondary_options)}")

    secondary_options = filter_secondary_cut_options_2d(secondary_options)
    print(f"[OPT_2D] После фильтрации осталось: {len(secondary_options)} вторичных опций")

    ilp = build_two_d_cutting_ilp(
        demand_2d=demand_2d,
        primary_options=primary_options,
        secondary_options=secondary_options,
        solid_widths=solid_widths,
        plate_width=plate_width,
        demand_tolerance_width=demand_tolerance_width,
        opt_config=opt_config,
    )
    prob = ilp.prob
    x_prim = ilp.x_prim
    x_sec = ilp.x_sec
    z_prim = ilp.z_prim
    z_sec = ilp.z_sec
    slack_solid = ilp.slack_solid
    unmet = ilp.unmet
    dk_list = ilp.dk_list
    primary_pairs_per_dk = ilp.primary_pairs_per_dk
    secondary_pairs_per_dk = ilp.secondary_pairs_per_dk
    solid_pairs_per_dk = ilp.solid_pairs_per_dk
    no_sources_keys = ilp.no_sources_keys


    print(
        f"[OPT_2D] Запуск решателя: {len(primary_options)} primary, "
        f"{len(secondary_options)} secondary, {len(z_prim) + len(z_sec)} z-vars, "
        f"{len(dk_list)} demand keys"
    )
    prob.solve(PULP_CBC_CMD(msg=0, timeLimit=60, gapRel=0.005))

    _solver_status = LpStatus[prob.status]
    print(f"[OPT_2D] Статус решателя: {_solver_status}")
    if _solver_status not in ("Optimal",):
        import logging as _solver_status_log

        _solver_status_log.getLogger(__name__).warning(
            "[OPT_2D] Решатель завершился со статусом %s — "
            "будет извлечён частичный результат, остатки добёрет post-correction",
            _solver_status,
        )
        if _solver_status in ("Infeasible", "Undefined"):
            print(f"[OPT_2D] ⚠️ Решение не найдено! Статус: {_solver_status}")
            _code = ERROR_SOLVER_UNDEFINED if _solver_status == "Undefined" else ERROR_SOLVER_INFEASIBLE
            return None, opt_error(
                _code,
                f"Решатель завершился со статусом {_solver_status}.",
                solver_status=_solver_status,
            )

    _slack_total = int(round(sum((value(s) or 0) for s in slack_solid.values())))
    _unmet_total = int(round(sum((value(u) or 0) for u in unmet.values())))
    if _slack_total > 0 or _unmet_total > 0:
        import logging as _slack_diag_log

        _slack_logger = _slack_diag_log.getLogger(__name__)
        _slack_logger.warning(
            "[OPT_2D] [SLACK] solid_slack=%d, unmet=%d (см. unmet_d* / slack_solid_d*)",
            _slack_total,
            _unmet_total,
        )
        for dk, sv in unmet.items():
            v = value(sv) or 0
            if v > 0.5:
                _slack_logger.warning("[OPT_2D] [UNMET] %s = %d", dk, int(round(v)))
        for dk, sv in slack_solid.items():
            v = value(sv) or 0
            if v > 0.5:
                _slack_logger.warning("[OPT_2D] [SOLID_SLACK] %s = %d", dk, int(round(v)))

    state = TwoDPhaseAState(
        orders_2d=orders_2d,
        demand_2d=demand_2d,
        order_info_list=order_info_list,
        slot_lists=slot_lists,
        slot_cursors=slot_cursors,
        geometry_config=geometry_config,
        primary_options=primary_options,
        secondary_options=secondary_options,
        solid_widths=solid_widths,
        ilp=ilp,
        solver_status=_solver_status,
    )
    return state, None
