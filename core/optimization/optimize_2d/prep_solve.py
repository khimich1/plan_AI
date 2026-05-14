# -*- coding: utf-8 -*-
"""Phase A: demand prep, cut-option generation, ILP build, solver run."""

from __future__ import annotations

import json
import time
from core.config_and_data import canonical_plate_key
from core.optimization.debug_log import (
    _DEBUG_LOG_5b5324,
    _DEBUG_LOG_COMMON,
    _dbg_open_append,
    _opt_debug_enabled,
)
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
from core.optimization.optimization_debug_impl import (
    _DEBUG_AGENT_LOG_EBB546,
    _DEBUG_LOG_2D5C43,
    _DEBUG_LOG_7E420E,
    _debug_runtime_write_648532,
)
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

    # #region agent log
    if _opt_debug_enabled():
        try:
            with _dbg_open_append(_DEBUG_LOG_7E420E) as _lf:
                _lf.write(
                    json.dumps(
                        {
                            "sessionId": "7e420e",
                            "hypothesisId": "H_OPT_ENTER",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "2D ILP optimizer entered (fresh plan build)",
                            "data": {"n_orders": len(orders_2d), "plate_width": int(plate_width)},
                            "timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

    print(f"\n[OPT_2D] === ПОЛНАЯ 2D ОПТИМИЗАЦИЯ ===")
    print(f"[OPT_2D] Заказ:")
    for order in orders_2d:
        print(order_line_for_console(order))

    # 1. ПОДГОТОВКА: Группируем спрос по (length, width, load_code)
    demand_2d = {}
    for order in orders_2d:
        key = canonical_plate_key(order["length"], order["width"], order.get("load_code", 800))
        demand_2d[key] = demand_2d.get(key, 0) + order["qty"]
    # #region agent log: demand keys 59/10 (H3,H4)
    if _opt_debug_enabled():
        try:
            _demand_59_10 = [
                (list(k), v)
                for k, v in demand_2d.items()
                if abs(k[0] - 5.99) < 0.02 and (k[2] == 10 or abs(float(k[2]) - 10) < 0.01)
            ]
            if _demand_59_10:
                with _dbg_open_append(_DEBUG_LOG_COMMON) as _df:
                    _df.write(
                        json.dumps(
                            {
                                "hypothesisId": "H_59_10_demand",
                                "location": "core/optimization/optimize_2d/prep_solve.py",
                                "message": "demand_2d: ключи 5.99м 10п (length, width, load_code)",
                                "data": {"keys": _demand_59_10},
                                "timestamp": time.time(),
                            },
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
        except Exception:
            pass
    # #endregion
    order_info_list = build_order_info_list(orders_2d, _order_view)

    slot_lists, slot_cursors = _build_proportional_slot_lists(orders_2d, demand_2d)
    # #region agent log (session 5b5324) после построения slot_lists
    if _opt_debug_enabled():
        try:
            _slot_summary = [(list(k), len(slots)) for k, slots in list(slot_lists.items())[:20]]
            _sample_slot = []
            for k, slots in list(slot_lists.items())[:3]:
                if slots:
                    _sample_slot.append(
                        {
                            "key": list(k),
                            "first_identity": [
                                slots[0].get("kp_id"),
                                (slots[0].get("plate_name") or "")[:50],
                            ],
                        }
                    )
            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                _f.write(
                    json.dumps(
                        {
                            "sessionId": "5b5324",
                            "hypothesisId": "H_slots",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "slot_lists summary",
                            "data": {
                                "slot_summary": _slot_summary,
                                "sample_slot_identity": _sample_slot,
                            },
                            "timestamp": time.time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
    # #region agent log
    if _opt_debug_enabled():
        try:
            _slot_mismatch = []
            _slot_empty = []
            for _k, _need in demand_2d.items():
                _slots_len = len(slot_lists.get(_k, []))
                if _slots_len == 0 and _need > 0:
                    _slot_empty.append({"key": list(_k), "need": int(_need)})
                if _slots_len < _need:
                    _slot_mismatch.append({"key": list(_k), "need": int(_need), "slots": int(_slots_len)})
            _debug_runtime_write_648532(
                "run1",
                "H1_slot_key_alignment",
                "core/optimization/optimize_2d/prep_solve.py",
                "Demand keys versus proportional slot capacity",
                {
                    "demand_total": int(sum(demand_2d.values())),
                    "slot_total": int(sum(len(v) for v in slot_lists.values())),
                    "demand_keys": int(len(demand_2d)),
                    "slot_keys": int(len(slot_lists)),
                    "empty_slot_keys": _slot_empty[:50],
                    "short_slot_keys": _slot_mismatch[:50],
                },
            )
        except Exception:
            pass
    # #endregion

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
    # #region agent log (2d5c43) Plan B: опции для 5.1/320 и 6/530 до фильтра
    if _opt_debug_enabled():
        try:
            _log = _DEBUG_LOG_2D5C43
            _opts_320 = [
                {"id": o["id"], "length": o["length"], "main": o["main"], "type": o.get("type"), "load_code": o.get("load_code")}
                for o in primary_options
                if o.get("main") == 320 or o.get("target_width") == 320
            ]
            _opts_530 = [
                {"id": o["id"], "length": o["length"], "main": o["main"], "type": o.get("type"), "load_code": o.get("load_code")}
                for o in primary_options
                if o.get("main") == 530 or o.get("target_width") == 530
            ]
            with _dbg_open_append(_log) as _f:
                _f.write(
                    json.dumps(
                        {
                            "sessionId": "2d5c43",
                            "hypothesisId": "H_opt_gen",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "options for 320 and 530 before filter",
                            "data": {"opts_320": _opts_320, "opts_530": _opts_530, "solid_widths": solid_widths},
                            "timestamp": time.time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
    primary_options = primary_result.options
    print(f"[OPT_2D] После фильтрации осталось: {len(primary_options)} первичных опций")
    # #region agent log (2d5c43) Plan B: опции для 320/530 после фильтра
    if _opt_debug_enabled():
        try:
            _log = _DEBUG_LOG_2D5C43
            _opts_320 = [
                {"id": o["id"], "length": o["length"], "main": o["main"], "type": o.get("type")}
                for o in primary_options
                if o.get("main") == 320 or o.get("target_width") == 320
            ]
            _opts_530 = [
                {"id": o["id"], "length": o["length"], "main": o["main"], "type": o.get("type")}
                for o in primary_options
                if o.get("main") == 530 or o.get("target_width") == 530
            ]
            with _dbg_open_append(_log) as _f:
                _f.write(
                    json.dumps(
                        {
                            "sessionId": "2d5c43",
                            "hypothesisId": "H_opt_filter",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "options for 320 and 530 after filter",
                            "data": {"opts_320": _opts_320, "opts_530": _opts_530},
                            "timestamp": time.time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
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

    # #region agent log
    if _opt_debug_enabled():
        try:
            _supply_diag = []
            for _dk in dk_list:
                _supply_diag.append(
                    {
                        "dk": list(_dk),
                        "need_qty": int(demand_2d.get(_dk, 0)),
                        "primary_opts": len(primary_pairs_per_dk.get(_dk) or []),
                        "secondary_opts": len(secondary_pairs_per_dk.get(_dk) or []),
                        "solid_opts": len(solid_pairs_per_dk.get(_dk) or []),
                    }
                )
            with _dbg_open_append(_DEBUG_AGENT_LOG_EBB546) as _agent_f:
                _agent_f.write(
                    json.dumps(
                        {
                            "sessionId": "ebb546",
                            "runId": "solver-localization",
                            "hypothesisId": "O1,O2",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "Demand keys и количество доступных опций до solve",
                            "data": {
                                "demand_total": int(sum(demand_2d.values())),
                                "demand_keys": int(len(dk_list)),
                                "no_sources_keys": [{"dk": list(k), "qty": int(q)} for k, q in (no_sources_keys or [])],
                                "supply_diag": _supply_diag[:200],
                            },
                            "timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

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
    # #region agent log
    if _opt_debug_enabled():
        try:
            _coverage_diag = []
            _missing_after_solver = []
            for _dk in dk_list:
                _need = int(demand_2d.get(_dk, 0))
                _z_prim = int(
                    round(
                        sum(
                            (value(z_prim.get((_oid, _dk))) or 0)
                            for _oid in (primary_pairs_per_dk.get(_dk) or [])
                        )
                    )
                )
                _z_sec = int(
                    round(
                        sum(
                            (value(z_sec.get((_oid, _dk))) or 0)
                            for _oid in (secondary_pairs_per_dk.get(_dk) or [])
                        )
                    )
                )
                _unmet_v = int(round(value(unmet.get(_dk)) or 0))
                _covered = _z_prim + _z_sec
                _coverage_diag.append(
                    {
                        "dk": list(_dk),
                        "need_qty": _need,
                        "z_prim": _z_prim,
                        "z_sec": _z_sec,
                        "covered": _covered,
                        "unmet": _unmet_v,
                        "supply_primary_opts": len(primary_pairs_per_dk.get(_dk) or []),
                        "supply_secondary_opts": len(secondary_pairs_per_dk.get(_dk) or []),
                    }
                )
                if _covered < _need:
                    _missing_after_solver.append(
                        {
                            "dk": list(_dk),
                            "need_qty": _need,
                            "covered": _covered,
                            "missing": _need - _covered,
                            "unmet": _unmet_v,
                        }
                    )
            with _dbg_open_append(_DEBUG_AGENT_LOG_EBB546) as _agent_f:
                _agent_f.write(
                    json.dumps(
                        {
                            "sessionId": "ebb546",
                            "runId": "solver-localization",
                            "hypothesisId": "O3,O4,O5",
                            "location": "core/optimization/optimize_2d/prep_solve.py",
                            "message": "Покрытие спроса по dk после solve через z_prim/z_sec/unmet",
                            "data": {
                                "solver_status": _solver_status,
                                "demand_total": int(sum(demand_2d.values())),
                                "z_prim_total": int(round(sum((value(v) or 0) for v in z_prim.values()))),
                                "z_sec_total": int(round(sum((value(v) or 0) for v in z_sec.values()))),
                                "unmet_total": _unmet_total,
                                "slack_total": _slack_total,
                                "missing_after_solver_total": int(sum(x["missing"] for x in _missing_after_solver)),
                                "missing_after_solver": _missing_after_solver[:120],
                                "coverage_diag": _coverage_diag[:250],
                            },
                            "timestamp": int(time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

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
