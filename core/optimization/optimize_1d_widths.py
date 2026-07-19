#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""1D PuLP-оптимизация раскроя только по ширине (каскадные продольные резы)."""

from __future__ import annotations

from core.optimization.geometry import (
    generate_primary_cut_options_1d,
    generate_secondary_cut_options_1d,
)
from core.optimization.pulp_qty import _opt_1d_pulp_nonneg_qty
from core.optimization.ports.order_data import PlateOrderDataPort, resolve_plate_order_port
from core.optimization.result_contract import (
    ERROR_EMPTY_ORDERS_1D,
    ERROR_PULP_MISSING,
    ERROR_SOLVER_INFEASIBLE,
    ERROR_SOLVER_UNDEFINED,
    opt_error,
    opt_ok,
)


def _optimize_1d_widths_only(
    orders: dict,
    plate_width: int = 1200,
    min_useful_width: int = 200,
    order_data: PlateOrderDataPort | None = None,
) -> dict:
    """
    ПРИВАТНАЯ функция: Оптимизация только по ширинам (1D).
    Длины НЕ учитываются в оптимизации, присваиваются позже.

    Args:
        orders: {300: 4, 500: 3} — спрос по ширинам в мм (без учёта длин!)
        plate_width: ширина исходной плиты в мм (1200)
        min_useful_width: минимальная полезная ширина остатка

    Returns:
        {
            'primary_cuts': [{'width': 320, 'rest': 880, 'qty': 2}, ...],
            'secondary_cuts': [{'source': 880, 'cuts': [320, 560], 'qty': 1}, ...],
            'total_plates': 5,
            'total_cost': 75000,
            'waste_width': 120
        }
    """
    try:
        from pulp import (
            LpInteger,
            LpMinimize,
            LpProblem,
            LpStatus,
            LpVariable,
            PULP_CBC_CMD,
            lpSum,
            value,
        )
    except ImportError:
        print("[OPT_CASCADING] PuLP не установлен, пропускаем.")
        return opt_error(
            ERROR_PULP_MISSING,
            "PuLP не установлен — 1D оптимизация недоступна.",
        )

    if not orders:
        return opt_error(ERROR_EMPTY_ORDERS_1D, "Пустой словарь заказов по ширине (1D).")

    target_widths = sorted(orders.keys())

    tolerance = 20

    primary_cut_options = generate_primary_cut_options_1d(
        target_widths=target_widths,
        plate_width=plate_width,
        min_useful_width=min_useful_width,
    )

    secondary_cut_options = generate_secondary_cut_options_1d(
        primary_cut_options=primary_cut_options,
        target_widths=target_widths,
        tolerance=tolerance,
    )
    possible_rests = {opt["rest"] for opt in primary_cut_options if opt["rest"] > 0}

    print(f"[DEBUG] Найдено вариантов вторичных резов: {len(secondary_cut_options)}")
    for opt in secondary_cut_options[:3]:
        print(
            f"  {opt['source_rest']}мм -> {opt.get('pieces', 1)}x{opt['output1']}мм "
            f"(отход {opt['waste']}мм)"
        )

    prob = LpProblem("cascading_longitudinal_cuts", LpMinimize)

    x_prim = {}
    for i, opt in enumerate(primary_cut_options):
        x_prim[i] = LpVariable(f"prim_{i}_{opt['id']}", lowBound=0, cat=LpInteger)

    x_sec = {}
    for i, opt in enumerate(secondary_cut_options):
        x_sec[i] = LpVariable(f"sec_{i}_{opt['id']}", lowBound=0, cat=LpInteger)

    target_widths_keys = list(orders.keys())
    primary_pairs_per_w: dict = {w: [] for w in target_widths_keys}
    secondary_pairs_per_w: dict = {w: [] for w in target_widths_keys}

    for w in target_widths_keys:
        for i, opt in enumerate(primary_cut_options):
            if abs(opt["main"] - w) <= tolerance:
                primary_pairs_per_w[w].append(i)
        for i, opt in enumerate(secondary_cut_options):
            if abs(opt["output1"] - w) <= tolerance:
                secondary_pairs_per_w[w].append((i, opt.get("pieces", 1)))
            if opt["output2"] > 0 and abs(opt["output2"] - w) <= tolerance:
                secondary_pairs_per_w[w].append((i, 1))

    z_prim_w: dict = {}
    z_sec_w: dict = {}
    unmet_w: dict = {}
    for w in target_widths_keys:
        for i in primary_pairs_per_w[w]:
            z_prim_w[(i, w)] = LpVariable(f"z1d_prim_{i}_w{w}", lowBound=0, cat=LpInteger)
        for (i, _) in secondary_pairs_per_w[w]:
            z_sec_w[(i, w)] = LpVariable(f"z1d_sec_{i}_w{w}", lowBound=0, cat=LpInteger)
        unmet_w[w] = LpVariable(f"unmet_w{w}", lowBound=0, cat=LpInteger)

    for w, qty in orders.items():
        parts = [z_prim_w[(i, w)] for i in primary_pairs_per_w[w]]
        parts += [z_sec_w[(i, w)] for (i, _) in secondary_pairs_per_w[w]]
        parts.append(unmet_w[w])
        prob += lpSum(parts) == qty, f"demand_w{w}"

    prim_to_ws_1d: dict = {}
    for w, ids in primary_pairs_per_w.items():
        for i in ids:
            prim_to_ws_1d.setdefault(i, []).append(w)
    for i, ws in prim_to_ws_1d.items():
        prob += (
            lpSum(z_prim_w[(i, w)] for w in ws) <= x_prim[i],
            f"cap_prim_w_{i}",
        )

    sec_to_ws_1d: dict = {}
    for w, pairs in secondary_pairs_per_w.items():
        for (i, contrib) in pairs:
            sec_to_ws_1d.setdefault(i, []).append((w, contrib))
    for i, ws_contribs in sec_to_ws_1d.items():
        opt = secondary_cut_options[i]
        outputs_per_app = opt.get("pieces", 1)
        if opt.get("output2", 0) > 0:
            outputs_per_app += 1
        prob += (
            lpSum(z_sec_w[(i, w)] for (w, _) in ws_contribs) <= x_sec[i] * outputs_per_app,
            f"cap_sec_w_{i}",
        )

    for rest_w in possible_rests:
        produced = [
            x_prim[i] for i, opt in enumerate(primary_cut_options) if opt["rest"] == rest_w
        ]
        consumed = [
            x_sec[i]
            for i, opt in enumerate(secondary_cut_options)
            if opt["source_rest"] == rest_w
        ]
        if produced and consumed:
            prob += lpSum(consumed) <= lpSum(produced), f"balance_rest_{rest_w}"

    M_UNMET_W = 1e7

    total_plates = lpSum(x_prim.values())

    unused_rests_penalty = 0
    for rest_w in possible_rests:
        produced = [
            x_prim[i] for i, opt in enumerate(primary_cut_options) if opt["rest"] == rest_w
        ]
        consumed = [
            x_sec[i]
            for i, opt in enumerate(secondary_cut_options)
            if opt["source_rest"] == rest_w
        ]
        if produced and consumed:
            unused = lpSum(produced) - lpSum(consumed)
            unused_rests_penalty += unused * (rest_w / 1000.0) * 0.05

    waste_penalty = 0
    for i, opt in enumerate(secondary_cut_options):
        waste_penalty += x_sec[i] * opt.get("waste", 0) * 0.0001

    obj = total_plates + unused_rests_penalty + waste_penalty
    if unmet_w:
        obj = obj + M_UNMET_W * lpSum(unmet_w.values())
    prob += obj
    prob.solve(PULP_CBC_CMD(msg=0, timeLimit=60, gapRel=0.005))

    _solver_status_1d = LpStatus[prob.status]
    if _solver_status_1d not in ("Optimal",):
        import logging as _solver_status_log_1d

        _solver_status_log_1d.getLogger(__name__).warning(
            "[OPT_1D] Решатель завершился со статусом %s — извлекаем частичный результат",
            _solver_status_1d,
        )
        if _solver_status_1d in ("Infeasible", "Undefined"):
            _code = (
                ERROR_SOLVER_UNDEFINED
                if _solver_status_1d == "Undefined"
                else ERROR_SOLVER_INFEASIBLE
            )
            return opt_error(
                _code,
                f"1D решатель: {_solver_status_1d}.",
                solver_status=_solver_status_1d,
            )

    _unmet_total_1d = int(round(sum((value(u) or 0) for u in unmet_w.values())))
    if _unmet_total_1d > 0:
        import logging as _unmet_log_1d

        _unmet_log_1d.getLogger(__name__).warning(
            "[OPT_1D] [UNMET] Дефицит: %d плит — модель не нашла источников",
            _unmet_total_1d,
        )

    result = {
        "primary_cuts": [],
        "secondary_cuts": [],
        "total_plates": 0,
        "total_cost": 0,
        "waste_width": 0,
    }

    for i, opt in enumerate(primary_cut_options):
        qty = _opt_1d_pulp_nonneg_qty(value, x_prim[i], context=f"x_prim[{i}]")
        if qty > 0:
            result["primary_cuts"].append(
                {
                    "width": opt["main"],
                    "rest": opt["rest"],
                    "qty": qty,
                }
            )
            result["total_plates"] += qty

    for i, opt in enumerate(secondary_cut_options):
        qty = _opt_1d_pulp_nonneg_qty(value, x_sec[i], context=f"x_sec[{i}]")
        if qty > 0:
            cuts = [opt["output1"]]
            if opt["output2"] > 0:
                cuts.append(opt["output2"])
            result["secondary_cuts"].append(
                {
                    "source": opt["source_rest"],
                    "cuts": cuts,
                    "pieces": opt.get("pieces", 1),
                    "qty": qty,
                    "waste": opt.get("waste", 0),
                }
            )
            result["waste_width"] += opt.get("waste", 0) * qty

    print("[OPT_1D] 🔧 Применяем правила завода для порядка плит...")

    solid_plates = []
    cut_plates = []

    for cut in result["primary_cuts"]:
        if cut["rest"] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)

    if solid_plates:
        solid_plates.sort(key=lambda x: (-x["width"]))

    cut_plates.sort(key=lambda x: (-x["rest"], -x["width"]))

    result["primary_cuts"] = solid_plates + cut_plates

    print(f"[OPT_1D] ✓ Целых плит в начале: {len(solid_plates)}")
    print(f"[OPT_1D] ✓ Плит с резом (сгруппировано): {len(cut_plates)}")

    print(f"[DEBUG] Оптимизатор выбрал вторичных резов: {len(result['secondary_cuts'])}")

    _odata = resolve_plate_order_port(order_data)
    plate_price = _odata.one_d_plate_unit_price_rub()
    long_cut_cost = _odata.one_d_long_cut_cost_rub()

    result["total_cost"] = (
        result["total_plates"] * plate_price
        + result["total_plates"] * long_cut_cost
        + len(result["secondary_cuts"]) * long_cut_cost
    )

    return opt_ok(result, partial=(_solver_status_1d != "Optimal"))
