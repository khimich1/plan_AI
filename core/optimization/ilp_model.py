#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PuLP ILP construction for 2D plate cutting (assignment + residual balance)."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from core.config.constants import LONG_CUT_PRICE_PER_M, TRANSVERSE_CUT_PRICE
from core.domain.plate_order import normalize_load_code
from core.optimization.debug_log import _DEBUG_LOG_COMMON, _dbg_open_append
from core.project_paths import PRICE_DB_PATH
from core.optimization.geometry import _canonical_length
from core.price_db import get_price


def _residual_phys_key(length, rest_width) -> tuple[float, int]:
    """Physical residual band key shared by optimizer constraints and parent fallback."""
    return (_canonical_length(length), int(round(float(rest_width or 0))))


def _build_residual_balance_constraints(
    *,
    prob: Any,
    primary_options: list[dict],
    secondary_options: list[dict],
    x_prim: dict,
    x_sec: dict,
) -> dict:
    """
    Enforce residual supply/consumption with downgrade load-code policy.

    A primary residual with higher/equal load_code may serve secondary demand with
    lower/equal load_code. The reverse is forbidden by cumulative constraints.
    """
    from pulp import lpSum

    supply_by_phys: dict[tuple[float, int], dict[float | int, list[int]]] = defaultdict(
        lambda: defaultdict(list)
    )
    demand_by_phys: dict[tuple[float, int], dict[float | int, list[int]]] = defaultdict(
        lambda: defaultdict(list)
    )

    for opt in primary_options:
        rest_w = opt.get("rest", 0)
        if rest_w > 0 and opt.get("type") != "solid":
            phys = _residual_phys_key(opt.get("length"), rest_w)
            prim_lc = normalize_load_code(opt.get("load_code", 8), default=8)
            opt["load_code"] = prim_lc
            supply_by_phys[phys][prim_lc].append(opt["id"])

    for opt in secondary_options:
        target_key = opt.get("target_order_key", (0, 0, 8))
        sec_lc = target_key[2] if len(target_key) == 3 else 8
        sec_lc = normalize_load_code(sec_lc, default=8)
        phys = _residual_phys_key(opt.get("source_length"), opt.get("source_rest"))
        demand_by_phys[phys][sec_lc].append(opt["id"])

    constraint_count = 0
    blocked_no_supply = 0
    rests_for_objective: dict = {}

    for phys, demand_by_lc in demand_by_phys.items():
        supply_by_lc = supply_by_phys.get(phys, {})
        if not supply_by_lc:
            for opt_ids in demand_by_lc.values():
                for opt_id in opt_ids:
                    prob += x_sec[opt_id] == 0, f"residual_no_supply_sec_{opt_id}"
                    blocked_no_supply += 1
            continue

        produced_all = [opt_id for ids in supply_by_lc.values() for opt_id in ids]
        consumed_all = [opt_id for ids in demand_by_lc.values() for opt_id in ids]
        rests_for_objective[(phys[0], phys[1], "all")] = {
            "produced": produced_all,
            "consumed": consumed_all,
        }

        levels = sorted(set(supply_by_lc.keys()) | set(demand_by_lc.keys()), reverse=True)
        for level in levels:
            consumed = [
                opt_id
                for target_lc, opt_ids in demand_by_lc.items()
                if target_lc >= level
                for opt_id in opt_ids
            ]
            produced = [
                opt_id
                for prim_lc, opt_ids in supply_by_lc.items()
                if prim_lc >= level
                for opt_id in opt_ids
            ]
            if consumed:
                prob += (
                    lpSum(x_sec[i] for i in consumed) <= lpSum(x_prim[i] for i in produced),
                    f"residual_balance_L{phys[0]}_R{phys[1]}_LC{level}",
                )
                constraint_count += 1

    try:
        import json as _json
        import time as _time

        with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
            _f.write(
                _json.dumps(
                    {
                        "hypothesisId": "residual_balance_constraints_added",
                        "location": "core/optimization/ilp_model.py:_build_residual_balance_constraints",
                        "message": "Residual balance constraints with downgrade load-code policy",
                        "data": {
                            "constraints_added": constraint_count,
                            "blocked_secondary_without_supply": blocked_no_supply,
                            "physical_supply_keys": len(supply_by_phys),
                            "physical_demand_keys": len(demand_by_phys),
                        },
                        "timestamp": int(_time.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass

    return rests_for_objective


@dataclass
class TwoDCuttingILPArtifacts:
    """PuLP model and indexing structures for 2D cutting (pre-solve)."""

    prob: Any
    x_prim: dict
    x_sec: dict
    z_prim: dict
    z_sec: dict
    slack_solid: dict
    unmet: dict
    dk_list: list
    dk_to_idx: dict
    primary_pairs_per_dk: dict
    secondary_pairs_per_dk: dict
    solid_pairs_per_dk: dict
    primary_options_by_id: dict
    secondary_options_by_id: dict
    no_sources_keys: list
    rests_by_lkey: dict


def build_two_d_cutting_ilp(
    *,
    demand_2d: dict,
    primary_options: list[dict],
    secondary_options: list[dict],
    solid_widths: Any,
    plate_width: int,
    demand_tolerance_width: int,
    opt_config: Any,
) -> TwoDCuttingILPArtifacts:
    """
    Build assignment-model ILP: variables, demand/cap constraints, residual balance, objective.
    Does not call the solver (caller runs solve and reads variable values).
    """
    from pulp import LpInteger, LpMinimize, LpProblem, LpVariable, lpSum

    prob = LpProblem("2D_Optimization", LpMinimize)

    x_prim = {
        opt["id"]: LpVariable(f"prim_{opt['id']}", lowBound=0, cat=LpInteger) for opt in primary_options
    }
    x_sec = {
        opt["id"]: LpVariable(f"sec_{opt['id']}", lowBound=0, cat=LpInteger) for opt in secondary_options
    }

    primary_options_by_id = {o["id"]: o for o in primary_options}
    secondary_options_by_id = {o["id"]: o for o in secondary_options}

    dk_list = list(demand_2d.keys())
    dk_to_idx = {dk: i for i, dk in enumerate(dk_list)}
    primary_pairs_per_dk: dict = {dk: [] for dk in dk_list}
    secondary_pairs_per_dk: dict = {dk: [] for dk in dk_list}
    solid_pairs_per_dk: dict = {dk: [] for dk in dk_list}
    no_sources_keys: list = []

    for dk in dk_list:
        target_length, target_width, target_load_code = dk
        for opt in primary_options:
            if (
                _canonical_length(opt["length"]) != _canonical_length(target_length)
                or opt.get("load_code", 800) != target_load_code
            ):
                continue
            opt_type = opt.get("type")
            if opt_type in ("direct", "solid"):
                if abs(opt["main"] - target_width) <= demand_tolerance_width:
                    primary_pairs_per_dk[dk].append(opt["id"])
                    if opt_type == "solid" and target_width in solid_widths:
                        solid_pairs_per_dk[dk].append(opt["id"])
            elif opt_type == "indirect":
                if abs(opt.get("target_width", 0) - target_width) <= demand_tolerance_width:
                    primary_pairs_per_dk[dk].append(opt["id"])

        for opt in secondary_options:
            opt_target_key = opt.get("target_order_key", (0, 0, 800))
            opt_target_load = opt_target_key[2] if len(opt_target_key) == 3 else 800
            if (
                _canonical_length(opt["output_length"]) == _canonical_length(target_length)
                and abs(opt["output_width"] - target_width) <= demand_tolerance_width
                and opt_target_load == target_load_code
            ):
                secondary_pairs_per_dk[dk].append(opt["id"])

        if not primary_pairs_per_dk[dk] and not secondary_pairs_per_dk[dk]:
            no_sources_keys.append((dk, demand_2d[dk]))
            import logging as _no_src_log

            _no_src_log.getLogger(__name__).error(
                "[OPT_2D] ❌ НЕТ ИСТОЧНИКОВ для плиты: %sм x %sмм (load=%s) x%dшт — закроется через unmet/post-correction",
                target_length,
                target_width,
                target_load_code,
                demand_2d[dk],
            )

    z_prim: dict = {}
    z_sec: dict = {}
    slack_solid: dict = {}
    unmet: dict = {}

    for dk in dk_list:
        di = dk_to_idx[dk]
        for opt_id in primary_pairs_per_dk[dk]:
            z_prim[(opt_id, dk)] = LpVariable(
                f"z_prim_{opt_id}_d{di}",
                lowBound=0,
                cat=LpInteger,
            )
        for opt_id in secondary_pairs_per_dk[dk]:
            z_sec[(opt_id, dk)] = LpVariable(
                f"z_sec_{opt_id}_d{di}",
                lowBound=0,
                cat=LpInteger,
            )
        if solid_pairs_per_dk[dk]:
            slack_solid[dk] = LpVariable(
                f"slack_solid_d{di}",
                lowBound=0,
                cat=LpInteger,
            )
        unmet[dk] = LpVariable(f"unmet_d{di}", lowBound=0, cat=LpInteger)

    for dk in dk_list:
        qty = demand_2d[dk]
        parts = [z_prim[(oid, dk)] for oid in primary_pairs_per_dk[dk]]
        parts += [z_sec[(oid, dk)] for oid in secondary_pairs_per_dk[dk]]
        parts.append(unmet[dk])
        prob += lpSum(parts) == qty, f"demand_d{dk_to_idx[dk]}"

    prim_to_dks: dict = {}
    for _dk, _opts in primary_pairs_per_dk.items():
        for _oid in _opts:
            prim_to_dks.setdefault(_oid, []).append(_dk)
    for _oid, _dks in prim_to_dks.items():
        prob += (
            lpSum(z_prim[(_oid, _dk)] for _dk in _dks) <= x_prim[_oid],
            f"cap_prim_{_oid}",
        )

    sec_to_dks: dict = {}
    for _dk, _opts in secondary_pairs_per_dk.items():
        for _oid in _opts:
            sec_to_dks.setdefault(_oid, []).append(_dk)
    for _oida, _dks in sec_to_dks.items():
        _pieces = secondary_options_by_id[_oida].get("pieces", 1)
        prob += (
            lpSum(z_sec[(_oida, _dk)] for _dk in _dks) <= x_sec[_oida] * _pieces,
            f"cap_sec_{_oida}",
        )

    for _dk, _solid_ids in solid_pairs_per_dk.items():
        if not _solid_ids:
            continue
        _qty = demand_2d[_dk]
        prob += (
            lpSum(z_prim[(oid, _dk)] for oid in _solid_ids) + slack_solid[_dk] >= _qty,
            f"solid_priority_d{dk_to_idx[_dk]}",
        )

    if no_sources_keys:
        import logging as _no_src_summary_log

        _no_src_summary_log.getLogger(__name__).warning(
            "[OPT_2D] no_sources: %d ключей, %d плит — закроются через unmet/post-correction",
            len(no_sources_keys),
            sum(q for _, q in no_sources_keys),
        )

    rests_by_lkey = _build_residual_balance_constraints(
        prob=prob,
        primary_options=primary_options,
        secondary_options=secondary_options,
        x_prim=x_prim,
        x_sec=x_sec,
    )

    M_UNMET = 1e7
    M_SOLID = 1e5
    obj_terms: list = []

    print("[OPT_2D] Расчёт стоимости первичных резов...")
    for opt in primary_options:
        plate_price = get_price(opt["length"], 8, PRICE_DB_PATH) or 10000
        cut_cost = (
            LONG_CUT_PRICE_PER_M * opt["length"] if opt["type"] in ("direct", "indirect") else 0
        )
        obj_terms.append(x_prim[opt["id"]] * (plate_price + cut_cost))

    for opt in secondary_options:
        if opt["type"] in ("narrowing", "multiple", "multiple_transverse"):
            obj_terms.append(x_sec[opt["id"]] * LONG_CUT_PRICE_PER_M * opt["source_length"])
        if opt["type"] in ("transverse", "multiple_transverse"):
            obj_terms.append(x_sec[opt["id"]] * TRANSVERSE_CUT_PRICE)

    for rkey, rec in rests_by_lkey.items():
        if not (rec["produced"] and rec["consumed"]):
            continue
        unused_expr = lpSum(x_prim[i] for i in rec["produced"]) - lpSum(
            x_sec[i] for i in rec["consumed"]
        )
        base_price = get_price(rkey[0], 6, PRICE_DB_PATH) or 5000
        rest_price = base_price * (rkey[1] / float(plate_width))
        obj_terms.append(unused_expr * rest_price * opt_config.unused_rest_penalty_coeff)

    for opt in secondary_options:
        waste_w = opt.get("waste", 0)
        if waste_w > 0:
            waste_area_m2 = (waste_w / 1000.0) * opt["source_length"]
            obj_terms.append(x_sec[opt["id"]] * waste_area_m2 * 1000)
        waste_l = opt.get("length_waste", 0)
        if waste_l > 0:
            waste_area_m2 = (waste_l / 1000.0) * (opt["source_rest"] / 1000.0)
            obj_terms.append(x_sec[opt["id"]] * waste_area_m2 * 1000)

    if opt_config.secondary_reuse_bonus:
        for opt in secondary_options:
            obj_terms.append(x_sec[opt["id"]] * opt_config.secondary_reuse_bonus)

    obj_terms.append(lpSum(x_prim.values()) * 5000.0)

    if slack_solid:
        obj_terms.append(M_SOLID * lpSum(slack_solid.values()))
    if unmet:
        obj_terms.append(M_UNMET * lpSum(unmet.values()))

    print(
        f"[OPT_2D] Конфиг: unused_penalty={opt_config.unused_rest_penalty_coeff}, "
        f"reuse_bonus={opt_config.secondary_reuse_bonus}"
    )
    prob += lpSum(t for t in obj_terms if t != 0)

    return TwoDCuttingILPArtifacts(
        prob=prob,
        x_prim=x_prim,
        x_sec=x_sec,
        z_prim=z_prim,
        z_sec=z_sec,
        slack_solid=slack_solid,
        unmet=unmet,
        dk_list=dk_list,
        dk_to_idx=dk_to_idx,
        primary_pairs_per_dk=primary_pairs_per_dk,
        secondary_pairs_per_dk=secondary_pairs_per_dk,
        solid_pairs_per_dk=solid_pairs_per_dk,
        primary_options_by_id=primary_options_by_id,
        secondary_options_by_id=secondary_options_by_id,
        no_sources_keys=no_sources_keys,
        rests_by_lkey=rests_by_lkey,
    )
