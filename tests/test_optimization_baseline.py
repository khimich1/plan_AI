#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baseline-тесты ILP-оптимизатора раскроя.

Цель — зафиксировать инварианты текущего поведения перед переходом на
assignment-модель, чтобы регрессии было видно сразу:

1. verify_coverage(demand, primary_cuts, secondary_cuts).ok == True
   (после оптимизатора и его post-correction плита не должна теряться).
2. total_plates >= demand_total после деления secondary на pieces
   (физическая реализуемость раскроя).
3. Исторически нестабильные ключи (4.78/4.79, 5.98/665, 5.08/320, 5.98/530)
   гарантированно покрыты.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Импортируем то, что тестируем
from core.optimization.result_contract import is_optimization_success
from core.optimization import (  # noqa: E402
    _build_residual_balance_constraints,
    optimize_with_cascading_longitudinal_cuts,
    verify_coverage,
)
from core.optimization.geometry import (  # noqa: E402
    GeometryConfig,
    generate_primary_cut_options_2d,
    generate_secondary_cut_options_2d,
)
from core.config_and_data import canonical_plate_key  # noqa: E402

pulp = pytest.importorskip("pulp")


# ---------- helper'ы ----------------------------------------------------------

def _demand_from_orders(orders_2d: list[dict]) -> dict:
    """Сворачивает orders_2d в demand_2d (ключи через canonical_plate_key)."""
    demand: dict = {}
    for o in orders_2d:
        key = canonical_plate_key(o["length"], o["width"], o.get("load_code", 8))
        demand[key] = demand.get(key, 0) + int(o.get("qty", 1))
    return demand


def _physical_units(result: dict) -> int:
    """
    Число физических плит-юнитов: primary count + сумма secondary * pieces.
    Должно покрывать demand_total (>=).
    """
    units = 0
    for cut in result.get("primary_cuts", []) or []:
        units += int(cut.get("qty", 1) or 1)
    for cut in result.get("secondary_cuts", []) or []:
        units += int(cut.get("qty", 1) or 1) * int(cut.get("pieces", 1) or 1)
    return units


# ---------- unit: verify_coverage --------------------------------------------

def test_verify_coverage_ok_when_full_match():
    demand = {(6.0, 1200, 8): 3}
    primary = [
        {"assignment_key": (6.0, 1200, 8)},
        {"assignment_key": (6.0, 1200, 8)},
        {"assignment_key": (6.0, 1200, 8)},
    ]
    secondary: list = []
    cov = verify_coverage(demand, primary, secondary)
    assert cov["ok"] is True
    assert cov["missing"] == {}
    assert cov["demand_total"] == 3
    assert cov["covered_total"] == 3


def test_verify_coverage_detects_missing():
    demand = {(6.0, 1200, 8): 5}
    primary = [{"assignment_key": (6.0, 1200, 8)}, {"assignment_key": (6.0, 1200, 8)}]
    cov = verify_coverage(demand, primary, [])
    assert cov["ok"] is False
    assert cov["missing"][(6.0, 1200, 8)] == 3


def test_verify_coverage_detects_surplus():
    demand = {(6.0, 1200, 8): 1}
    primary = [{"assignment_key": (6.0, 1200, 8)}, {"assignment_key": (6.0, 1200, 8)}]
    cov = verify_coverage(demand, primary, [])
    assert cov["ok"] is True  # дефицита нет
    assert cov["surplus"][(6.0, 1200, 8)] == 1


def test_verify_coverage_handles_load_code_normalization():
    # 800 и 8 должны нормализоваться в один ключ через canonical_plate_key
    demand = {(6.0, 1200, 800): 2}
    primary = [
        {"assignment_key": (6.0, 1200, 8)},
        {"assignment_key": (6.0, 1200, 8)},
    ]
    cov = verify_coverage(demand, primary, [])
    assert cov["ok"] is True
    assert cov["missing"] == {}


def test_residual_balance_normalizes_load_code_and_blocks_ghost_secondary():
    prob = pulp.LpProblem("residual_balance_normalized", pulp.LpMaximize)
    primary_options = [
        {"id": 1, "length": 7.3, "rest": 880, "type": "direct", "load_code": 800}
    ]
    secondary_options = [
        {
            "id": 10,
            "source_length": 7.3,
            "source_rest": 880,
            "target_order_key": (7.3, 720, 8),
        }
    ]
    x_prim = {1: pulp.LpVariable("p1", lowBound=0, upBound=0, cat=pulp.LpInteger)}
    x_sec = {10: pulp.LpVariable("s10", lowBound=0, cat=pulp.LpInteger)}

    _build_residual_balance_constraints(
        prob=prob,
        primary_options=primary_options,
        secondary_options=secondary_options,
        x_prim=x_prim,
        x_sec=x_sec,
    )
    prob += x_sec[10]
    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    assert pulp.value(x_sec[10]) == 0


def test_residual_balance_allows_only_load_code_downgrade():
    prob = pulp.LpProblem("residual_balance_downgrade", pulp.LpMaximize)
    primary_options = [
        {"id": 1, "length": 7.3, "rest": 880, "type": "direct", "load_code": 10},
        {"id": 2, "length": 8.4, "rest": 480, "type": "direct", "load_code": 8},
    ]
    secondary_options = [
        {
            "id": 10,
            "source_length": 7.3,
            "source_rest": 880,
            "target_order_key": (7.3, 720, 8),
        },
        {
            "id": 11,
            "source_length": 8.4,
            "source_rest": 480,
            "target_order_key": (8.4, 320, 10),
        },
    ]
    x_prim = {
        1: pulp.LpVariable("p1", lowBound=1, upBound=1, cat=pulp.LpInteger),
        2: pulp.LpVariable("p2", lowBound=1, upBound=1, cat=pulp.LpInteger),
    }
    x_sec = {
        10: pulp.LpVariable("s10", lowBound=0, cat=pulp.LpInteger),
        11: pulp.LpVariable("s11", lowBound=0, cat=pulp.LpInteger),
    }

    _build_residual_balance_constraints(
        prob=prob,
        primary_options=primary_options,
        secondary_options=secondary_options,
        x_prim=x_prim,
        x_sec=x_sec,
    )
    prob += x_sec[10] + x_sec[11]
    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    assert pulp.value(x_sec[10]) == 1
    assert pulp.value(x_sec[11]) == 0


def test_geometry_generates_representative_narrowing_option():
    demand = {
        (6.0, 720, 8): 1,
        (6.0, 460, 8): 1,
    }

    primary_result = generate_primary_cut_options_2d(
        demand_2d=demand,
        order_info_list={},
        order_info_getter=lambda _order_info_list, _key: {},
        config=GeometryConfig(plate_width=1200, min_useful_width=200, tolerance_width=20),
    )
    secondary_options = generate_secondary_cut_options_2d(
        primary_options=primary_result.options,
        demand_2d=demand,
        config=GeometryConfig(plate_width=1200, min_useful_width=200, tolerance_width=20),
    )

    assert any(
        opt["type"] == "narrowing"
        and opt["source_rest"] == 480
        and opt["output_width"] == 460
        for opt in secondary_options
    )


def test_geometry_secondary_pieces_capped_when_primary_already_emits_same_width():
    """
    Плита 1200: первичный direct 300+900 даёт одну 300 мм; из остатка 900 недопустимо
    нарезать три 300 мм (1+3=4 с одной базовой) — максимум две в вторичке (итого 3).
    """
    demand = {(6.0, 300, 8): 4}
    cfg = GeometryConfig(plate_width=1200, min_useful_width=200, tolerance_width=20)
    primary_result = generate_primary_cut_options_2d(
        demand_2d=demand,
        order_info_list={},
        order_info_getter=lambda _order_info_list, _key: {},
        config=cfg,
    )
    secondary_options = generate_secondary_cut_options_2d(
        primary_options=primary_result.options,
        demand_2d=demand,
        config=cfg,
    )
    bad = [
        o
        for o in secondary_options
        if o.get("type") == "multiple"
        and int(o.get("source_rest") or 0) == 900
        and int(o.get("output_width") or 0) == 300
        and int(o.get("pieces") or 1) >= 3
    ]
    assert not bad, f"unexpected 3+ secondary strips from 900 after primary 300: {bad}"


# ---------- integration: golden cases ----------------------------------------

@pytest.mark.parametrize(
    "case_name, orders_2d",
    [
        # Исторический фикс из docs/plate-loss-fix-summary.md:
        # конкурирующие длины 4.78 и 4.79 м, ширина 1200 мм (solid).
        (
            "competing_lengths_4_78_4_79",
            [
                {"length": 4.78, "width": 1200, "qty": 8, "load_code": 8, "kp_id": 1},
                {"length": 4.79, "width": 1200, "qty": 17, "load_code": 8, "kp_id": 1},
            ],
        ),
        # 5.98/665 — точечный safeguard demand_598665_min покрывает этот ключ.
        (
            "key_5_98_665",
            [
                {"length": 5.98, "width": 665, "qty": 4, "load_code": 8, "kp_id": 1},
            ],
        ),
        # 5.08/320 — РЕСКЬЮ-ключ из debug-95694e
        (
            "key_5_08_320",
            [
                {"length": 5.08, "width": 320, "qty": 6, "load_code": 8, "kp_id": 1},
            ],
        ),
        # 5.98/530 — второй РЕСКЬЮ-ключ из debug-95694e
        (
            "key_5_98_530",
            [
                {"length": 5.98, "width": 530, "qty": 4, "load_code": 8, "kp_id": 1},
            ],
        ),
        # Простой smoke-кейс (один solid).
        (
            "simple_solid",
            [
                {"length": 6.0, "width": 1200, "qty": 3, "load_code": 8, "kp_id": 1},
            ],
        ),
    ],
)
def test_optimizer_covers_demand_for_golden_cases(case_name, orders_2d):
    """
    Главный инвариант: после оптимизатора demand покрыт полностью.
    verify_coverage.ok должен быть True (включая post-correction safety net).
    """
    result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
    assert is_optimization_success(result), f"{case_name}: оптимизатор не вернул успешный план"

    demand = _demand_from_orders(orders_2d)
    cov = verify_coverage(
        demand,
        result.get("primary_cuts", []),
        result.get("secondary_cuts", []),
    )
    assert cov["ok"], (
        f"{case_name}: missing={cov['missing']}, "
        f"demand_total={cov['demand_total']}, covered={cov['covered_total']}"
    )


@pytest.mark.parametrize(
    "case_name, orders_2d, max_units_factor",
    [
        # Для simple_solid: ровно qty solid-плит, без раздутия.
        ("simple_solid", [{"length": 6.0, "width": 1200, "qty": 3, "load_code": 8}], 1.0),
        # Для 5.08/320: одна плита 1200мм даёт 3 куска по 320мм, поэтому
        # достаточно ceil(6/3)=2 первичных плит. Допустим запас x2.
        ("key_5_08_320", [{"length": 5.08, "width": 320, "qty": 6, "load_code": 8}], 2.0),
    ],
)
def test_optimizer_does_not_overproduce(case_name, orders_2d, max_units_factor):
    """
    Snapshot-инвариант: total_plates не должен катастрофически раздуваться
    относительно demand_total. Это защита от случайной регрессии.
    """
    result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
    assert is_optimization_success(result), f"{case_name}: оптимизатор не вернул успешный план"

    demand = _demand_from_orders(orders_2d)
    demand_total = sum(demand.values())

    units = _physical_units(result)
    # Покрытие не должно превышать demand_total в max_units_factor раз.
    assert units <= max(demand_total * max_units_factor, 2), (
        f"{case_name}: оптимизатор произвёл слишком много плит — "
        f"{units} при спросе {demand_total} (max factor {max_units_factor})"
    )


def test_optimizer_competing_lengths_no_loss():
    """
    Регрессия из docs/plate-loss-fix-summary.md: 4.78×8 + 4.79×17 = 25 плит.
    После фикса должно производиться >= 25 единиц.
    """
    orders_2d = [
        {"length": 4.78, "width": 1200, "qty": 8, "load_code": 8, "kp_id": 1},
        {"length": 4.79, "width": 1200, "qty": 17, "load_code": 8, "kp_id": 1},
    ]
    result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
    assert is_optimization_success(result), "оптимизатор не вернул успешный план"

    demand = _demand_from_orders(orders_2d)
    cov = verify_coverage(
        demand,
        result.get("primary_cuts", []),
        result.get("secondary_cuts", []),
    )
    assert cov["ok"], f"конкурирующие длины: missing={cov['missing']}"

    units = _physical_units(result)
    assert units >= sum(demand.values()), (
        f"конкурирующие длины: получено {units} единиц, нужно минимум "
        f"{sum(demand.values())} (баг shared-sources снова регрессировал?)"
    )
