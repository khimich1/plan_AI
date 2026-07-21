#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Legacy width-plan adapters: fill OPT_WIDTH_PRIORITY and OPT_PLAN['actions']."""

from __future__ import annotations

from collections import Counter

from core.optimization.context import OPT_PLAN, OPT_WIDTH_PRIORITY
from core.plate_runtime_state import get_plate_mutable_runtime
from core.optimization.result_contract import (
    ERROR_EMPTY_ORDERS_1D,
    is_optimization_success,
    opt_error,
)


def _group_plate_lengths(plates: list[float]) -> dict[float, int]:
    """Группируем длины (в метрах) -> количество."""
    return Counter(round(float(length), 2) for length in (plates or []))


def _append_actions(actions: list, width_mm: int, lengths: dict, long_cuts: int, src_type: str):
    """Добавляем aggregated записи в OPT_PLAN['actions']."""
    rest_mm = max(0, 1200 - width_mm)
    for length_m, qty in lengths.items():
        actions.append((
            src_type,
            width_mm,
            rest_mm if src_type != 'solid' else 0,
            length_m,
            qty,
            long_cuts,
            0  # поперечные резы в legacy-плане не моделируем
        ))


def apply_width_optimization() -> dict:
    """
    Упрощённый наследуемый оптимизатор ширин.
    Наполняет OPT_WIDTH_PRIORITY и OPT_PLAN['actions'] из текущего plate runtime.
    """
    rt = get_plate_mutable_runtime()
    priority = []
    actions = []

    priority_groups = [
        ('0_32', 320, rt.plates_0_32),
        ('0_46', 460, rt.plates_0_46),
        ('0_70', 700, rt.plates_0_70),
        ('0_72', 720, rt.plates_0_72),
        ('0_86', 860, rt.plates_0_86),
        ('0_74', 740, rt.plates_0_74),
        ('0_88', 880, rt.plates_0_88),
        ('0_48', 480, rt.plates_0_48),
        ('0_50', 500, rt.plates_0_50),
        ('0_34', 340, rt.plates_0_34),
    ]

    for code, width_mm, plate_list in priority_groups:
        if not plate_list:
            continue
        priority.append(code)
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    solid_groups = [
        (1200, rt.plates_1_2),
    ]
    split_groups = [
        (1080, rt.plates_1_08),
        (1000, rt.plates_1_0),
    ]

    for width_mm, plate_list in solid_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=0, src_type='solid')

    for width_mm, plate_list in split_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    OPT_WIDTH_PRIORITY.clear()
    OPT_WIDTH_PRIORITY.extend(priority)

    OPT_PLAN.clear()
    OPT_PLAN.update({
        'orders': {width: len(lst) for (_, width, lst) in priority_groups if lst},
        'actions': actions
    })

    return OPT_PLAN


def optimize_cuts_pulp(orders: dict | None = None) -> dict:
    """
    Legacy-обёртка над новой каскадной оптимизацией (для совместимости с визуализатором).
    """
    if orders is None:
        rt = get_plate_mutable_runtime()
        orders = {}
        for width_mm, plates in [
            (320, rt.plates_0_32), (460, rt.plates_0_46), (700, rt.plates_0_70),
            (720, rt.plates_0_72), (860, rt.plates_0_86), (880, rt.plates_0_88),
            (740, rt.plates_0_74), (480, rt.plates_0_48), (500, rt.plates_0_50),
            (340, rt.plates_0_34)
        ]:
            if plates:
                orders[width_mm] = len(plates)

    if not orders:
        return opt_error(
            ERROR_EMPTY_ORDERS_1D,
            "Нет исходных заказов по ширине для optimize_cuts_pulp.",
        )
    # Локальный импорт: избегаем цикла при загрузке legacy_width_plan из _implementation (DIP-001).
    from core.optimization.orchestrator import optimize_with_cascading_longitudinal_cuts

    result = optimize_with_cascading_longitudinal_cuts(orders=orders)
    if is_optimization_success(result):
        OPT_PLAN.clear()
        OPT_PLAN.update(result)
    return result
