#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Публичная точка входа оптимизации (OPT-007): валидация (OPT-006) + делегирование в
``core.optimization._implementation`` (бывший монолит, OPT-008).
"""

from __future__ import annotations

from typing import Any

from core.optimization.result_contract import ERROR_NO_INPUT, opt_error
from core.optimization.validation import validate_optimize_entrypoint


def optimize_with_cascading_longitudinal_cuts(
    orders: dict | None = None,
    orders_2d: list | None = None,
    plate_width: int = 1200,
    min_useful_width: int = 200,
    opt_config: Any = None,
) -> dict:
    """
    Универсальная оптимизация с каскадными резами (PUBLIC API).

    Возвращает словарь плана полей (`primary_cuts`, …) или структурированную
    ошибку: ключ ``_opt_status`` == ``\"error\"`` и ``_opt_error_code``
    (`no_input`, `pulp_missing`, `solver_infeasible`, …). Успешный ответ
    помечается ``_opt_status``: ``\"ok\"`` или ``\"partial\"`` (неоптимальный
    статус решателя CBC, но допустимый план).
    """
    validate_optimize_entrypoint(
        orders=orders,
        orders_2d=orders_2d,
        plate_width=plate_width,
        min_useful_width=min_useful_width,
    )

    # Ленивый импорт пакета — избегаем циклов при загрузке оркестратора и реализации.
    import core.optimization as pkg

    if orders_2d is not None and len(orders_2d) > 0:
        print("[OPT] Режим: ПОЛНАЯ 2D оптимизация (длина + ширина)")
        return pkg._optimize_2d_with_lengths(orders_2d, plate_width, min_useful_width, opt_config)

    if orders is not None and len(orders) > 0:
        print("[OPT] Режим: 1D оптимизация (только ширина, обратная совместимость)")
        return pkg._optimize_1d_widths_only(orders, plate_width, min_useful_width)

    print("[OPT] ⚠️ Не указаны ни orders, ни orders_2d!")
    return opt_error(
        ERROR_NO_INPUT,
        "Не переданы orders (1D) и orders_2d (2D): нечего оптимизировать.",
    )
