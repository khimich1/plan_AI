#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Публичная точка входа оптимизации (OPT-007): валидация (OPT-006) + делегирование в
``core.optimization._implementation`` (бывший монолит, OPT-008).
"""

from __future__ import annotations

from typing import Any

from core.optimization.context import optimization_context_scope
from core.optimization.optimize_1d_widths import _optimize_1d_widths_only
from core.optimization.optimize_2d.with_lengths import _optimize_2d_with_lengths
from core.optimization.ports.order_data import PlateOrderDataPort
from core.optimization.result_contract import ERROR_NO_INPUT, opt_error
from core.optimization.validation import validate_optimize_entrypoint


def optimize_with_cascading_longitudinal_cuts(
    orders: dict | None = None,
    orders_2d: list | None = None,
    plate_width: int = 1200,
    min_useful_width: int = 200,
    opt_config: Any = None,
    order_data: PlateOrderDataPort | None = None,
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

    # OPT-005 / A1: каждый прогон — своё состояние OPT_* (ContextVar + TLS), иначе
    # повторные вызовы на одном потоке (в т.ч. пул asyncio.to_thread) делят словари плана.
    with optimization_context_scope():
        # Точки входа тянем напрямую из модулей 1D/2D (DIP-001): без ленивого импорта пакета
        # и без обратной дуги orchestrator ↔ _implementation.

        if orders_2d is not None and len(orders_2d) > 0:
            print("[OPT] Режим: ПОЛНАЯ 2D оптимизация (длина + ширина)")
            return _optimize_2d_with_lengths(
                orders_2d, plate_width, min_useful_width, opt_config, order_data
            )

        if orders is not None and len(orders) > 0:
            print("[OPT] Режим: 1D оптимизация (только ширина, обратная совместимость)")
            return _optimize_1d_widths_only(orders, plate_width, min_useful_width, order_data)

        print("[OPT] ⚠️ Не указаны ни orders, ни orders_2d!")
        return opt_error(
            ERROR_NO_INPUT,
            "Не переданы orders (1D) и orders_2d (2D): нечего оптимизировать.",
        )
