#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Параметры настройки целевой функции оптимизации раскроя (штрафы, бонусы)."""

from dataclasses import dataclass


@dataclass
class OptimizationConfig:
    """
    Конфигурация параметров оптимизации.
    Позволяет экспериментировать с разными коэффициентами штрафов и бонусов.
    """
    # Коэффициент штрафа за неиспользованные остатки
    # OLD: 0.5 (50% стоимости остатка)
    # NEW: 0.15 (15% стоимости остатка)
    unused_rest_penalty_coeff: float = 0.15

    # Бонус за использование вторичных резов (отрицательное значение = бонус)
    # OLD: -500 (экономический стимул использовать остатки)
    # NEW: 0 (нет бонуса, остатки используются только если это выгодно)
    secondary_reuse_bonus: float = 0.0


# Дефолтная конфигурация (NEW поведение)
DEFAULT_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.15,
    secondary_reuse_bonus=0.0,
)

# Старая конфигурация (OLD поведение, для экспериментов)
OLD_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.5,
    secondary_reuse_bonus=-500.0,
)
