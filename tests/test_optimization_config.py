"""Регрессия для OPT-REF-001: конфиг целевой функции вынесен в optimization_config.py.

Проверяем публичный импорт с уровня `core.optimization` и стабильные значения
дефолтов (NEW vs OLD), чтобы рефакторинг не ломал контракт для вызывающего кода.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.optimization as optimization_pkg  # noqa: E402
from core.optimization import (  # noqa: E402
    DEFAULT_CONFIG,
    OLD_CONFIG,
    OptimizationConfig,
)
from core.optimization import optimization_config as opt_cfg_mod  # noqa: E402


def test_optimization_config_importable_from_package() -> None:
    assert optimization_pkg.OptimizationConfig is OptimizationConfig
    assert optimization_pkg.DEFAULT_CONFIG is DEFAULT_CONFIG
    assert optimization_pkg.OLD_CONFIG is OLD_CONFIG


def test_package_reexports_same_objects_as_optimization_config_module() -> None:
    """Один источник правды: пакет не должен дублировать экземпляры/класс."""
    assert DEFAULT_CONFIG is opt_cfg_mod.DEFAULT_CONFIG
    assert OLD_CONFIG is opt_cfg_mod.OLD_CONFIG
    assert OptimizationConfig is opt_cfg_mod.OptimizationConfig


def test_default_config_values() -> None:
    assert DEFAULT_CONFIG.unused_rest_penalty_coeff == pytest.approx(0.15)
    assert DEFAULT_CONFIG.secondary_reuse_bonus == pytest.approx(0.0)


def test_old_config_values() -> None:
    assert OLD_CONFIG.unused_rest_penalty_coeff == pytest.approx(0.5)
    assert OLD_CONFIG.secondary_reuse_bonus == pytest.approx(-500.0)


def test_optimization_config_dataclass_defaults_match_default_config() -> None:
    """Явные поля DEFAULT_CONFIG совпадают с дефолтами dataclass (NEW-поведение)."""
    fresh = OptimizationConfig()
    assert fresh.unused_rest_penalty_coeff == DEFAULT_CONFIG.unused_rest_penalty_coeff
    assert fresh.secondary_reuse_bonus == DEFAULT_CONFIG.secondary_reuse_bonus
