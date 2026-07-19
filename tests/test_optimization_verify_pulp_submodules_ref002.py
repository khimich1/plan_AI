#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OPT-REF-002: проверки извлечённых модулей `coverage_verify` и `pulp_qty`.

Не дублируем сценарии из `test_opt_1d_pulp_qty_extraction.py` / baseline — здесь только
канонический импорт и равенство объектов между подмодулями и пакетом / `_implementation`.
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.optimization as optimization_pkg  # noqa: E402
from core.optimization import verify_coverage as verify_from_package  # noqa: E402
from core.optimization._implementation import (  # noqa: E402
    _opt_1d_pulp_nonneg_qty as pulp_nonneg_from_implementation,
    verify_coverage as verify_from_implementation,
)
from core.optimization.coverage_verify import verify_coverage as verify_from_submodule  # noqa: E402
from core.optimization.pulp_qty import _opt_1d_pulp_nonneg_qty as pulp_nonneg_from_submodule  # noqa: E402


def test_verify_coverage_submodule_same_object_as_package_and_implementation() -> None:
    assert verify_from_submodule is verify_from_package
    assert verify_from_submodule is verify_from_implementation
    assert optimization_pkg.verify_coverage is verify_from_submodule
    mod = importlib.import_module("core.optimization.coverage_verify")
    assert mod.verify_coverage is verify_from_package


def test_verify_coverage_one_shot_call_via_submodule() -> None:
    """Минимальный вызов через прямой импорт (без дублирования baseline)."""
    demand = {(800, 400, 2): 1}
    primary = [{"assignment_key": (800, 400, 2)}]
    out = verify_from_submodule(demand, primary, [])
    assert out["ok"] is True
    assert out["missing"] == {}


def test_pulp_qty_nonneg_fn_same_object_as_implementation() -> None:
    """Функция не в `__all__`, но должна оставаться одним объектом с `_implementation`."""
    assert pulp_nonneg_from_submodule is pulp_nonneg_from_implementation


def test_pulp_qty_nonneg_minimal_roundtrip_not_duplicate_extraction_suite() -> None:
    """Один сквозной sanity-check; детальные ветки — в `test_opt_1d_pulp_qty_extraction.py`."""
    assert pulp_nonneg_from_submodule(lambda v: v, 2.0, context="parity") == 2
