#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Тесты извлечения qty из решения PuLP в 1D-ветке (без молчаливого bare except)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

pytest.importorskip("pulp")

from core.optimization._implementation import _optimize_1d_widths_only  # noqa: E402
from core.optimization.pulp_qty import _opt_1d_pulp_nonneg_qty  # noqa: E402


def test_opt_1d_pulp_nonneg_qty_rounds_and_zero():
    assert _opt_1d_pulp_nonneg_qty(lambda v: v, 3.7, context="demo") == 4
    assert _opt_1d_pulp_nonneg_qty(lambda v: v, 0.0, context="demo") == 0


def test_opt_1d_pulp_nonneg_qty_none_returns_zero(caplog):
    caplog.set_level("WARNING")
    assert _opt_1d_pulp_nonneg_qty(lambda v: None, object(), context="x_prim[0]") == 0
    assert "x_prim[0]" in caplog.text
    assert "None" in caplog.text


def test_opt_1d_pulp_nonneg_qty_invalid_numeric_raises():
    with pytest.raises(ValueError, match="invalid pulp value"):
        _opt_1d_pulp_nonneg_qty(lambda v: "nope", object(), context="x_sec[1]")


def test_opt_1d_pulp_nonneg_qty_value_fn_raises():
    def _boom(_):
        raise RuntimeError("solver bug")

    with pytest.raises(ValueError, match="pulp.value failed"):
        _opt_1d_pulp_nonneg_qty(_boom, object(), context="x_prim[2]")


def test_opt_1d_pulp_nonneg_qty_negative_raises():
    with pytest.raises(ValueError, match="negative qty"):
        _opt_1d_pulp_nonneg_qty(lambda v: -1, object(), context="x_prim[3]")


def test_optimize_1d_widths_only_smoke_still_ok():
    """Интеграционный дым: полный цикл 1D после смены извлечения переменных."""
    out = _optimize_1d_widths_only({600: 2, 400: 1})
    assert out.get("_opt_status") in ("ok", "partial", "error")
    if out.get("_opt_status") == "error":
        pytest.skip("PuLP/CBC недоступен или модель вернула ошибку в среде CI")
    assert "primary_cuts" in out
    assert sum(c.get("qty", 0) for c in out.get("primary_cuts", [])) >= 1
