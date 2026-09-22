#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PLI-006: единый rounding-helper для извлечения qty из решения 2D-ILP."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

pytest.importorskip("pulp")

from core.optimization.optimize_2d.extract_cuts import round_cut_qty  # noqa: E402


def test_round_cut_qty_integer_unchanged(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.WARNING, logger="core.optimization.optimize_2d.extract_cuts")
    assert round_cut_qty(0) == 0
    assert round_cut_qty(2.0) == 2
    assert round_cut_qty(3) == 3
    assert not any("Дробное" in rec.getMessage() for rec in caplog.records)


def test_round_cut_qty_within_eps_no_warning(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.WARNING, logger="core.optimization.optimize_2d.extract_cuts")
    assert round_cut_qty(2.0 + 1e-7) == 2
    assert round_cut_qty(2.0 - 1e-7) == 2
    assert round_cut_qty(1e-7) == 0
    assert not any("Дробное" in rec.getMessage() for rec in caplog.records)


def test_round_cut_qty_fractional_beyond_eps_warns_and_rounds(
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.WARNING, logger="core.optimization.optimize_2d.extract_cuts")
    assert round_cut_qty(1.4, context="z_prim") == 1
    messages = [rec.getMessage() for rec in caplog.records]
    assert any("Дробное" in msg and "1.4" in msg and "z_prim" in msg for msg in messages)


def test_extract_cuts_uses_round_cut_qty_at_three_sites() -> None:
    src = (PROJECT_ROOT / "core/optimization/optimize_2d/extract_cuts.py").read_text(
        encoding="utf-8"
    )
    assert "math.ceil" not in src
    assert src.count("round_cut_qty(") >= 4  # def + три точки извлечения
