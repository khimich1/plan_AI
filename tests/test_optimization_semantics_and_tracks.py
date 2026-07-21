#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Дополнительные тесты для core.optimization.py.

Цели:
1. Зафиксировать контракт verify_coverage vs поля qty/pieces (частый источник
   ложных «багов раскроя», когда смешивают агрегированные строки плана и
   посчёт физических плит).
2. Покрыть чистые хелперы (_canonical_length) и дорожечный FFD без PuLP.

Интеграция ILP уже в tests/test_optimization_baseline.py.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.optimization import (  # noqa: E402
    Piece,
    Track,
    first_fit_decreasing,
    optimize_tracks,
    optimize_with_cascading_longitudinal_cuts,
    verify_coverage,
)
from core.optimization.geometry import _canonical_length  # noqa: E402
from core.optimization.result_contract import ERROR_NO_INPUT  # noqa: E402


class TestCanonicalLength:
    """Канонические длины для ILP."""

    def test_round_trip_two_decimals(self) -> None:
        assert _canonical_length(5.678) == 5.68
        assert _canonical_length("6.001") == 6.0

    def test_invalid_returns_zero(self) -> None:
        assert _canonical_length(None) == 0.0
        assert _canonical_length("not-a-float") == 0.0


class TestVerifyCoverageSemantics:
    """
    verify_coverage считает по одному «слоту покрытия» на каждую запись списков
    primary/secondary с валидным ключом assignment_key / target_order_key.
    Поле qty строки игнорируется — см. код verify_coverage в optimization.py.
    """

    key = (6.0, 1200, 8)

    def test_single_primary_row_always_counts_as_one_even_with_qty_gt_one(self) -> None:
        demand = {self.key: 5}
        primary = [{"assignment_key": self.key, "qty": 5}]
        cov = verify_coverage(demand, primary, [])
        assert cov["ok"] is False
        assert cov["missing"][self.key] == 4
        assert cov["covered_total"] == 1

    def test_secondary_row_counts_once_even_if_qty_times_pieces_would_need_more(self) -> None:
        """
        Если вручную собрать агрегированную вторичную запись с qty*pieces >> 1,
        verify_coverage всё равно добавляет +1 по target_order_key —
        расходится с _physical_units() из baseline, который умножает qty*pieces.

        Такой кейс маловероятен для реального вывода _optimize_2d_with_lengths,
        там вторичные строки расшиваются по одной плите (qty=pieces=1).
        Тест нужен как сигнал при смешении 1D-агрегатов или ручном JSON.
        """
        demand = {self.key: 6}
        primary: list = []
        secondary = [
            {
                "target_order_key": self.key,
                "qty": 2,
                "pieces": 3,
            }
        ]
        cov = verify_coverage(demand, primary, secondary)
        assert cov["covered_total"] == 1
        assert cov["ok"] is False
        missing = cov["missing"][self.key]
        assert missing == 5

    def test_secondary_without_target_order_key_contributes_zero(self) -> None:
        """Типичный вывод 1D-пути вторичных резов — без target_order_key."""
        demand = {self.key: 1}
        primary = [{"assignment_key": self.key}]
        secondary = [{"qty": 3, "pieces": 3, "source": 880, "cuts": [320]}]
        cov = verify_coverage(demand, primary, secondary)
        assert cov["covered_total"] == 1
        assert cov["ok"] is True

    def test_mixed_primary_and_secondary_keys(self) -> None:
        a = (5.5, 1200, 8)
        b = (5.5, 320, 8)
        demand = {a: 1, b: 1}
        primary = [{"assignment_key": a}]
        secondary = [{"target_order_key": b}]
        cov = verify_coverage(demand, primary, secondary)
        assert cov["ok"] is True
        assert cov["demand_total"] == 2
        assert cov["covered_total"] == 2


class TestOptimizeWithCascadingEmptyPublicApi:
    def test_empty_inputs_yield_structured_error(self, caplog) -> None:
        with caplog.at_level("WARNING", logger="core.optimization.orchestrator"):
            out = optimize_with_cascading_longitudinal_cuts(orders=None, orders_2d=None)
        assert out.get("_opt_status") == "error"
        assert out.get("_opt_error_code") == ERROR_NO_INPUT

        messages = " ".join(r.getMessage() for r in caplog.records)
        assert "Не указаны" in messages or "orders" in messages.lower()


class TestFirstFitDecreasingAndTracks:
    def test_ffd_splits_when_sum_exceeds_stock(self) -> None:
        pieces = [Piece(length_m=6.0, qty=2, kind="standard", load_class=8.0)]
        tracks = first_fit_decreasing(pieces, stock_len_m=9.88)
        assert len(tracks) == 2
        assert all(t.leftover_m >= 0 for t in tracks)
        total_placed = sum(t.total_m for t in tracks)
        assert pytest.approx(total_placed, rel=1e-6) == 12.0

    def test_ffd_prefers_packed_layout(self) -> None:
        pieces = [
            Piece(length_m=5.0, qty=1, kind="standard", load_class=8.0),
            Piece(length_m=4.0, qty=1, kind="standard", load_class=8.0),
        ]
        tracks = first_fit_decreasing(pieces, stock_len_m=9.88)
        assert len(tracks) == 1
        assert tracks[0].total_m == 9.0

    def test_optimize_tracks_dict_structure(self) -> None:
        res = optimize_tracks(
            [
                {"length_m": 3.0, "qty": 2, "kind": "standard", "load_class": 8.0},
            ],
            stock_len_m=9.88,
        )
        assert res["total_tracks"] == 1
        assert isinstance(res["tracks"], list)
        assert all(isinstance(t, Track) for t in res["tracks"])
        assert "efficiency_pct" in res
        assert res["stock_length_m"] == 9.88
