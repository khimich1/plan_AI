#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Регрессия для XLSX-fallback производственной сметы: ключ длины через length_m_to_price_length_dm (ceil),
в отличие от find_price_for_plate (round), чтобы совпадать с прайсом при краевых длинах.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.price_db import length_m_to_price_length_dm  # noqa: E402
from viz_modules.price_utils import find_price_for_plate  # noqa: E402
from viz_modules.procurement import _find_price_for_plate_production_fallback  # noqa: E402


def test_production_fallback_matches_length_m_to_price_length_dm_key_exact() -> None:
    """Точное попадание в ключ таблицы = ceil-дм для метража."""
    length_m = 2.73
    dm = length_m_to_price_length_dm(length_m)
    price_table = {dm: {8: 42_600.0}}
    got = _find_price_for_plate_production_fallback(price_table, length_m, 8)
    assert got == 42_600.0


def test_production_fallback_finds_nearby_dm_within_one() -> None:
    dm = length_m_to_price_length_dm(5.5)
    price_table = {dm + 1: {10: 99_999.0}}
    got = _find_price_for_plate_production_fallback(price_table, 5.5, 10)
    assert got == 99_999.0


def test_production_fallback_vs_legacy_round_key_gap_gt_one() -> None:
    """
    При ключе только на dm=30 round-для 2.81 даёт 28 → расстояние до 30 больше 1 → legacy None,
    production берёт ключ ceil (=29) и находит соседа 30.
    """
    length_m = 2.81
    price_table = {30: {8: 1000.0}}
    assert length_m_to_price_length_dm(length_m) == 29
    assert int(round(length_m * 10)) == 28

    assert find_price_for_plate(price_table, length_m, 8) is None
    assert _find_price_for_plate_production_fallback(price_table, length_m, 8) == 1000.0


def test_production_fallback_load_code_floor_matches_legacy() -> None:
    price_table = {55: {12: 77_777.0}}
    got_production = _find_price_for_plate_production_fallback(price_table, 5.5, 12.5)
    got_legacy = find_price_for_plate(price_table, 5.5, 12.5)
    assert got_production == got_legacy == 77_777.0


def test_production_fallback_invalid_load_defaults_to_eight() -> None:
    price_table = {27: {8: 1.0}}
    assert _find_price_for_plate_production_fallback(price_table, 2.7, "bad") == 1.0


def test_production_fallback_returns_none_when_no_match() -> None:
    assert _find_price_for_plate_production_fallback({}, 6.0, 8) is None
