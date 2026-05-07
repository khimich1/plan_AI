# -*- coding: utf-8 -*-
"""Tests for ``core.config.constants`` (REF-001): dimension helpers."""

from __future__ import annotations

import pytest

from core.config.constants import (
    length_dm_to_m,
    normalize_dimension,
    parse_pb_width_to_m,
)


@pytest.mark.parametrize(
    "raw,expected_m",
    [
        # Целый номинал в дм → метры без дробной части в строке
        ("38", 3.8),
        ("69", 6.9),
        ("1", 0.1),
        # Пробелы обрезаются
        ("  75 ", 7.5),
        # Точное значение с разделителем → мм через целочисленную схему, /1000
        ("38,0", 3.8),
        ("38.0", 3.8),
        ("75.5", 7.55),
        ("59,8", 5.98),
        # До трёх знаков дробной части дм
        ("10,12", 1.012),
        # strip() снимает пробелы/перевод строки до разбора
        (" 59,8 ", 5.98),
        ("1,5\n", 0.15),
    ],
)
def test_length_dm_to_m_documented_happy_path(raw: str, expected_m: float) -> None:
    assert length_dm_to_m(raw) == expected_m


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("", 0.0),
        ("   ", 0.0),
        ("abc", 0.0),
        ("12..34", 0.0),  # float() падает → 0.0
    ],
)
def test_length_dm_to_m_invalid_or_empty(raw: str, expected: float) -> None:
    assert length_dm_to_m(raw) == expected


def test_length_dm_to_m_mm_branch_negative_returns_zero() -> None:
    # Через дробную ветку: отрицательная длина в мм отсекается
    assert length_dm_to_m("-1,0") == 0.0


@pytest.mark.parametrize(
    "value_str,expected_dm",
    [
        ("1,2", 1.2),
        ("6.65", 6.65),
        # мм без десятичного разделителя: (20, 1000) → /100
        ("530", 5.3),
        ("665", 6.65),
        ("900", 9.0),
        ("32", 0.32),
        ("86", 0.86),
        ("21", 0.21),  # 20 < 21 < 1000
    ],
)
def test_normalize_dimension_mm_heuristic_and_plain(value_str: str, expected_dm: float) -> None:
    assert normalize_dimension(value_str) == pytest.approx(expected_dm, abs=1e-9)


@pytest.mark.parametrize(
    "value_str,expected_dm",
    [
        # Границы интервала (20; 1000) не включаются
        ("20", 20.0),
        ("1000", 1000.0),
        # Ниже порога — без деления
        ("19.99", 19.99),
    ],
)
def test_normalize_dimension_boundaries_excluded(value_str: str, expected_dm: float) -> None:
    assert normalize_dimension(value_str) == pytest.approx(expected_dm, abs=1e-9)


@pytest.mark.parametrize(
    "value_str",
    ["", "  \t ", "not_a_number"],
)
def test_normalize_dimension_bad_input_zero(value_str: str) -> None:
    assert normalize_dimension(value_str) == 0.0


@pytest.mark.parametrize(
    "width_str,expected_m",
    [
        # Уже метры (спец-ленты)
        ("0.2", 0.2),
        ("0,2", 0.2),
        ("0.3", 0.3),
        ("0,3", 0.3),
        # мм без точки: /100 в normalize → 0.3 дм, дальше не делим на 10
        ("30", 0.3),
        ("200", 0.2),  # 200 мм → 2 дм → /10 = 0.2 м
        # Остальное: дециметры → /10 в метры
        ("3,2", 0.32),
        ("12", 1.2),
        ("2", 0.2),  # 2 дм → 0.2 м (не путаем с литерой 0.2)
    ],
)
def test_parse_pb_width_to_m_special_two_vs_dm_division(
    width_str: str,
    expected_m: float,
) -> None:
    assert parse_pb_width_to_m(width_str) == pytest.approx(expected_m, abs=1e-9)


def test_parse_pb_width_to_m_nonpositive_zero() -> None:
    assert parse_pb_width_to_m("") == 0.0
    assert parse_pb_width_to_m("0") == 0.0
    assert parse_pb_width_to_m("-3") == 0.0
