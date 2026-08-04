#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unit tests for pile OCR text normalizer (R1–R4)."""

from __future__ import annotations

from core.pile_line_parser import parse_pile_line
from core.pile_text_normalizer import normalize_pile_order_text

PILOT_LINES = [
    "Сваи 90.30-11 189",
    "Свай 110.30-13 26",
    "Свай 120.30-12 20",
]

EXPECTED = [
    ("С90.30-11", 189),
    ("С110.30-13", 26),
    ("С120.30-12", 20),
]


def test_pilot_strings_r1_r2():
    result = normalize_pile_order_text("\n".join(PILOT_LINES))
    assert result.normalized_lines == [
        "С90.30-11 189",
        "С110.30-13 26",
        "С120.30-12 B25 20",
    ]


def test_r3_strip_sht():
    result = normalize_pile_order_text("Сваи 90.30-11 189 шт")
    assert result.normalized_lines == ["С90.30-11 189"]


def test_r4_dash_and_whitespace_cleanup():
    result = normalize_pile_order_text("Сваи  90.30–11   189")
    assert result.normalized_lines == ["С90.30-11 189"]


def test_pilot_strings_parse_after_normalizer():
    result = normalize_pile_order_text("\n".join(PILOT_LINES))
    for line, (mark, qty) in zip(result.normalized_lines, EXPECTED):
        parsed = parse_pile_line(line)
        assert parsed.parsed is True
        assert parsed.mark == mark
        assert parsed.qty == qty
        assert parsed.concrete_grade == "B25"


def test_ocr_disambiguate_qty_20_after_svay_prefix():
    result = normalize_pile_order_text("Свай 120.30-12 20")
    parsed = parse_pile_line(result.normalized_lines[0])
    assert parsed.parsed is True
    assert parsed.qty == 20
    assert parsed.concrete_grade == "B25"


def test_idempotent_on_clean_marks():
    clean = "С90.30-11 189\nС110.30-13 26"
    once = normalize_pile_order_text(clean)
    twice = normalize_pile_order_text(once.normalized_text)
    assert once.normalized_text == twice.normalized_text
