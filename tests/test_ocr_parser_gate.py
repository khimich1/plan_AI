#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import copy

import pytest

from core.ocr.bridge_pile_parser_gate import apply_bridge_pile_parser_gate
from core.ocr.fbs_parser_gate import apply_fbs_parser_gate
from core.ocr.march_parser_gate import apply_march_parser_gate
from core.ocr.parser_gate import apply_parser_gate
from core.ocr.pile_parser_gate import apply_pile_parser_gate
from core.ocr.step_parser_gate import apply_step_parser_gate


def _plate(**kwargs):
    base = {
        "raw_name": "ПБ 60-12-8п",
        "normalized_candidate": "ПБ 60-12-8п",
        "qty": 1,
        "confidence": 0.95,
        "issues": [],
    }
    base.update(kwargs)
    return base


def test_apply_parser_gate_valid_plates_unchanged():
    plates = [_plate(), _plate(normalized_candidate="ПБ 78-12-8п", qty=2)]
    original = copy.deepcopy(plates)

    result = apply_parser_gate(plates)

    assert result is plates
    assert plates == original


def test_apply_parser_gate_invalid_mark_adds_issue_and_lowers_confidence():
    plates = [_plate(normalized_candidate="Непонятный текст", confidence=0.99)]

    apply_parser_gate(plates)

    assert plates[0]["issues"] == ["parser_rejected"]
    assert plates[0]["confidence"] == 0.5


def _bridge(**kwargs):
    base = {
        "raw_name": "C8-35T1",
        "normalized_candidate": "C8-35T1",
        "qty": 2,
        "confidence": 0.95,
        "issues": [],
    }
    base.update(kwargs)
    return base


def test_bridge_gate_accepts_gost_and_joined_sht():
    items = [
        _bridge(normalized_candidate="C 14.35-T7", qty=80, raw_name="C 14.35-T7"),
        _bridge(normalized_candidate="C14-35T7 80шт", qty=80, raw_name="C14-35T7 80шт"),
    ]
    original_raw = [item["raw_name"] for item in items]

    apply_bridge_pile_parser_gate(items)

    assert items[0]["issues"] == []
    assert items[1]["issues"] == []
    assert [item["raw_name"] for item in items] == original_raw


def test_plate_gate_accepts_joined_sht():
    plates = [_plate(normalized_candidate="ПБ 78-12-8п 5шт", qty=5)]
    apply_parser_gate(plates)
    assert plates[0]["issues"] == []


def test_fbs_step_march_gates_accept_joined_sht():
    fbs = [
        {
            "raw_name": "ФБС 9.3.6-Т 2шт",
            "normalized_candidate": "ФБС 9.3.6-Т 2шт",
            "qty": 2,
            "confidence": 0.95,
            "issues": [],
        }
    ]
    apply_fbs_parser_gate(fbs)
    assert fbs[0]["issues"] == []

    steps = [
        {
            "raw_name": "ЛС11",
            "normalized_candidate": "ЛС11 2шт",
            "qty": 2,
            "confidence": 0.95,
            "issues": [],
        }
    ]
    apply_step_parser_gate(steps)
    assert steps[0]["issues"] == []

    marches = [
        {
            "raw_name": "1ЛМ 27-11-14-4",
            "normalized_candidate": "1ЛМ 27-11-14-4 2шт",
            "qty": 2,
            "confidence": 0.95,
            "issues": [],
        }
    ]
    apply_march_parser_gate(marches)
    assert marches[0]["issues"] == []


def test_apply_parser_gate_caps_existing_low_confidence():
    plates = [_plate(normalized_candidate="???", confidence=0.3)]

    apply_parser_gate(plates)

    assert "parser_rejected" in plates[0]["issues"]
    assert plates[0]["confidence"] == 0.3


def test_apply_parser_gate_does_not_duplicate_parser_rejected():
    plates = [_plate(normalized_candidate="???", issues=["parser_rejected", "other"])]

    apply_parser_gate(plates)

    assert plates[0]["issues"].count("parser_rejected") == 1


def test_apply_parser_gate_preserves_existing_issues():
    plates = [_plate(normalized_candidate="???", issues=["prefix_separator_dot"])]

    apply_parser_gate(plates)

    assert plates[0]["issues"] == ["prefix_separator_dot", "parser_rejected"]


def test_apply_parser_gate_empty_list():
    plates = []
    assert apply_parser_gate(plates) == []


def test_apply_parser_gate_uses_raw_name_fallback():
    plates = [
        {
            "raw_name": "ПБ 60-12-8п",
            "qty": 1,
            "confidence": 0.95,
            "issues": [],
        }
    ]

    apply_parser_gate(plates)

    assert plates[0]["issues"] == []


def _pile(**kwargs):
    base = {
        "raw_name": "С90.30-11",
        "normalized_candidate": "С90.30-11",
        "qty": 189,
        "confidence": 0.95,
        "issues": [],
    }
    base.update(kwargs)
    return base


def test_apply_pile_parser_gate_valid_piles_unchanged():
    piles = [_pile(), _pile(normalized_candidate="С110.30-13", qty=26)]
    original = copy.deepcopy(piles)

    result = apply_pile_parser_gate(piles)

    assert result is piles
    assert piles == original


def test_apply_pile_parser_gate_invalid_mark_adds_issue():
    piles = [_pile(normalized_candidate="Непонятный текст", confidence=0.99)]

    apply_pile_parser_gate(piles)

    assert piles[0]["issues"] == ["parser_rejected"]
    assert piles[0]["confidence"] == 0.5
