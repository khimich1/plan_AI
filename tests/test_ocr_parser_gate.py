#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import copy

import pytest

from core.ocr.parser_gate import apply_parser_gate


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
