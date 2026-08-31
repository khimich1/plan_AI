#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from core.ocr.result import build_result_payload
from core.ocr.verify_apply import select_ocr_items


def _item(name: str, qty: int = 1) -> dict:
    return {
        "raw_name": name,
        "normalized_candidate": name,
        "qty": qty,
        "confidence": 0.95,
        "issues": [],
    }


def _gate(items: list[dict]) -> list[dict]:
    gated = []
    for item in items:
        copy = dict(item)
        copy["gated"] = True
        gated.append(copy)
    return gated


def test_empty_corrections_keeps_extract_even_if_plates_differ() -> None:
    extract = [_item("ПБ 36-12-8п")]
    verified = [_item("ПБ 63-12-8п")]
    result = select_ocr_items(
        extract,
        {"plates": verified, "corrections": [], "row_count_on_image": 1},
        _gate,
    )

    assert result.items == extract
    assert result.verify_failed is False
    assert result.select_reason == "kept_extract_empty_corrections"


def test_empty_corrections_and_different_row_count_still_keeps_extract() -> None:
    extract = [_item("ПБ 60-12-8п"), _item("ПБ 72-12-8п")]
    verified = [_item("ПБ 60-12-8п")]
    result = select_ocr_items(
        extract,
        {"plates": verified, "corrections": [], "row_count_on_image": 10},
        _gate,
    )

    assert result.items == extract
    assert result.verify_failed is False
    assert result.select_reason == "kept_extract_empty_corrections"


def test_missing_corrections_treated_as_empty() -> None:
    extract = [_item("ПБ 36-12-8п")]
    result = select_ocr_items(
        extract,
        {"plates": [_item("ПБ 63-12-8п")]},
        _gate,
    )

    assert result.items == extract
    assert result.select_reason == "kept_extract_empty_corrections"


def test_nonempty_corrections_applies_gated_verified() -> None:
    extract = [_item("ПБ 36-12-8п")]
    verified = [_item("ПБ 63-12-8п")]
    corrections = [{"action": "changed_mark", "row_index": 1}]
    result = select_ocr_items(
        extract,
        {"plates": verified, "corrections": corrections},
        _gate,
    )

    assert result.items == _gate(verified)
    assert result.items[0]["gated"] is True
    assert result.verify_failed is False
    assert result.select_reason == "applied"


def test_empty_verified_plates_keeps_extract_and_fails() -> None:
    extract = [_item("ПБ 60-12-8п")]
    result = select_ocr_items(
        extract,
        {"plates": [], "corrections": [{"action": "removed"}]},
        _gate,
    )

    assert result.items == extract
    assert result.verify_failed is True
    assert result.select_reason == "empty_verified_plates"


def test_missing_plates_keeps_extract_and_fails() -> None:
    extract = [_item("ПБ 60-12-8п")]
    result = select_ocr_items(extract, {"corrections": []}, _gate)

    assert result.items == extract
    assert result.verify_failed is True
    assert result.select_reason == "empty_verified_plates"


def test_gate_is_not_called_when_corrections_empty() -> None:
    calls: list[list[dict]] = []

    def recording_gate(items: list[dict]) -> list[dict]:
        calls.append(items)
        return items

    select_ocr_items(
        [_item("ПБ 36-12-8п")],
        {"plates": [_item("ПБ 63-12-8п")], "corrections": []},
        recording_gate,
    )
    assert calls == []


def test_gate_is_not_called_when_verified_empty() -> None:
    calls: list[list[dict]] = []

    def recording_gate(items: list[dict]) -> list[dict]:
        calls.append(items)
        return items

    select_ocr_items(
        [_item("ПБ 60-12-8п")],
        {"plates": [], "corrections": [{"action": "added"}]},
        recording_gate,
    )
    assert calls == []


def _minimal_payload_kwargs() -> dict:
    plate = _item("ПБ 60-12-8п")
    return {
        "plates": [plate],
        "draft_plates": [plate],
        "corrections": [],
        "row_count_on_image": 1,
        "method": "GPT-4o",
        "verify_applied": False,
        "verify_failed": False,
        "cost_usd": 0.0,
    }


def test_build_result_payload_defaults_new_fields_to_none() -> None:
    payload = build_result_payload(**_minimal_payload_kwargs())
    assert payload["ocr_verify_select_reason"] is None
    assert payload["ocr_preprocess"] is None


def test_build_result_payload_passes_select_and_preprocess() -> None:
    payload = build_result_payload(
        **_minimal_payload_kwargs(),
        verify_select_reason="kept_extract_empty_corrections",
        ocr_preprocess="2x_lanczos",
    )
    assert payload["ocr_verify_select_reason"] == "kept_extract_empty_corrections"
    assert payload["ocr_preprocess"] == "2x_lanczos"
