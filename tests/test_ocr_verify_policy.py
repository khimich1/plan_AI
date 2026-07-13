#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pytest

from core.ocr.verify_policy import OcrVerifySettings, should_run_verify

DEFAULT_SETTINGS = OcrVerifySettings(
    max_rows=10,
    min_confidence=0.92,
    max_bytes=819_200,
)


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


def _good_plates(count: int = 3):
    return [_plate() for _ in range(count)]


# --- never / max_api_calls ---


@pytest.mark.parametrize("mode", ["never", "auto", "always"])
def test_never_mode_or_max_one_call_always_skips(mode):
    run, reason = should_run_verify(
        mode=mode,
        max_api_calls=1,
        image_size_bytes=1000,
        plates=_good_plates(),
        settings=DEFAULT_SETTINGS,
    )
    assert run is False
    assert reason == "max_api_calls_or_never"


def test_never_mode_skips_even_with_suspicious_plates():
    run, reason = should_run_verify(
        mode="never",
        max_api_calls=2,
        image_size_bytes=10_000_000,
        plates=[_plate(normalized_candidate="???", confidence=0.1, issues=["x"])],
        settings=DEFAULT_SETTINGS,
    )
    assert run is False
    assert reason == "max_api_calls_or_never"


# --- always ---


def test_always_mode_runs_when_max_api_calls_allows():
    run, reason = should_run_verify(
        mode="always",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=_good_plates(),
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "mode_always"


# --- auto: all checks pass → skip ---


def test_auto_all_checks_passed_skips_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=DEFAULT_SETTINGS.max_bytes,
        plates=_good_plates(10),
        settings=DEFAULT_SETTINGS,
    )
    assert run is False
    assert reason == "auto_all_checks_passed"


def test_auto_at_exact_thresholds_skips():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=DEFAULT_SETTINGS.max_bytes,
        plates=_good_plates(DEFAULT_SETTINGS.max_rows),
        settings=DEFAULT_SETTINGS,
    )
    assert run is False
    assert reason == "auto_all_checks_passed"


# --- auto: individual violations → run ---


def test_auto_empty_plates_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=[],
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_empty_plates"


def test_auto_file_too_large_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=DEFAULT_SETTINGS.max_bytes + 1,
        plates=_good_plates(),
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_file_too_large"


def test_auto_too_many_rows_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=_good_plates(DEFAULT_SETTINGS.max_rows + 1),
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_too_many_rows"


def test_auto_low_confidence_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=[_plate(confidence=DEFAULT_SETTINGS.min_confidence - 0.01)],
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_low_confidence"


def test_auto_min_confidence_exact_passes():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=[_plate(confidence=DEFAULT_SETTINGS.min_confidence)],
        settings=DEFAULT_SETTINGS,
    )
    assert run is False
    assert reason == "auto_all_checks_passed"


def test_auto_unparsed_plate_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=[_plate(normalized_candidate="Непонятный текст")],
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_unparsed_plate"


def test_auto_has_issues_runs_verify():
    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=[_plate(issues=["prefix_separator_dot"])],
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "auto_has_issues"


def test_auto_custom_settings_thresholds():
    tight = OcrVerifySettings(max_rows=2, min_confidence=0.99, max_bytes=5000)

    run, reason = should_run_verify(
        mode="auto",
        max_api_calls=2,
        image_size_bytes=4000,
        plates=_good_plates(3),
        settings=tight,
    )
    assert run is True
    assert reason == "auto_too_many_rows"


def test_unknown_mode_runs_verify():
    run, reason = should_run_verify(
        mode="bogus",
        max_api_calls=2,
        image_size_bytes=1000,
        plates=_good_plates(),
        settings=DEFAULT_SETTINGS,
    )
    assert run is True
    assert reason == "unknown_mode_bogus"
