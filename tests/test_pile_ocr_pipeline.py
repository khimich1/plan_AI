#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Integration tests for pile OCR pipeline (mock provider)."""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest

from core.config.settings import Settings
from core.ocr.pipeline import run_pile_ocr_pipeline


def _make_settings(**overrides) -> Settings:
    base = {
        "app_secret_key": "x" * 32,
        "ocr_provider": "gigachat",
        "gigachat_credentials": "dGVzdC1jcmVkZW50aWFscw==",
        "gigachat_model": "GigaChat-2-Max",
        "ocr_max_api_calls": 2,
        "ocr_verify_mode": "auto",
        "ocr_verify_auto_max_rows": 10,
        "ocr_verify_auto_min_confidence": 0.92,
        "ocr_verify_auto_max_bytes": 819_200,
    }
    base.update(overrides)
    return Settings(**base)


def _good_pile_rows():
    return [
        {
            "raw_name": "С90.30-11",
            "normalized_candidate": "С90.30-11",
            "qty": 189,
            "concrete_grade": "B25",
            "confidence": 0.95,
            "issues": [],
        },
        {
            "raw_name": "С110.30-13",
            "normalized_candidate": "С110.30-13",
            "qty": 26,
            "concrete_grade": "B25",
            "confidence": 0.95,
            "issues": [],
        },
        {
            "raw_name": "С120.30-12",
            "normalized_candidate": "С120.30-12",
            "qty": 20,
            "concrete_grade": "B25",
            "confidence": 0.95,
            "issues": [],
        },
    ]


def _mock_provider(*, extract_rows, verify_rows=None):
    provider = MagicMock()
    provider.extract_piles = AsyncMock(return_value=(extract_rows, 0.5))
    if verify_rows is not None:
        provider.verify_plates = AsyncMock(
            return_value=(
                {
                    "row_count_on_image": len(verify_rows),
                    "plates": verify_rows,
                    "corrections": [],
                },
                0.3,
            )
        )
    return provider


async def _run_pipeline(tmp_path: Path, provider, settings: Settings):
    image_path = tmp_path / "pilot.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 32)
    return await run_pile_ocr_pipeline(
        image_path=str(image_path),
        provider=provider,
        settings=settings,
        show_cost=False,
    )


def test_pile_pipeline_skips_verify_on_clean_rows(tmp_path: Path):
    settings = _make_settings()
    provider = _mock_provider(extract_rows=_good_pile_rows())

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings))

    assert result is not None
    assert result["ocr_api_calls"] == 1
    assert result["verify_applied"] is False
    assert result["ocr_verify_skipped_reason"] == "auto_all_checks_passed"
    assert "С90.30-11 189" in result["text"]
    provider.extract_piles.assert_awaited_once()
    provider.verify_plates.assert_not_called()


def test_pile_pipeline_runs_verify_on_parser_rejected(tmp_path: Path):
    settings = _make_settings()
    bad_rows = [
        {
            "raw_name": "???",
            "normalized_candidate": "???",
            "qty": 1,
            "confidence": 0.99,
            "issues": [],
        }
    ]
    fixed_rows = _good_pile_rows()
    provider = _mock_provider(extract_rows=bad_rows, verify_rows=fixed_rows)

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings))

    assert result is not None
    assert result["ocr_api_calls"] == 2
    assert result["verify_applied"] is True
    assert result["ocr_verify_applied_reason"] in {"auto_unparsed_pile", "auto_low_confidence"}
    provider.verify_plates.assert_awaited_once()


def test_pile_pipeline_loads_fixture_json():
    fixture_path = Path(__file__).parent / "fixtures" / "pile_ocr" / "gigachat_extract_response.json"
    payload = json.loads(fixture_path.read_text(encoding="utf-8"))
    assert len(payload) == 3
    assert payload[0]["normalized_candidate"].startswith("С90")
