#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Integration tests for pile OCR pipeline (mock provider)."""

from __future__ import annotations

import asyncio
import base64
import json
from io import BytesIO
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from PIL import Image

from core.config.settings import Settings
from core.ocr.pipeline import run_ocr_pipeline, run_pile_ocr_pipeline


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


def _mock_provider(*, extract_rows, verify_rows=None, corrections=None, extract_attr="extract_piles"):
    provider = MagicMock()
    setattr(provider, extract_attr, AsyncMock(return_value=(extract_rows, 0.5)))
    if verify_rows is not None:
        provider.verify_plates = AsyncMock(
            return_value=(
                {
                    "row_count_on_image": len(verify_rows),
                    "plates": verify_rows,
                    "corrections": list(corrections or []),
                },
                0.3,
            )
        )
    return provider


def _other_pile_rows():
    return [
        {
            "raw_name": "С70.30-8",
            "normalized_candidate": "С70.30-8",
            "qty": 4,
            "concrete_grade": "B25",
            "confidence": 0.95,
            "issues": [],
        }
    ]


async def _run_pipeline(
    tmp_path: Path,
    provider,
    settings: Settings,
    *,
    width: int = 1600,
    height: int = 1200,
):
    image_path = tmp_path / "pilot.png"
    Image.new("RGB", (width, height), color=(255, 255, 255)).save(image_path)
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


def test_pile_pipeline_runs_verify_on_small_clean_image(tmp_path: Path):
    settings = _make_settings()
    provider = _mock_provider(extract_rows=_good_pile_rows(), verify_rows=_good_pile_rows())

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings, width=400, height=300))

    assert result is not None
    assert result["ocr_api_calls"] == 2
    assert result["verify_applied"] is True
    assert result["ocr_verify_applied_reason"] == "auto_small_image"
    assert result["ocr_verify_select_reason"] == "kept_extract_empty_corrections"
    assert result["ocr_preprocess"] == "2x_lanczos"
    provider.verify_plates.assert_awaited_once()


def test_pile_pipeline_keeps_extract_when_verify_plates_differ_without_corrections(tmp_path: Path):
    settings = _make_settings()
    extract = _good_pile_rows()
    provider = _mock_provider(extract_rows=extract, verify_rows=_other_pile_rows(), corrections=[])

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings, width=400, height=300))

    assert result is not None
    assert result["ocr_verify_select_reason"] == "kept_extract_empty_corrections"
    assert result["verify_failed"] is False
    assert "С90.30-11 189" in result["text"]
    assert "С70.30-8" not in result["text"]


def test_pile_pipeline_applies_verify_when_corrections_present(tmp_path: Path):
    settings = _make_settings()
    provider = _mock_provider(
        extract_rows=_good_pile_rows(),
        verify_rows=_other_pile_rows(),
        corrections=[{"action": "changed_mark", "row_index": 1}],
    )

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings, width=400, height=300))

    assert result is not None
    assert result["ocr_verify_select_reason"] == "applied"
    assert "С70.30-8 4" in result["text"]
    assert "С90.30-11 189" not in result["text"]


def test_pile_pipeline_small_image_sends_2x_png(tmp_path: Path):
    settings = _make_settings()
    provider = _mock_provider(extract_rows=_good_pile_rows(), verify_rows=_good_pile_rows())

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings, width=400, height=300))

    assert result is not None
    assert result["ocr_preprocess"] == "2x_lanczos"
    kwargs = provider.extract_piles.await_args.kwargs
    assert kwargs["mime_type"] == "image/png"
    with Image.open(BytesIO(base64.b64decode(kwargs["image_base64"]))) as image:
        assert image.size == (800, 600)
    verify_kwargs = provider.verify_plates.await_args.kwargs
    assert verify_kwargs["mime_type"] == "image/png"


def test_pile_pipeline_never_does_not_replace(tmp_path: Path):
    settings = _make_settings(ocr_verify_mode="never")
    provider = _mock_provider(
        extract_rows=_good_pile_rows(),
        verify_rows=_other_pile_rows(),
        corrections=[{"action": "changed_mark"}],
    )
    image_path = tmp_path / "pilot.png"
    Image.new("RGB", (400, 300), color=(255, 255, 255)).save(image_path)

    result = asyncio.run(
        run_pile_ocr_pipeline(
            image_path=str(image_path),
            provider=provider,
            settings=settings,
            verify_enabled=False,
            show_cost=False,
        )
    )

    assert result is not None
    assert result["verify_applied"] is False
    assert result["ocr_verify_select_reason"] is None
    assert "С90.30-11 189" in result["text"]
    provider.verify_plates.assert_not_called()


def test_plates_pipeline_keeps_extract_on_empty_corrections(tmp_path: Path):
    settings = _make_settings()
    extract = [
        {
            "raw_name": "ПБ 36-12-8п",
            "normalized_candidate": "ПБ 36-12-8п",
            "qty": 1,
            "confidence": 0.95,
            "issues": [],
        }
    ]
    verified = [
        {
            "raw_name": "ПБ 63-12-8п",
            "normalized_candidate": "ПБ 63-12-8п",
            "qty": 1,
            "confidence": 0.95,
            "issues": [],
        }
    ]
    provider = _mock_provider(
        extract_rows=extract,
        verify_rows=verified,
        corrections=[],
        extract_attr="extract_plates",
    )
    image_path = tmp_path / "plates.png"
    Image.new("RGB", (400, 300), color=(255, 255, 255)).save(image_path)

    result = asyncio.run(
        run_ocr_pipeline(
            image_path=str(image_path),
            provider=provider,
            settings=settings,
            show_cost=False,
        )
    )

    assert result is not None
    assert result["ocr_verify_select_reason"] == "kept_extract_empty_corrections"
    assert result["plates"][0]["normalized_candidate"] == "ПБ 36-12-8п"
    assert result["ocr_preprocess"] == "2x_lanczos"


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
    provider = _mock_provider(
        extract_rows=bad_rows,
        verify_rows=fixed_rows,
        corrections=[{"action": "changed_mark", "row_index": 1}],
    )

    result = asyncio.run(_run_pipeline(tmp_path, provider, settings))

    assert result is not None
    assert result["ocr_api_calls"] == 2
    assert result["verify_applied"] is True
    assert result["ocr_verify_applied_reason"] in {"auto_unparsed_pile", "auto_low_confidence"}
    assert result["ocr_verify_select_reason"] == "applied"
    assert "С90.30-11 189" in result["text"]
    provider.verify_plates.assert_awaited_once()


def test_pile_pipeline_loads_fixture_json():
    fixture_path = Path(__file__).parent / "fixtures" / "pile_ocr" / "gigachat_extract_response.json"
    payload = json.loads(fixture_path.read_text(encoding="utf-8"))
    assert len(payload) == 3
    assert payload[0]["normalized_candidate"].startswith("С90")
