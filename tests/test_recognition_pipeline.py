import json
import asyncio
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from PIL import Image

import core.config_and_data as cfg
from core.config.settings import get_settings
from core.ocr_gpt import (
    _validate_plate_item,
    format_corrections_for_user,
    parse_gpt_response,
    parse_verify_response,
    recognize_text_smart,
)


def test_normalize_prefix_with_dot_and_comma():
    from core.plate_text_normalizer import normalize_plate_prefixes

    assert normalize_plate_prefixes("ПБ.19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПБ,19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПБ19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПВ.19,6-12-10 7") == "ПБ 19,6-12-10 7"


def test_parse_line_tolerant_pbpk():
    from core.plate_line_parser import parse_line

    result = parse_line("ПБ.19,6-12-10 7")
    assert result.parsed is True
    assert result.stage == "tolerant_pbpk"
    assert result.length_m == 1.96
    assert result.width_m == 1.2
    assert result.qty == 7
    assert result.load_code == 10.0


def test_parse_gpt_response_new_contract():
    response = """
    [
      {
        "raw_name": "ПБ.19,6-12-10",
        "normalized_candidate": "ПБ 19,6-12-10",
        "qty": "7",
        "confidence": 0.8,
        "issues": ["prefix_separator_dot"]
      }
    ]
    """
    parsed = parse_gpt_response(response)
    assert len(parsed) == 1
    assert parsed[0]["raw_name"] == "ПБ.19,6-12-10"
    assert parsed[0]["normalized_candidate"] == "ПБ 19,6-12-10"
    assert parsed[0]["qty"] == 7
    assert parsed[0]["confidence"] == 0.8
    assert parsed[0]["issues"] == ["prefix_separator_dot"]


def test_validate_plate_item_shared():
    valid = _validate_plate_item(
        {
            "raw_name": "ПБ 66,2-12-8п",
            "normalized_candidate": "ПБ 66,2-12-8п",
            "qty": 3,
            "confidence": 0.91,
            "issues": [],
        }
    )
    assert valid is not None
    assert valid["qty"] == 3

    assert _validate_plate_item({"qty": 1}) is None
    assert _validate_plate_item({"raw_name": "ПБ 60-12-8п", "qty": "x"}) is None
    assert _validate_plate_item("not-a-dict") is None


def test_parse_verify_response_valid():
    response = json.dumps(
        {
            "row_count_on_image": 2,
            "plates": [
                {
                    "raw_name": "ПБ 90,9-10,2-6п",
                    "normalized_candidate": "ПБ 90,9-10,2-6п",
                    "qty": 1,
                    "confidence": 0.99,
                    "issues": [],
                },
                {
                    "raw_name": "ПБ 66,2-2,6-8п",
                    "normalized_candidate": "ПБ 66,2-2,6-8п",
                    "qty": 1,
                    "confidence": 0.97,
                    "issues": [],
                },
            ],
            "corrections": [
                {
                    "action": "added",
                    "row_index": 5,
                    "before": None,
                    "after": {"normalized_candidate": "ПБ 90,9-10,2-6п", "qty": 1},
                    "reason": "пропущена строка",
                }
            ],
        },
        ensure_ascii=False,
    )
    parsed = parse_verify_response(response)
    assert parsed["row_count_on_image"] == 2
    assert len(parsed["plates"]) == 2
    assert parsed["plates"][0]["normalized_candidate"] == "ПБ 90,9-10,2-6п"
    assert len(parsed["corrections"]) == 1
    assert parsed["corrections"][0]["action"] == "added"


def test_parse_verify_response_fallback_array():
    response = """
    [
      {
        "raw_name": "ПБ 56-12-8п",
        "normalized_candidate": "ПБ 56-12-8п",
        "qty": 7,
        "confidence": 0.95,
        "issues": []
      }
    ]
    """
    parsed = parse_verify_response(response)
    assert len(parsed["plates"]) == 1
    assert parsed["plates"][0]["qty"] == 7
    assert parsed["corrections"] == []
    assert parsed["row_count_on_image"] == 1


def test_parse_verify_response_invalid_json():
    parsed = parse_verify_response("not json at all")
    assert parsed["plates"] == []
    assert parsed["corrections"] == []
    assert parsed["row_count_on_image"] is None


def test_format_corrections_for_user():
    corrections = [
        {
            "action": "added",
            "row_index": 5,
            "after": {"normalized_candidate": "ПБ 90,9-10,2-6п", "qty": 1},
        },
        {
            "action": "changed_mark",
            "row_index": 11,
            "before": {"normalized_candidate": "ПБ 66,2-9,2-6п", "qty": 1},
            "after": {"normalized_candidate": "ПБ 66,2-2,6-8п", "qty": 1},
        },
        {
            "action": "changed_qty",
            "row_index": 22,
            "before": {"normalized_candidate": "ПБ 56-2,6-8п", "qty": 2},
            "after": {"normalized_candidate": "ПБ 56-2,6-8п", "qty": 1},
        },
    ]
    text = format_corrections_for_user(corrections, max_items=2)
    assert "Автоисправлено 3 строк" in text
    assert "добавлено" in text
    assert "ещё 1" in text


def test_format_corrections_for_user_verify_failed_only():
    text = format_corrections_for_user([{"action": "verify_failed"}])
    assert text == ""


def test_set_plate_lists_diagnostics_for_unparsed():
    unparsed, _, _ = cfg.set_plate_lists_from_text("непонятная строка")
    assert len(unparsed) == 1
    diagnostics = cfg.get_last_parse_diagnostics()
    assert diagnostics
    assert diagnostics[0]["validation_status"] == "failed"
    assert diagnostics[0]["reason_code"] in {"pattern_not_matched", "empty_line"}


async def _run_recognize_text_smart_single_call(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("OCR_PROVIDER", "openai")
    get_settings.cache_clear()

    image_path = tmp_path / "table.png"
    Image.new("RGB", (1600, 1200), color=(255, 255, 255)).save(image_path)

    draft_response = MagicMock()
    draft_response.choices = [MagicMock(message=MagicMock(content=json.dumps([
        {
            "raw_name": "ПБ 90,9-12-6п",
            "normalized_candidate": "ПБ 90,9-12-6п",
            "qty": 3,
            "confidence": 0.95,
            "issues": [],
        }
    ])))]
    draft_response.usage = MagicMock(total_tokens=1000)

    mock_create = AsyncMock(return_value=draft_response)
    mock_client = MagicMock()
    mock_client.chat.completions.create = mock_create

    with patch("core.ocr_gpt.AsyncOpenAI", return_value=mock_client):
        result = await recognize_text_smart(str(image_path), show_cost=False)

    assert result is not None
    assert result["verify_applied"] is False
    assert result["method"] == "GPT-4o"
    assert len(result["plates"]) == 1
    assert mock_create.await_count == 1
    call_kwargs = mock_create.await_args.kwargs
    assert call_kwargs["temperature"] == 0.0
    messages = call_kwargs["messages"]
    assert messages[0]["role"] == "system"
    assert messages[1]["role"] == "user"


def test_recognize_text_smart_single_call(tmp_path, monkeypatch):
    asyncio.run(_run_recognize_text_smart_single_call(tmp_path, monkeypatch))


async def _run_recognize_text_smart_small_image_runs_verify(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("OCR_PROVIDER", "openai")
    get_settings.cache_clear()

    image_path = tmp_path / "table.png"
    Image.new("RGB", (400, 300), color=(255, 255, 255)).save(image_path)

    plate = {
        "raw_name": "ПБ 90,9-12-6п",
        "normalized_candidate": "ПБ 90,9-12-6п",
        "qty": 3,
        "confidence": 0.95,
        "issues": [],
    }
    extract_response = MagicMock()
    extract_response.choices = [MagicMock(message=MagicMock(content=json.dumps([plate])))]
    extract_response.usage = MagicMock(total_tokens=1000)

    verify_response = MagicMock()
    verify_response.choices = [
        MagicMock(
            message=MagicMock(
                content=json.dumps(
                    {
                        "row_count_on_image": 1,
                        "plates": [plate],
                        "corrections": [],
                    }
                )
            )
        )
    ]
    verify_response.usage = MagicMock(total_tokens=800)

    mock_create = AsyncMock(side_effect=[extract_response, verify_response])
    mock_client = MagicMock()
    mock_client.chat.completions.create = mock_create

    with patch("core.ocr_gpt.AsyncOpenAI", return_value=mock_client):
        result = await recognize_text_smart(str(image_path), show_cost=False)

    assert result is not None
    assert result["verify_applied"] is True
    assert result["ocr_verify_applied_reason"] == "auto_small_image"
    assert mock_create.await_count == 2


def test_recognize_text_smart_small_image_runs_verify(tmp_path, monkeypatch):
    asyncio.run(_run_recognize_text_smart_small_image_runs_verify(tmp_path, monkeypatch))


async def _run_recognize_text_smart_verify_disabled(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("OCR_PROVIDER", "openai")
    monkeypatch.setenv("OCR_VERIFY_MODE", "never")
    monkeypatch.setenv("OCR_VERIFY_ENABLED", "false")
    get_settings.cache_clear()

    image_path = tmp_path / "table.png"
    Image.new("RGB", (1600, 1200), color=(255, 255, 255)).save(image_path)

    draft_response = MagicMock()
    draft_response.choices = [MagicMock(message=MagicMock(content=json.dumps([
        {
            "raw_name": "ПБ 60-12-8п",
            "normalized_candidate": "ПБ 60-12-8п",
            "qty": 7,
            "confidence": 0.95,
            "issues": [],
        }
    ])))]
    draft_response.usage = MagicMock(total_tokens=800)

    mock_create = AsyncMock(return_value=draft_response)
    mock_client = MagicMock()
    mock_client.chat.completions.create = mock_create

    with patch("core.ocr_gpt.AsyncOpenAI", return_value=mock_client):
        result = await recognize_text_smart(str(image_path), show_cost=False)

    assert result is not None
    assert result["verify_applied"] is False
    assert result["method"] == "GPT-4o"
    assert mock_create.await_count == 1


def test_recognize_text_smart_verify_disabled(tmp_path, monkeypatch):
    asyncio.run(_run_recognize_text_smart_verify_disabled(tmp_path, monkeypatch))


async def _run_apply_plates_with_ai(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("OCR_PROVIDER", "openai")
    get_settings.cache_clear()

    ai_response = MagicMock()
    ai_response.choices = [MagicMock(message=MagicMock(content=json.dumps([
        {
            "raw_name": "ПБ 60-12-8п",
            "normalized_candidate": "ПБ 60-12-8п",
            "qty": 7,
            "confidence": 0.99,
            "issues": [],
        }
    ])))]
    ai_response.usage = MagicMock(total_tokens=900)

    mock_create = AsyncMock(return_value=ai_response)
    mock_client = MagicMock()
    mock_client.chat.completions.create = mock_create

    from core.ocr_gpt import apply_plates_with_ai

    with patch("core.ocr_gpt.AsyncOpenAI", return_value=mock_client):
        result = await apply_plates_with_ai(
            current_plates_text="ПБ 78-12-8п 2",
            user_instruction="замени 78 на 60 и qty 7",
        )

    assert result is not None
    assert result["method"] == "GPT-4o+ai"
    assert "ПБ 60-12-8п 7" in result["text"]
    assert mock_create.await_count == 1


def test_apply_plates_with_ai(tmp_path, monkeypatch):
    asyncio.run(_run_apply_plates_with_ai(tmp_path, monkeypatch))


async def _run_apply_plates_with_ai_gigachat_text_only(monkeypatch):
    monkeypatch.setenv("OCR_PROVIDER", "gigachat")
    monkeypatch.setenv("GIGACHAT_CREDENTIALS", "dGVzdA==")
    monkeypatch.setenv("GIGACHAT_MODEL", "GigaChat-2-Max")
    get_settings.cache_clear()

    plates_json = json.dumps(
        [
            {
                "raw_name": "ПБ 70-12-8п",
                "normalized_candidate": "ПБ 70-12-8п",
                "qty": 5,
                "confidence": 0.99,
                "issues": [],
            },
            {
                "raw_name": "ПБ 90-12-8п",
                "normalized_candidate": "ПБ 90-12-8п",
                "qty": 5,
                "confidence": 0.99,
                "issues": [],
            },
        ],
        ensure_ascii=False,
    )
    mock_client = MagicMock()
    mock_client.chat.return_value = MagicMock(
        choices=[MagicMock(message=MagicMock(content=plates_json))],
        usage=MagicMock(total_tokens=1500),
    )

    from core.ocr_gpt import apply_plates_with_ai

    with patch(
        "core.ocr.providers.gigachat.require_gigachat_client",
        return_value=mock_client,
    ):
        result = await apply_plates_with_ai(
            current_plates_text="ПБ 70.12-8К7 5\nПБ 90.12-8К7 5",
            user_instruction="ПБ 70.12-8К7 5 это ПБ 70-12-8п 5 исправь по аналогии",
        )

    assert result is not None
    assert "GigaChat" in result["method"]
    assert result["method"].endswith("+ai")
    assert "ПБ 70-12-8п 5" in result["text"]
    assert "ПБ 90-12-8п 5" in result["text"]
    assert result["ocr_cost_rub"] > 0
    assert result["cost_usd"] == 0.0
    mock_client.upload_file.assert_not_called()
    mock_client.chat.assert_called_once()


def test_apply_plates_with_ai_gigachat_text_only(monkeypatch):
    asyncio.run(_run_apply_plates_with_ai_gigachat_text_only(monkeypatch))


async def _run_apply_plates_with_ai_gigachat_with_image(tmp_path, monkeypatch):
    monkeypatch.setenv("OCR_PROVIDER", "gigachat")
    monkeypatch.setenv("GIGACHAT_CREDENTIALS", "dGVzdA==")
    monkeypatch.setenv("GIGACHAT_MODEL", "GigaChat-2-Max")
    get_settings.cache_clear()

    image_path = tmp_path / "plates.png"
    Image.new("RGB", (32, 32), color=(200, 200, 200)).save(image_path)

    plates_json = json.dumps(
        [
            {
                "raw_name": "ПБ 60-12-8п",
                "normalized_candidate": "ПБ 60-12-8п",
                "qty": 7,
                "confidence": 0.99,
                "issues": [],
            }
        ],
        ensure_ascii=False,
    )
    uploaded = MagicMock(id_="file-ai")
    mock_client = MagicMock()
    mock_client.upload_file.return_value = uploaded
    mock_client.chat.return_value = MagicMock(
        choices=[MagicMock(message=MagicMock(content=plates_json))],
        usage=MagicMock(total_tokens=2000),
    )

    from core.ocr_gpt import apply_plates_with_ai

    with patch(
        "core.ocr.providers.gigachat.require_gigachat_client",
        return_value=mock_client,
    ):
        result = await apply_plates_with_ai(
            current_plates_text="ПБ 78-12-8п 2",
            user_instruction="замени 78 на 60 и qty 7",
            image_path=str(image_path),
        )

    assert result is not None
    assert "GigaChat-2-Max+ai" == result["method"]
    assert "ПБ 60-12-8п 7" in result["text"]
    mock_client.upload_file.assert_called_once()
    mock_client.chat.assert_called_once()


def test_apply_plates_with_ai_gigachat_with_image(tmp_path, monkeypatch):
    asyncio.run(_run_apply_plates_with_ai_gigachat_with_image(tmp_path, monkeypatch))
