#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import base64
import json
from unittest.mock import MagicMock, patch

import pytest

from core.config.settings import Settings
from core.ocr.providers.gigachat import (
    GIGACHAT_COST_PER_1K_TOKENS_RUB,
    GigaChatProvider,
    _estimate_cost_rub,
    require_gigachat_client,
)


def _make_settings(**overrides) -> Settings:
    base = {
        "app_secret_key": "x" * 32,
        "gigachat_credentials": "dGVzdC1jcmVkZW50aWFscw==",
        "gigachat_model": "GigaChat-2-Max",
        "gigachat_scope": "GIGACHAT_API_PERS",
    }
    base.update(overrides)
    return Settings(**base)


def test_estimate_cost_rub():
    assert _estimate_cost_rub(1000) == pytest.approx(GIGACHAT_COST_PER_1K_TOKENS_RUB)
    assert _estimate_cost_rub(2000) == pytest.approx(GIGACHAT_COST_PER_1K_TOKENS_RUB * 2)


def test_require_gigachat_client_missing_credentials():
    settings = _make_settings(gigachat_credentials="")
    with pytest.raises(ValueError, match="GIGACHAT_CREDENTIALS"):
        require_gigachat_client(settings)


def _mock_chat_response(content: str, *, tokens: int = 4000):
    response = MagicMock()
    response.choices = [MagicMock(message=MagicMock(content=content))]
    response.usage = MagicMock(total_tokens=tokens)
    return response


async def _run_extract_plates_mocked():
    plates_json = json.dumps(
        [
            {
                "raw_name": "ПБ 60-12-8п",
                "normalized_candidate": "ПБ 60-12-8п",
                "qty": 2,
                "confidence": 0.95,
                "issues": [],
            }
        ],
        ensure_ascii=False,
    )
    uploaded = MagicMock(id_="file-123")
    mock_client = MagicMock()
    mock_client.upload_file.return_value = uploaded
    mock_client.chat.return_value = _mock_chat_response(plates_json, tokens=5000)

    settings = _make_settings()
    provider = GigaChatProvider(settings=settings, client=mock_client)

    image_b64 = base64.b64encode(b"\x89PNG\r\n\x1a\nfake").decode()
    plates, cost_rub = await provider.extract_plates(
        user_text="Распознай таблицу",
        image_base64=image_b64,
        mime_type="image/png",
    )

    assert len(plates) == 1
    assert plates[0]["qty"] == 2
    assert cost_rub == pytest.approx(_estimate_cost_rub(5000))
    mock_client.upload_file.assert_called_once()
    mock_client.chat.assert_called_once()


def test_gigachat_extract_plates_mocked():
    asyncio.run(_run_extract_plates_mocked())


async def _run_extract_piles_mocked():
    piles_json = json.dumps(
        [
            {
                "raw_name": "С90.30-11",
                "normalized_candidate": "С90.30-11",
                "qty": 189,
                "concrete_grade": "B25",
                "confidence": 0.95,
                "issues": [],
            }
        ],
        ensure_ascii=False,
    )
    uploaded = MagicMock(id_="file-pile")
    mock_client = MagicMock()
    mock_client.upload_file.return_value = uploaded
    mock_client.chat.return_value = _mock_chat_response(piles_json, tokens=5000)

    settings = _make_settings()
    provider = GigaChatProvider(settings=settings, client=mock_client)

    image_b64 = base64.b64encode(b"\x89PNG\r\n\x1a\nfake").decode()
    piles, cost_rub = await provider.extract_piles(
        user_text="Распознай таблицу свай",
        image_base64=image_b64,
        mime_type="image/png",
    )

    assert len(piles) == 1
    assert piles[0]["qty"] == 189
    assert cost_rub == pytest.approx(_estimate_cost_rub(5000))
    mock_client.chat.assert_called_once()
    system_msg = mock_client.chat.call_args[0][0].messages[0].content
    assert "сваи" in system_msg.lower()


def test_gigachat_extract_piles_mocked():
    asyncio.run(_run_extract_piles_mocked())


async def _run_verify_plates_mocked():
    verify_json = json.dumps(
        {
            "row_count_on_image": 1,
            "plates": [
                {
                    "raw_name": "ПБ 60-12-8п",
                    "normalized_candidate": "ПБ 60-12-8п",
                    "qty": 3,
                    "confidence": 0.99,
                    "issues": [],
                }
            ],
            "corrections": [
                {
                    "action": "changed_qty",
                    "row_index": 1,
                    "before": {"normalized_candidate": "ПБ 60-12-8п", "qty": 2},
                    "after": {"normalized_candidate": "ПБ 60-12-8п", "qty": 3},
                    "reason": "qty на фото = 3",
                }
            ],
        },
        ensure_ascii=False,
    )
    uploaded = MagicMock(id_="file-verify")
    mock_client = MagicMock()
    mock_client.upload_file.return_value = uploaded
    mock_client.chat.return_value = _mock_chat_response(verify_json, tokens=6000)

    settings = _make_settings()
    provider = GigaChatProvider(settings=settings, client=mock_client)
    draft = [
        {
            "raw_name": "ПБ 60-12-8п",
            "normalized_candidate": "ПБ 60-12-8п",
            "qty": 2,
            "confidence": 0.95,
            "issues": [],
        }
    ]

    image_b64 = base64.b64encode(b"\x89PNG\r\n\x1a\nfake").decode()
    result, cost_rub = await provider.verify_plates(
        image_base64=image_b64,
        mime_type="image/png",
        draft_plates=draft,
    )

    assert result["plates"][0]["qty"] == 3
    assert len(result["corrections"]) == 1
    assert cost_rub == pytest.approx(_estimate_cost_rub(6000))
    assert mock_client.chat.call_count == 1


def test_gigachat_verify_plates_mocked():
    asyncio.run(_run_verify_plates_mocked())


async def _run_provider_uses_to_thread(monkeypatch):
    plates_json = json.dumps(
        [
            {
                "raw_name": "ПБ 60-12-8п",
                "normalized_candidate": "ПБ 60-12-8п",
                "qty": 1,
                "confidence": 0.95,
                "issues": [],
            }
        ]
    )
    uploaded = MagicMock(id_="file-thread")
    mock_client = MagicMock()
    mock_client.upload_file.return_value = uploaded
    mock_client.chat.return_value = _mock_chat_response(plates_json, tokens=1000)

    to_thread_calls: list = []

    async def fake_to_thread(func, *args, **kwargs):
        to_thread_calls.append(func.__name__)
        return func(*args, **kwargs)

    monkeypatch.setattr(
        "core.ocr.providers.gigachat.asyncio.to_thread",
        fake_to_thread,
    )

    settings = _make_settings()
    provider = GigaChatProvider(settings=settings, client=mock_client)
    image_b64 = base64.b64encode(b"img").decode()

    await provider.extract_plates(
        user_text="test",
        image_base64=image_b64,
        mime_type="image/jpeg",
    )

    assert to_thread_calls == ["_sync_extract_plates"]


def test_gigachat_provider_uses_to_thread(monkeypatch):
    asyncio.run(_run_provider_uses_to_thread(monkeypatch))


def test_require_gigachat_client_creates_client():
    settings = _make_settings()
    with patch("core.ocr.providers.gigachat._GigaChat") as mock_cls:
        mock_cls.return_value = MagicMock()
        client = require_gigachat_client(settings)
        assert client is mock_cls.return_value
        mock_cls.assert_called_once()
        call_kwargs = mock_cls.call_args.kwargs
        assert call_kwargs["credentials"] == settings.gigachat_credentials
        assert call_kwargs["scope"] == "GIGACHAT_API_PERS"
        assert call_kwargs["model"] == "GigaChat-2-Max"
