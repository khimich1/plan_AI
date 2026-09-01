#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GigaChat Vision OCR provider (sync SDK via asyncio.to_thread)."""

from __future__ import annotations

import asyncio
import base64
import json
import logging
from typing import Any, Callable, Dict, List, Optional, TYPE_CHECKING

from core.config.settings import Settings, get_settings
from core.ocr.parsing import parse_gpt_response, parse_verify_response
from core.ocr.prompts import get_verification_prompt
from core.bridge_pile_format_prompt import build_bridge_pile_parser_system_prompt
from core.fbs_format_prompt import build_fbs_parser_system_prompt
from core.march_format_prompt import build_march_parser_system_prompt
from core.pile_format_prompt import build_pile_parser_system_prompt
from core.plate_format_prompt import build_plate_parser_system_prompt
from core.step_format_prompt import build_step_parser_system_prompt

if TYPE_CHECKING:
    from gigachat import GigaChat

try:
    from gigachat import GigaChat as _GigaChat
    from gigachat.models.chat import Chat, Messages, MessagesRole

    GIGACHAT_AVAILABLE = True
except ImportError:
    GIGACHAT_AVAILABLE = False
    _GigaChat = None  # type: ignore[misc, assignment]
    Chat = Messages = MessagesRole = None  # type: ignore[misc, assignment]

_logger = logging.getLogger(__name__)

GIGACHAT_COST_PER_1K_TOKENS_RUB = 0.65

_MIME_TO_EXT = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "application/pdf": ".pdf",
}


def _estimate_cost_rub(tokens_used: int) -> float:
    return (tokens_used / 1000) * GIGACHAT_COST_PER_1K_TOKENS_RUB


def _upload_filename(mime_type: str) -> str:
    return f"ocr-upload{_MIME_TO_EXT.get(mime_type, '.jpg')}"


def require_gigachat_client(settings: Settings | None = None) -> "GigaChat":
    if not GIGACHAT_AVAILABLE or _GigaChat is None:
        raise RuntimeError("GigaChat недоступен. Установите: pip install gigachat")

    cfg = settings or get_settings()
    credentials = (cfg.gigachat_credentials or "").strip()
    if not credentials:
        raise ValueError(
            "Для распознавания через GigaChat задайте GIGACHAT_CREDENTIALS в окружении backend."
        )

    return _GigaChat(
        credentials=credentials,
        scope=cfg.gigachat_scope,
        model=cfg.gigachat_model,
        verify_ssl_certs=False,
        timeout=600,
    )


def _sync_extract(
    *,
    client: "GigaChat",
    system_prompt: str,
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    """Extract items via GigaChat. Image is optional (text-only chat when omitted)."""
    attachments: list[str] | None = None
    if image_base64 and mime_type:
        image_bytes = base64.b64decode(image_base64)
        uploaded = client.upload_file((_upload_filename(mime_type), image_bytes))
        attachments = [uploaded.id_]

    user_kwargs: dict[str, Any] = {
        "role": MessagesRole.USER,
        "content": user_text,
    }
    if attachments:
        user_kwargs["attachments"] = attachments

    response = client.chat(
        Chat(
            messages=[
                Messages(
                    role=MessagesRole.SYSTEM,
                    content=system_prompt,
                ),
                Messages(**user_kwargs),
            ],
            temperature=0,
            max_tokens=max_tokens,
        )
    )

    result_text = response.choices[0].message.content or ""
    items = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    return items, _estimate_cost_rub(tokens_used)


def _sync_extract_plates(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_plate_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_extract_piles(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_pile_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_extract_steps(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_step_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_extract_marches(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_march_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_extract_bridge_piles(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_bridge_pile_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_extract_fbs(
    *,
    client: "GigaChat",
    user_text: str,
    image_base64: str | None,
    mime_type: str | None,
    max_tokens: int,
) -> tuple[List[Dict[str, Any]], float]:
    return _sync_extract(
        client=client,
        system_prompt=build_fbs_parser_system_prompt(),
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
        max_tokens=max_tokens,
    )


def _sync_verify_plates(
    *,
    client: "GigaChat",
    image_base64: str,
    mime_type: str,
    draft_plates: List[Dict[str, Any]],
) -> tuple[Dict[str, Any], float]:
    draft_json = json.dumps(draft_plates, ensure_ascii=False, indent=2)
    image_bytes = base64.b64decode(image_base64)
    uploaded = client.upload_file((_upload_filename(mime_type), image_bytes))

    response = client.chat(
        Chat(
            messages=[
                Messages(
                    role=MessagesRole.USER,
                    content=get_verification_prompt(draft_json),
                    attachments=[uploaded.id_],
                ),
            ],
            temperature=0,
            max_tokens=2500,
        )
    )

    result_text = response.choices[0].message.content or ""
    verify_result = parse_verify_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    _logger.info(
        "[OCR] GigaChat Verify: tokens=%s rows=%s corrections=%s",
        tokens_used,
        len(verify_result.get("plates") or []),
        len(verify_result.get("corrections") or []),
    )
    return verify_result, _estimate_cost_rub(tokens_used)


class GigaChatProvider:
    """GigaChat Vision implementation of OcrProvider."""

    def __init__(
        self,
        *,
        settings: Settings | None = None,
        client: Optional["GigaChat"] = None,
    ) -> None:
        self._settings = settings
        self._client = client

    def _get_client(self) -> "GigaChat":
        if self._client is None:
            self._client = require_gigachat_client(self._settings)
        return self._client

    async def _extract(
        self,
        sync_fn: Callable[..., tuple[List[Dict[str, Any]], float]],
        *,
        user_text: str,
        image_base64: str | None,
        mime_type: str | None,
        max_tokens: int,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await asyncio.to_thread(
            sync_fn,
            client=self._get_client(),
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_plates(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_plates,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_piles(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_piles,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_steps(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_steps,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_marches(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_marches,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_bridge_piles(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_bridge_piles,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def extract_fbs(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await self._extract(
            _sync_extract_fbs,
            user_text=user_text,
            image_base64=image_base64,
            mime_type=mime_type,
            max_tokens=max_tokens,
        )

    async def verify_plates(
        self,
        *,
        image_base64: str,
        mime_type: str,
        draft_plates: List[Dict[str, Any]],
    ) -> tuple[Dict[str, Any], float]:
        return await asyncio.to_thread(
            _sync_verify_plates,
            client=self._get_client(),
            image_base64=image_base64,
            mime_type=mime_type,
            draft_plates=draft_plates,
        )
