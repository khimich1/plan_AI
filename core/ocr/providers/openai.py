#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OpenAI GPT-4o Vision OCR provider."""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Type

from core.ocr.parsing import parse_gpt_response, parse_verify_response
from core.ocr.prompts import OCR_USER_PROMPT, get_verification_prompt
from core.plate_format_prompt import build_plate_parser_system_prompt

try:
    from openai import AsyncOpenAI as _AsyncOpenAI
    GPT_AVAILABLE = True
except ImportError:
    GPT_AVAILABLE = False
    _AsyncOpenAI = None  # type: ignore[misc, assignment]

# Re-exported by core.ocr_gpt shim; tests patch core.ocr_gpt.AsyncOpenAI.
AsyncOpenAI: Optional[Type[Any]] = _AsyncOpenAI

if not GPT_AVAILABLE:
    print("[GPT OCR] ⚠️ OpenAI не установлен. Установите: pip install openai")


def _async_openai_cls() -> Type[Any]:
    """Late binding so unittest.mock.patch('core.ocr_gpt.AsyncOpenAI') works."""
    import core.ocr_gpt as shim

    return shim.AsyncOpenAI


def _estimate_cost_usd(tokens_used: int) -> float:
    # GPT-4o: $2.50 за 1M токенов (упрощённая оценка по total_tokens)
    return (tokens_used / 1_000_000) * 2.5


def load_image_payload(image_path: str) -> Tuple[bytes, str, str]:
    """Читает файл и возвращает (bytes, base64, mime_type)."""
    with open(image_path, "rb") as f:
        image_data = f.read()
    image_base64 = base64.b64encode(image_data).decode()
    mime_type = image_mime_type(image_path, image_data)
    return image_data, image_base64, mime_type


def image_mime_type(image_path: str, image_data: bytes) -> str:
    """MIME для data: URL Vision API (PNG раньше ошибочно слали как image/jpeg)."""
    if len(image_data) >= 8 and image_data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if len(image_data) >= 3 and image_data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if len(image_data) >= 12 and image_data.startswith(b"RIFF") and image_data[8:12] == b"WEBP":
        return "image/webp"
    if len(image_data) >= 6 and image_data.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if len(image_data) >= 4 and image_data.startswith(b"%PDF"):
        return "application/pdf"
    suffix = Path(image_path).suffix.lower()
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif",
        ".pdf": "application/pdf",
    }.get(suffix, "image/jpeg")


def require_openai_client() -> Any:
    if not GPT_AVAILABLE:
        raise RuntimeError("GPT недоступен. Установите: pip install openai")
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "Для распознавания по фото задайте OPENAI_API_KEY в окружении backend "
            "(docker-compose: сервис backend; локально: .env или экспорт переменной)."
        )
    client_cls = _async_openai_cls()
    if client_cls is None:
        raise RuntimeError("GPT недоступен. Установите: pip install openai")
    return client_cls(api_key=api_key)


async def call_gpt_for_plates(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_plate_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    plates = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return plates, cost_usd


async def call_gpt_for_piles(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    from core.pile_format_prompt import build_pile_parser_system_prompt

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_pile_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    piles = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return piles, cost_usd


async def call_gpt_for_steps(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    from core.step_format_prompt import build_step_parser_system_prompt

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_step_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    steps = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return steps, cost_usd


async def call_gpt_for_marches(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    from core.march_format_prompt import build_march_parser_system_prompt

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_march_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    marches = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return marches, cost_usd


async def call_gpt_for_bridge_piles(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    from core.bridge_pile_format_prompt import build_bridge_pile_parser_system_prompt

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_bridge_pile_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    items = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return items, cost_usd




async def call_gpt_for_fbs(
    *,
    user_text: str,
    client: Any,
    image_base64: str | None = None,
    mime_type: str | None = None,
    max_tokens: int = 2500,
) -> tuple[List[Dict[str, Any]], float]:
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": user_text}]
    if image_base64 and mime_type:
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{image_base64}",
                    "detail": "high",
                },
            }
        )

    from core.fbs_format_prompt import build_fbs_parser_system_prompt

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": build_fbs_parser_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        max_tokens=max_tokens,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    items = parse_gpt_response(result_text)
    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    return items, cost_usd


async def recognize_with_gpt_vision(
    image_path: str,
    *,
    client: Optional[Any] = None,
    image_base64: Optional[str] = None,
    mime_type: Optional[str] = None,
) -> tuple[List[Dict], float]:
    """Legacy wrapper: один вызов GPT-4o Vision для OCR."""
    if client is None:
        client = require_openai_client()
    if image_base64 is None or mime_type is None:
        _, image_base64, mime_type = load_image_payload(image_path)
    plates, cost_usd = await call_gpt_for_plates(
        user_text=OCR_USER_PROMPT,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    return plates, cost_usd


async def verify_plates_with_gpt_vision(
    *,
    image_base64: str,
    mime_type: str,
    draft_plates: List[Dict[str, Any]],
    client: Any,
) -> tuple[Dict[str, Any], float]:
    """Этап 2 Verify: сверка черновика с изображением."""
    draft_json = json.dumps(draft_plates, ensure_ascii=False, indent=2)

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": get_verification_prompt(draft_json)},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{image_base64}",
                            "detail": "high",
                        },
                    },
                ],
            }
        ],
        max_tokens=2500,
        temperature=0.0,
    )

    result_text = response.choices[0].message.content or ""
    verify_result = parse_verify_response(result_text)

    tokens_used = response.usage.total_tokens if response.usage else 0
    cost_usd = _estimate_cost_usd(tokens_used)
    print(
        f"[GPT] Verify: токенов {tokens_used}, "
        f"строк {len(verify_result.get('plates') or [])}, "
        f"исправлений {len(verify_result.get('corrections') or [])}"
    )

    return verify_result, cost_usd


class OpenAIProvider:
    """GPT-4o Vision implementation of OcrProvider."""

    def __init__(self, client: Any | None = None) -> None:
        self._client = client

    def _get_client(self) -> Any:
        if self._client is None:
            self._client = require_openai_client()
        return self._client

    async def extract_plates(
        self,
        *,
        user_text: str,
        image_base64: str | None = None,
        mime_type: str | None = None,
        max_tokens: int = 2500,
    ) -> tuple[List[Dict[str, Any]], float]:
        return await call_gpt_for_plates(
            user_text=user_text,
            client=self._get_client(),
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
        return await call_gpt_for_piles(
            user_text=user_text,
            client=self._get_client(),
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
        return await call_gpt_for_steps(
            user_text=user_text,
            client=self._get_client(),
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
        return await call_gpt_for_marches(
            user_text=user_text,
            client=self._get_client(),
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
        return await call_gpt_for_bridge_piles(
            user_text=user_text,
            client=self._get_client(),
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
        return await call_gpt_for_fbs(
            user_text=user_text,
            client=self._get_client(),
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
        return await verify_plates_with_gpt_vision(
            image_base64=image_base64,
            mime_type=mime_type,
            draft_plates=draft_plates,
            client=self._get_client(),
        )
