#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Распознавание и редактирование списков плит через OCR-провайдеры."""

from __future__ import annotations

from typing import Any, Callable, Awaitable, Dict, Literal, Optional

from core.config.settings import get_settings
from core.ocr.pipeline import run_ocr_pipeline
from core.ocr.providers.openai import (
    GPT_AVAILABLE,
    call_gpt_for_plates,
    load_image_payload,
    require_openai_client,
)
from core.ocr.result import build_result_payload


async def recognize_text_smart(
    image_path: str,
    force_gpt: bool = False,
    show_cost: bool = True,
    mode: Literal["full_gpt", "hybrid"] = "full_gpt",
    verify_enabled: Optional[bool] = None,
    on_status: Optional[Callable[[str], Awaitable[None]]] = None,
) -> Optional[Dict]:
    """
    Распознавание таблицы через OCR pipeline (GigaChat или OpenAI).

    verify_enabled: legacy override (True → always verify, False → never).
    on_status оставлен для обратной совместимости.
    """
    _ = (force_gpt, on_status)
    if mode == "hybrid":
        print("[OCR] ℹ️ Режим hybrid отключен: используется провайдер из OCR_PROVIDER")

    settings = get_settings()
    provider_name = (settings.ocr_provider or "openai").strip().lower()

    if provider_name == "openai" and not GPT_AVAILABLE:
        print("[OCR] ❌ GPT недоступен. Установите: pip install openai")
        return None

    try:
        if on_status:
            await on_status("Распознавание таблицы...")

        result = await run_ocr_pipeline(
            image_path=image_path,
            settings=settings,
            verify_enabled=verify_enabled,
            show_cost=show_cost,
        )
        if result:
            print(
                f"[OCR] ✅ Итого {len(result['plates'])} строк(и), "
                f"method={result.get('method')}, api_calls={result.get('ocr_api_calls')}"
            )
        return result

    except Exception as e:
        print(f"[OCR] ❌ Ошибка OCR: {e}")
        import traceback
        traceback.print_exc()

    return None


async def apply_plates_with_ai(
    *,
    current_plates_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Применяет инструкцию пользователя к списку плит (опционально с изображением).
    Один вызов GPT-4o, temperature=0.
    """
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    client = require_openai_client()
    current_text = (current_plates_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список плит:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя)...")
    plates, cost_usd = await call_gpt_for_plates(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not plates:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(plates)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=plates,
        draft_plates=plates,
        corrections=[],
        row_count_on_image=len(plates),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )
