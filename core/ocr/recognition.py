#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Распознавание и редактирование списков изделий через OCR-провайдеры."""

from __future__ import annotations

from typing import Any, Callable, Awaitable, Dict, Literal, Optional

from core.config.settings import get_settings
from core.ocr.pipeline import (
    create_ocr_provider,
    run_ocr_pipeline,
    run_bridge_pile_ocr_pipeline,
    run_fbs_ocr_pipeline,
    run_march_ocr_pipeline,
    run_pile_ocr_pipeline,
    run_step_ocr_pipeline,
)
from core.ocr.providers.openai import (
    GPT_AVAILABLE,
    load_image_payload,
)
from core.ocr.result import build_result_payload

ProductType = Literal["plates", "piles", "steps", "marches", "bridge_piles", "fbs"]

_APPLY_AI_LABELS: dict[str, str] = {
    "plates": "плит",
    "piles": "свай",
    "steps": "ступеней",
    "marches": "маршей",
    "bridge_piles": "мостовых свай",
    "fbs": "ФБС",
}

_APPLY_AI_EXTRACT: dict[str, str] = {
    "plates": "extract_plates",
    "piles": "extract_piles",
    "steps": "extract_steps",
    "marches": "extract_marches",
    "bridge_piles": "extract_bridge_piles",
    "fbs": "extract_fbs",
}

_OPENAI_USD_TO_RUB = 75.0


async def recognize_text_smart(
    image_path: str,
    force_gpt: bool = False,
    show_cost: bool = True,
    mode: Literal["full_gpt", "hybrid"] = "full_gpt",
    verify_enabled: Optional[bool] = None,
    on_status: Optional[Callable[[str], Awaitable[None]]] = None,
    product_type: ProductType = "plates",
) -> Optional[Dict]:
    """
    Распознавание таблицы через OCR pipeline (GigaChat или OpenAI).

    verify_enabled: legacy override (True → always verify, False → never).
    product_type: "plates" (default), "piles", "steps", "marches", "bridge_piles" или "fbs".
    on_status оставлен для обратной совместимости.
    """
    _ = (force_gpt, on_status)
    if mode == "hybrid":
        print("[OCR] ℹ️ Режим hybrid отключен: используется провайдер из OCR_PROVIDER")

    settings = get_settings()
    provider_name = (settings.ocr_provider or "openai").strip().lower()
    normalized_product_type = (product_type or "plates").strip().lower()

    if provider_name == "openai" and not GPT_AVAILABLE:
        print("[OCR] ❌ GPT недоступен. Установите: pip install openai")
        return None

    try:
        if on_status:
            label = _APPLY_AI_LABELS.get(normalized_product_type, "плит")
            await on_status(f"Распознавание таблицы {label}...")

        if normalized_product_type == "piles":
            pipeline = run_pile_ocr_pipeline
        elif normalized_product_type == "bridge_piles":
            pipeline = run_bridge_pile_ocr_pipeline
        elif normalized_product_type == "fbs":
            pipeline = run_fbs_ocr_pipeline
        elif normalized_product_type == "steps":
            pipeline = run_step_ocr_pipeline
        elif normalized_product_type == "marches":
            pipeline = run_march_ocr_pipeline
        else:
            pipeline = run_ocr_pipeline
        result = await pipeline(
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


async def _apply_with_ai(
    *,
    product_type: ProductType,
    current_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Apply user instruction via OCR_PROVIDER (GigaChat or OpenAI), one extract call."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    label = _APPLY_AI_LABELS[product_type]
    extract_method = _APPLY_AI_EXTRACT[product_type]
    provider, provider_name, model_label = create_ocr_provider()

    current = (current_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список {label}:\n{current}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print(f"[AI] {model_label} (один вызов, инструкция пользователя, {label})...")
    extract_fn = getattr(provider, extract_method)
    items, extract_cost = await extract_fn(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not items:
        return None

    if provider_name == "gigachat":
        cost_usd = 0.0
        cost_rub = float(extract_cost)
    else:
        cost_usd = float(extract_cost)
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    if show_cost:
        if provider_name == "gigachat":
            print(f"[AI] 💰 Стоимость: ~{cost_rub:.2f}₽")
        else:
            print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    method = f"{model_label}+ai"
    print(f"[AI] ✅ Итого {len(items)} строк(и), method={method}")
    return build_result_payload(
        plates=items,
        draft_plates=items,
        corrections=[],
        row_count_on_image=len(items),
        method=method,
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )


async def apply_plates_with_ai(
    *,
    current_plates_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку плит (опционально с изображением)."""
    return await _apply_with_ai(
        product_type="plates",
        current_text=current_plates_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )


async def apply_piles_with_ai(
    *,
    current_piles_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку свай (опционально с изображением)."""
    return await _apply_with_ai(
        product_type="piles",
        current_text=current_piles_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )


async def apply_steps_with_ai(
    *,
    current_steps_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку ступеней (опционально с изображением)."""
    return await _apply_with_ai(
        product_type="steps",
        current_text=current_steps_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )


async def apply_marches_with_ai(
    *,
    current_marches_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку маршей (опционально с изображением)."""
    return await _apply_with_ai(
        product_type="marches",
        current_text=current_marches_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )


async def apply_bridge_piles_with_ai(
    *,
    current_bridge_piles_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку мостовых свай."""
    return await _apply_with_ai(
        product_type="bridge_piles",
        current_text=current_bridge_piles_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )


async def apply_fbs_with_ai(
    *,
    current_fbs_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку ФБС."""
    return await _apply_with_ai(
        product_type="fbs",
        current_text=current_fbs_text,
        user_instruction=user_instruction,
        image_path=image_path,
        show_cost=show_cost,
    )
