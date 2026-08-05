#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Распознавание и редактирование списков плит через OCR-провайдеры."""

from __future__ import annotations

from typing import Any, Callable, Awaitable, Dict, Literal, Optional

from core.config.settings import get_settings
from core.ocr.pipeline import (
    run_ocr_pipeline,
    run_bridge_pile_ocr_pipeline,
    run_fbs_ocr_pipeline,
    run_march_ocr_pipeline,
    run_pile_ocr_pipeline,
    run_step_ocr_pipeline,
)
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
    product_type: Literal["plates", "piles", "steps", "marches", "bridge_piles", "fbs"] = "plates",
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
            if normalized_product_type == "piles":
                label = "свай"
            elif normalized_product_type == "bridge_piles":
                label = "мостовых свай"
            elif normalized_product_type == "fbs":
                label = "ФБС"
            elif normalized_product_type == "steps":
                label = "ступеней"
            elif normalized_product_type == "marches":
                label = "маршей"
            else:
                label = "плит"
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


async def apply_piles_with_ai(
    *,
    current_piles_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку свай (опционально с изображением)."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    from core.ocr.providers.openai import call_gpt_for_piles

    client = require_openai_client()
    current_text = (current_piles_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список свай:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя, сваи)...")
    piles, cost_usd = await call_gpt_for_piles(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not piles:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(piles)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=piles,
        draft_plates=piles,
        corrections=[],
        row_count_on_image=len(piles),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )


async def apply_steps_with_ai(
    *,
    current_steps_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку ступеней (опционально с изображением)."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    from core.ocr.providers.openai import call_gpt_for_steps

    client = require_openai_client()
    current_text = (current_steps_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список ступеней:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя, ступени)...")
    steps, cost_usd = await call_gpt_for_steps(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not steps:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(steps)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=steps,
        draft_plates=steps,
        corrections=[],
        row_count_on_image=len(steps),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )


async def apply_marches_with_ai(
    *,
    current_marches_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку маршей (опционально с изображением)."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    from core.ocr.providers.openai import call_gpt_for_marches

    client = require_openai_client()
    current_text = (current_marches_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список маршей:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя, марши)...")
    marches, cost_usd = await call_gpt_for_marches(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not marches:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(marches)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=marches,
        draft_plates=marches,
        corrections=[],
        row_count_on_image=len(marches),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )


async def apply_bridge_piles_with_ai(
    *,
    current_bridge_piles_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку мостовых свай."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    from core.ocr.providers.openai import call_gpt_for_bridge_piles

    client = require_openai_client()
    current_text = (current_bridge_piles_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список мостовых свай:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя, мостовые сваи)...")
    items, cost_usd = await call_gpt_for_bridge_piles(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not items:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(items)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=items,
        draft_plates=items,
        corrections=[],
        row_count_on_image=len(items),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )


async def apply_fbs_with_ai(
    *,
    current_fbs_text: str,
    user_instruction: str,
    image_path: str | None = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    """Применяет инструкцию пользователя к списку ФБС."""
    instruction = (user_instruction or "").strip()
    if not instruction:
        raise ValueError("Инструкция для ИИ не может быть пустой.")

    from core.ocr.providers.openai import call_gpt_for_fbs

    client = require_openai_client()
    current_text = (current_fbs_text or "").strip() or "(пусто)"
    user_text = (
        f"Текущий список ФБС:\n{current_text}\n\n"
        f"Инструкция пользователя:\n{instruction}"
    )

    image_base64: str | None = None
    mime_type: str | None = None
    if image_path:
        _, image_base64, mime_type = load_image_payload(image_path)

    print("[AI] GPT-4o (один вызов, инструкция пользователя, ФБС)...")
    items, cost_usd = await call_gpt_for_fbs(
        user_text=user_text,
        client=client,
        image_base64=image_base64,
        mime_type=mime_type,
    )

    if not items:
        return None

    cost_rub = cost_usd * 75
    if show_cost:
        print(f"[AI] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    print(f"[AI] ✅ Итого {len(items)} строк(и), method=GPT-4o+ai")
    return build_result_payload(
        plates=items,
        draft_plates=items,
        corrections=[],
        row_count_on_image=len(items),
        method="GPT-4o+ai",
        verify_applied=False,
        verify_failed=False,
        cost_usd=cost_usd,
        cost_rub=cost_rub,
        api_calls=1,
    )
