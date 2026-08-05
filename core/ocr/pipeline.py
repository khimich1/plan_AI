#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OCR pipeline: Extract → Parser Gate → Verify policy → [Verify] → result."""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from core.config.settings import Settings, get_settings
from core.ocr.parser_gate import apply_parser_gate
from core.ocr.bridge_pile_parser_gate import apply_bridge_pile_parser_gate
from core.ocr.fbs_parser_gate import apply_fbs_parser_gate
from core.ocr.march_parser_gate import apply_march_parser_gate
from core.ocr.pile_parser_gate import apply_pile_parser_gate
from core.ocr.step_parser_gate import apply_step_parser_gate
from core.ocr.prompts import OCR_USER_PROMPT
from core.ocr.providers.base import OcrProvider
from core.ocr.providers.gigachat import GigaChatProvider
from core.ocr.providers.openai import OpenAIProvider, load_image_payload
from core.ocr.result import build_result_payload
from core.ocr.verify_policy import (
    OcrVerifySettings,
    should_run_bridge_pile_verify,
    should_run_fbs_verify,
    should_run_march_verify,
    should_run_pile_verify,
    should_run_step_verify,
    should_run_verify,
)
from core.bridge_pile_text_normalizer import normalize_bridge_pile_order_text
from core.fbs_text_normalizer import normalize_fbs_order_text
from core.march_text_normalizer import normalize_march_order_text
from core.pile_text_normalizer import normalize_pile_order_text
from core.step_text_normalizer import normalize_step_order_text

_logger = logging.getLogger(__name__)

_OPENAI_USD_TO_RUB = 75.0


def create_ocr_provider(settings: Settings | None = None) -> tuple[OcrProvider, str, str]:
    """Return (provider, provider_name, model_label)."""
    cfg = settings or get_settings()
    provider_name = (cfg.ocr_provider or "openai").strip().lower()
    if provider_name == "gigachat":
        return GigaChatProvider(settings=cfg), "gigachat", cfg.gigachat_model
    return OpenAIProvider(), "openai", "GPT-4o"


def resolve_verify_mode(
    settings: Settings,
    *,
    verify_enabled: Optional[bool] = None,
) -> str:
    if verify_enabled is not None:
        return "always" if verify_enabled else "never"
    return settings.ocr_verify_mode


async def run_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info("[OCR] image_size_kb=%.1f provider=%s verify_mode=%s", image_size_kb, provider_name, verify_mode)

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    plates, extract_cost = await provider.extract_plates(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not plates:
        return None

    plates = apply_parser_gate(plates)
    draft_plates = [dict(p) for p in plates]

    run_verify, verify_reason = should_run_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        plates=plates,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(plates)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_plates,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_plates = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_plates)

            if not verified_plates:
                verify_failed = True
            else:
                plates = apply_parser_gate(verified_plates)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] Verify failed: %s", exc)
            plates = draft_plates
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(plates),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    return build_result_payload(
        plates=plates,
        draft_plates=draft_plates,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )


async def run_pile_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    extract_piles = getattr(provider, "extract_piles", None)
    if extract_piles is None:
        raise RuntimeError(f"Провайдер {provider_name} не поддерживает OCR свай.")

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info(
        "[OCR] pile image_size_kb=%.1f provider=%s verify_mode=%s",
        image_size_kb,
        provider_name,
        verify_mode,
    )

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    piles, extract_cost = await extract_piles(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not piles:
        return None

    piles = apply_pile_parser_gate(piles)
    draft_piles = [dict(p) for p in piles]

    run_verify, verify_reason = should_run_pile_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        piles=piles,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(piles)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_piles,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_piles = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_piles)

            if not verified_piles:
                verify_failed = True
            else:
                piles = apply_pile_parser_gate(verified_piles)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] Pile verify failed: %s", exc)
            piles = draft_piles
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] pile api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(piles),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    payload = build_result_payload(
        plates=piles,
        draft_plates=draft_piles,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )
    normalized = normalize_pile_order_text(payload.get("text", ""))
    payload["text"] = normalized.normalized_text
    return payload


async def run_step_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    extract_steps = getattr(provider, "extract_steps", None)
    if extract_steps is None:
        raise RuntimeError(f"Провайдер {provider_name} не поддерживает OCR ступеней.")

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info(
        "[OCR] step image_size_kb=%.1f provider=%s verify_mode=%s",
        image_size_kb,
        provider_name,
        verify_mode,
    )

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    steps, extract_cost = await extract_steps(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not steps:
        return None

    steps = apply_step_parser_gate(steps)
    draft_steps = [dict(s) for s in steps]

    run_verify, verify_reason = should_run_step_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        steps=steps,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(steps)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_steps,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_steps = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_steps)

            if not verified_steps:
                verify_failed = True
            else:
                steps = apply_step_parser_gate(verified_steps)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] Step verify failed: %s", exc)
            steps = draft_steps
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] step api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(steps),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    payload = build_result_payload(
        plates=steps,
        draft_plates=draft_steps,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )
    normalized = normalize_step_order_text(payload.get("text", ""))
    payload["text"] = normalized.normalized_text
    return payload


async def run_march_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    extract_marches = getattr(provider, "extract_marches", None)
    if extract_marches is None:
        raise RuntimeError(f"Провайдер {provider_name} не поддерживает OCR маршей.")

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info(
        "[OCR] march image_size_kb=%.1f provider=%s verify_mode=%s",
        image_size_kb,
        provider_name,
        verify_mode,
    )

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    marches, extract_cost = await extract_marches(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not marches:
        return None

    marches = apply_march_parser_gate(marches)
    draft_marches = [dict(m) for m in marches]

    run_verify, verify_reason = should_run_march_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        marches=marches,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(marches)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_marches,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_marches = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_marches)

            if not verified_marches:
                verify_failed = True
            else:
                marches = apply_march_parser_gate(verified_marches)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] March verify failed: %s", exc)
            marches = draft_marches
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] march api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(marches),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    payload = build_result_payload(
        plates=marches,
        draft_plates=draft_marches,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )
    normalized = normalize_march_order_text(payload.get("text", ""))
    payload["text"] = normalized.normalized_text
    return payload


async def run_bridge_pile_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    extract_bridge_piles = getattr(provider, "extract_bridge_piles", None)
    if extract_bridge_piles is None:
        raise RuntimeError(f"Провайдер {provider_name} не поддерживает OCR мостовых свай.")

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info(
        "[OCR] bridge_pile image_size_kb=%.1f provider=%s verify_mode=%s",
        image_size_kb,
        provider_name,
        verify_mode,
    )

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    bridge_piles, extract_cost = await extract_bridge_piles(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not bridge_piles:
        return None

    bridge_piles = apply_bridge_pile_parser_gate(bridge_piles)
    draft_bridge_piles = [dict(item) for item in bridge_piles]

    run_verify, verify_reason = should_run_bridge_pile_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        bridge_piles=bridge_piles,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(bridge_piles)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_bridge_piles,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_bridge_piles = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_bridge_piles)

            if not verified_bridge_piles:
                verify_failed = True
            else:
                bridge_piles = apply_bridge_pile_parser_gate(verified_bridge_piles)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] Bridge pile verify failed: %s", exc)
            bridge_piles = draft_bridge_piles
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] bridge_pile api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(bridge_piles),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    payload = build_result_payload(
        plates=bridge_piles,
        draft_plates=draft_bridge_piles,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )
    normalized = normalize_bridge_pile_order_text(payload.get("text", ""))
    payload["text"] = normalized.normalized_text
    return payload


async def run_fbs_ocr_pipeline(
    *,
    image_path: str,
    provider: OcrProvider | None = None,
    settings: Settings | None = None,
    user_text: str = OCR_USER_PROMPT,
    verify_enabled: Optional[bool] = None,
    show_cost: bool = True,
) -> Optional[Dict[str, Any]]:
    cfg = settings or get_settings()
    if provider is None:
        provider, provider_name, model_label = create_ocr_provider(cfg)
    else:
        provider_name = (cfg.ocr_provider or "openai").strip().lower()
        model_label = cfg.gigachat_model if provider_name == "gigachat" else "GPT-4o"

    extract_fbs = getattr(provider, "extract_fbs", None)
    if extract_fbs is None:
        raise RuntimeError(f"Провайдер {provider_name} не поддерживает OCR ФБС.")

    verify_mode = resolve_verify_mode(cfg, verify_enabled=verify_enabled)
    verify_settings = OcrVerifySettings(
        max_rows=cfg.ocr_verify_auto_max_rows,
        min_confidence=cfg.ocr_verify_auto_min_confidence,
        max_bytes=cfg.ocr_verify_auto_max_bytes,
    )

    image_data, image_base64, mime_type = load_image_payload(image_path)
    image_size_kb = len(image_data) / 1024
    _logger.info(
        "[OCR] fbs image_size_kb=%.1f provider=%s verify_mode=%s",
        image_size_kb,
        provider_name,
        verify_mode,
    )

    api_calls = 0
    cost_usd = 0.0
    cost_rub = 0.0

    fbs, extract_cost = await extract_fbs(
        user_text=user_text,
        image_base64=image_base64,
        mime_type=mime_type,
    )
    api_calls += 1
    if provider_name == "gigachat":
        cost_rub += extract_cost
    else:
        cost_usd += extract_cost

    if not fbs:
        return None

    fbs = apply_fbs_parser_gate(fbs)
    draft_fbs = [dict(item) for item in fbs]

    run_verify, verify_reason = should_run_fbs_verify(
        mode=verify_mode,
        max_api_calls=cfg.ocr_max_api_calls,
        image_size_bytes=len(image_data),
        fbs=fbs,
        settings=verify_settings,
    )

    verify_applied = False
    verify_failed = False
    corrections: list[Dict[str, Any]] = []
    row_count_on_image: Optional[int] = len(fbs)
    verify_skipped_reason: Optional[str] = None
    verify_applied_reason: Optional[str] = None
    method = model_label

    if run_verify:
        verify_applied = True
        verify_applied_reason = verify_reason
        try:
            verify_result, verify_cost = await provider.verify_plates(
                image_base64=image_base64,
                mime_type=mime_type,
                draft_plates=draft_fbs,
            )
            api_calls += 1
            if provider_name == "gigachat":
                cost_rub += verify_cost
            else:
                cost_usd += verify_cost

            verified_fbs = verify_result.get("plates") or []
            corrections = list(verify_result.get("corrections") or [])
            row_count_on_image = verify_result.get("row_count_on_image")
            if row_count_on_image is None:
                row_count_on_image = len(verified_fbs)

            if not verified_fbs:
                verify_failed = True
            else:
                fbs = apply_fbs_parser_gate(verified_fbs)
        except Exception as exc:
            verify_failed = True
            _logger.exception("[OCR] FBS verify failed: %s", exc)
            fbs = draft_fbs
        method = f"{model_label}+verify"
    else:
        verify_skipped_reason = verify_reason

    if provider_name != "gigachat" and cost_rub == 0.0 and cost_usd > 0.0:
        cost_rub = cost_usd * _OPENAI_USD_TO_RUB

    verify_decision = verify_applied_reason or verify_skipped_reason or "unknown"
    _logger.info(
        "[OCR] fbs api_calls=%s cost_rub=%.4f verify_decision=%s rows=%s",
        api_calls,
        cost_rub,
        verify_decision,
        len(fbs),
    )

    if show_cost:
        if provider_name == "gigachat":
            print(f"[OCR] 💰 Стоимость: {cost_rub:.4f}₽")
        else:
            print(f"[OCR] 💰 Стоимость: ${cost_usd:.4f} (~{cost_rub:.2f}₽)")

    payload = build_result_payload(
        plates=fbs,
        draft_plates=draft_fbs,
        corrections=corrections,
        row_count_on_image=row_count_on_image,
        method=method,
        verify_applied=verify_applied,
        verify_failed=verify_failed,
        cost_usd=cost_usd if provider_name != "gigachat" else 0.0,
        cost_rub=cost_rub,
        api_calls=api_calls,
        verify_skipped_reason=verify_skipped_reason,
        verify_applied_reason=verify_applied_reason,
    )
    normalized = normalize_fbs_order_text(payload.get("text", ""))
    payload["text"] = normalized.normalized_text
    return payload
