#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Backward-compatibility shim for core.ocr package.

Импорты вида `from core.ocr_gpt import recognize_text_smart` продолжают работать.
"""

from core.ocr import (
    _validate_plate_item,
    apply_plates_with_ai,
    estimate_monthly_cost,
    format_corrections_for_user,
    get_recognition_prompt,
    get_verification_prompt,
    parse_gpt_response,
    parse_verify_response,
    recognize_text_smart,
)
from core.ocr.providers.openai import (
    AsyncOpenAI,
    GPT_AVAILABLE,
    call_gpt_for_plates,
    image_mime_type,
    load_image_payload,
    recognize_with_gpt_vision,
    require_openai_client,
    verify_plates_with_gpt_vision,
)

# Private aliases preserved for any legacy internal imports.
_load_image_payload = load_image_payload
_image_mime_type = image_mime_type
_require_openai_client = require_openai_client
_call_gpt_for_plates = call_gpt_for_plates

__all__ = [
    "AsyncOpenAI",
    "GPT_AVAILABLE",
    "_call_gpt_for_plates",
    "_image_mime_type",
    "_load_image_payload",
    "_require_openai_client",
    "_validate_plate_item",
    "apply_plates_with_ai",
    "call_gpt_for_plates",
    "estimate_monthly_cost",
    "format_corrections_for_user",
    "get_recognition_prompt",
    "get_verification_prompt",
    "image_mime_type",
    "load_image_payload",
    "parse_gpt_response",
    "parse_verify_response",
    "recognize_text_smart",
    "recognize_with_gpt_vision",
    "require_openai_client",
    "verify_plates_with_gpt_vision",
]

if __name__ == "__main__":
    print("💰 Оценка месячных затрат на OCR (один вызов GPT-4o):")
    print("=" * 50)

    for count in [100, 500, 1000, 5000]:
        costs = estimate_monthly_cost(count)
        print(f"\n📊 {count} фото в месяц:")
        print(f"  • GPT-4o: ${costs['gpt_only']:.2f} (~{costs['gpt_only'] * 75:.0f}₽)")
