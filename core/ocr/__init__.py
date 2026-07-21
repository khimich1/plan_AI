#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пакет OCR: распознавание списков ЖБ-плит через провайдеры (OpenAI, GigaChat).
"""

from core.ocr.parser_gate import apply_parser_gate
from core.ocr.pipeline import create_ocr_provider, run_ocr_pipeline
from core.ocr.parsing import (
    _validate_plate_item,
    parse_gpt_response,
    parse_verify_response,
)
from core.ocr.prompts import get_recognition_prompt, get_verification_prompt
from core.ocr.recognition import apply_plates_with_ai, recognize_text_smart
from core.ocr.result import estimate_monthly_cost, format_corrections_for_user
from core.ocr.verify_policy import OcrVerifySettings, should_run_verify

__all__ = [
    "apply_plates_with_ai",
    "recognize_text_smart",
    "create_ocr_provider",
    "run_ocr_pipeline",
    "apply_parser_gate",
    "should_run_verify",
    "OcrVerifySettings",
    "parse_gpt_response",
    "parse_verify_response",
    "_validate_plate_item",
    "get_recognition_prompt",
    "get_verification_prompt",
    "format_corrections_for_user",
    "estimate_monthly_cost",
]
