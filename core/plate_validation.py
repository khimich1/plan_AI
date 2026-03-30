#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Бизнес-валидация распознанной строки заказа плит.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ValidationResult:
    ok: bool
    reason_code: str = ""
    reason_text: str = ""


def validate_plate_values(width_m: float, length_m: float, qty: int) -> ValidationResult:
    """
    Единая проверка значений перед добавлением в рабочие списки.
    """
    if width_m <= 0 or length_m <= 0:
        return ValidationResult(False, "non_positive_dimensions", "длина/ширина должны быть больше 0")
    if qty is None or qty <= 0:
        return ValidationResult(False, "invalid_qty", "количество должно быть 1 или больше")
    if qty > 500:
        return ValidationResult(False, "qty_too_large", "слишком большое количество (макс 500 на строку)")
    if length_m < 0.5 or length_m > 15.0:
        return ValidationResult(False, "invalid_length", f"нереалистичная длина {length_m}м")
    if width_m < 0.1 or width_m > 2.0:
        return ValidationResult(False, "invalid_width", f"нереалистичная ширина {width_m}м")
    return ValidationResult(True)
