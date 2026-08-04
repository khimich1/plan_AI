"""Длина маркировки ПБ из plate_name (parse_line) с fallback."""

from __future__ import annotations

from core.plate_line_parser import parse_line


def marking_length_m(plate_name: str, length_m: float | None) -> tuple[float, bool]:
    """
    Возвращает (marking_length_m, used_fallback).

    При успешном разборе plate_name — length_m из парсера (номинал ПБ 64 → 6,4 м).
    Иначе — round(length_m, 1).
    """
    parsed = parse_line(plate_name or "")
    if parsed.parsed and parsed.length_m > 0:
        return round(parsed.length_m, 3), False
    fallback = round(float(length_m or 0.0), 1)
    return fallback, fallback > 0
