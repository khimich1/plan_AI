"""Правила укладки ПБ: ГОСТ-таблица, кусок, совместимость длин."""

from __future__ import annotations

PIECE_WIDTH_THRESHOLD_M = 1.2
LENGTH_MIX_TOLERANCE_M = 1.0


def is_piece(width_m: float | None) -> bool:
    """Кусок — ширина строго меньше 1,2 м."""
    return float(width_m or 0.0) < PIECE_WIDTH_THRESHOLD_M


def gost_stack_count(marking_length_m: float) -> int:
    """Число штабелей вдоль кузова 13,2 м по длине маркировки."""
    marking = float(marking_length_m)
    if marking <= 3.3:
        return 4
    if marking <= 4.4:
        return 3
    if marking <= 6.5:
        return 2
    return 1


def markings_compatible(a: float, b: float) -> bool:
    """Разница маркировок ≤ 1,0 м — можно в одну стопку."""
    return abs(float(a) - float(b)) <= LENGTH_MIX_TOLERANCE_M + 1e-9


def body_length_for_stacks(stack_markings: list[float]) -> float:
    """Суммарная занятая длина кузова: каждый штабель — max маркировки в нём."""
    return sum(stack_markings)
