"""Коды причин not_fit и warnings для движка укладки."""

from __future__ import annotations

from enum import Enum


class NotFitReason(str, Enum):
    WEIGHT_LIMIT = "weight_limit"
    TIER_LIMIT = "tier_limit"
    LENGTH_MIX = "length_mix"
    BODY_LENGTH = "body_length"
    NO_STACK_SLOT = "no_stack_slot"
    PIECE_PRIORITY = "piece_priority"
    NEXT_TRIP = "next_trip"


class WarningCode(str, Enum):
    KP_MIX = "kp_mix"
    PIECE_SUBOPTIMAL = "piece_suboptimal"
    MARKING_FALLBACK = "marking_fallback"


REASON_TEXT: dict[NotFitReason, str] = {
    NotFitReason.WEIGHT_LIMIT: "Превышен лимит веса класса ТС",
    NotFitReason.TIER_LIMIT: "Не помещается в 4 яруса",
    NotFitReason.LENGTH_MIX: "Нет стопки с допустимой разницей маркировок (≤1 м)",
    NotFitReason.BODY_LENGTH: "Не помещается в длину кузова 13,2 м",
    NotFitReason.NO_STACK_SLOT: "Нет места в ширине/ярусе стопки",
    NotFitReason.PIECE_PRIORITY: "Кусок не прошёл приоритеты добора",
    NotFitReason.NEXT_TRIP: "Оставлено на следующий рейс",
}

WARNING_TEXT: dict[WarningCode, str] = {
    WarningCode.KP_MIX: "В рейсе позиции из нескольких КП",
    WarningCode.PIECE_SUBOPTIMAL: "Кусок положен не по оптимальному приоритету",
    WarningCode.MARKING_FALLBACK: "Маркировка взята из length_m (plate_name не разобран)",
}
