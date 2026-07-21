"""Утилиты для приведения доменных словарей к JSON-safe виду.

В рантайме в ``optimization_result`` попадают вспомогательные объекты
(например, :class:`core.plate_audit.PlateAudit`), которые нельзя сериализовать
ни в JSON-файл, ни в HTTP-ответ FastAPI. Перед сохранением и отправкой клиенту
такие ключи нужно убирать.

Хелперы возвращают копии словарей и не мутируют исходные структуры —
продолжать использовать in-memory объекты с полной информацией допустимо.
Модуль лежит в shared-слое ``core/``, чтобы им могли пользоваться и
web-сервисы (``app/...``), и телеграм-бот (``bot/...``).
"""
from __future__ import annotations

from typing import Any

PLATE_AUDIT_KEY = "_plate_audit"


def strip_plate_audit(optimization_result: dict[str, Any] | None) -> dict[str, Any]:
    """Возвращает копию ``optimization_result`` без ключа ``_plate_audit``.

    Если на вход подан ``None`` или не-словарь — возвращается пустой dict,
    чтобы вызывающий код мог безопасно положить результат на место исходного
    значения.
    """
    if not isinstance(optimization_result, dict):
        return {}
    return {key: value for key, value in optimization_result.items() if key != PLATE_AUDIT_KEY}


def strip_plate_audit_from_plan(plan: dict[str, Any]) -> dict[str, Any]:
    """Возвращает копию плана, где ``optimization_result`` очищен от PlateAudit.

    Остальные поля плана копируются по ссылке (shallow copy), что ожидаемо:
    фикс не должен изменять форму плана и его lookup-таблицы.
    """
    safe = dict(plan)
    opt = plan.get("optimization_result")
    if isinstance(opt, dict):
        safe["optimization_result"] = strip_plate_audit(opt)
    return safe


def strip_plate_audit_from_plan_by_load(plan_by_load: dict[str, Any]) -> dict[str, Any]:
    """То же для ``plan_by_load``: значения — dict-и с тем же набором ключей."""
    out: dict[str, Any] = {}
    for key, value in plan_by_load.items():
        out[key] = strip_plate_audit(value) if isinstance(value, dict) else value
    return out
