"""Утилиты рабочего календаря для производственного планирования."""

from __future__ import annotations

import json
import logging
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

CALENDAR_PATH = Path(__file__).resolve().parent.parent / "data" / "work_calendar.json"


def _load_calendar_data() -> dict:
    """Загружает JSON-данные рабочего календаря."""
    if not CALENDAR_PATH.exists():
        return {"extra_holidays": [], "extra_workdays": []}

    try:
        with open(CALENDAR_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as exc:
        logger.exception("Не удалось загрузить рабочий календарь: %s", exc)
        return {"extra_holidays": [], "extra_workdays": []}

    if not isinstance(data, dict):
        return {"extra_holidays": [], "extra_workdays": []}
    return data


def _parse_iso_dates(values: list) -> set[date]:
    """Преобразует список строк ISO-дат в set[date]."""
    result: set[date] = set()
    for value in values:
        if not isinstance(value, str):
            continue
        try:
            result.add(date.fromisoformat(value))
        except ValueError:
            logger.warning("Пропущена дата с неверным форматом: %s", value)
    return result


def load_holidays() -> set[date]:
    """Возвращает множество дополнительных нерабочих дней."""
    data = _load_calendar_data()
    holidays = data.get("extra_holidays", [])
    if not isinstance(holidays, list):
        return set()
    return _parse_iso_dates(holidays)


def load_extra_workdays() -> set[date]:
    """Возвращает множество дополнительных рабочих дней (в т.ч. переносов)."""
    data = _load_calendar_data()
    workdays = data.get("extra_workdays", [])
    if not isinstance(workdays, list):
        return set()
    return _parse_iso_dates(workdays)


def is_working_day(
    day: date,
    holidays: Optional[set[date]] = None,
    extra_workdays: Optional[set[date]] = None,
) -> bool:
    """Проверяет, является ли день рабочим."""
    holidays = load_holidays() if holidays is None else holidays
    extra_workdays = load_extra_workdays() if extra_workdays is None else extra_workdays

    if day in extra_workdays:
        return True
    if day in holidays:
        return False
    return day.weekday() < 5


def nth_working_day(
    start_day: date,
    n: int,
    holidays: Optional[set[date]] = None,
    extra_workdays: Optional[set[date]] = None,
) -> date:
    """
    Возвращает дату N-го рабочего дня, считая от start_day включительно.

    n = 1 -> первый рабочий день в диапазоне от start_day.
    """
    if n < 1:
        raise ValueError("n должно быть >= 1")

    holidays = load_holidays() if holidays is None else holidays
    extra_workdays = load_extra_workdays() if extra_workdays is None else extra_workdays

    current = start_day
    found = 0
    while True:
        if is_working_day(current, holidays, extra_workdays):
            found += 1
            if found == n:
                return current
        current += timedelta(days=1)
