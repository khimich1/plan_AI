"""Парсинг и нормализация срока изготовления (КП): даты и относительные фразы.

Логика согласована с bot/handlers/commercial.py и UI-подсказками:
сначала абсолютные даты, затем «N дней» / «N недель» (чтобы даты не
пересекались с относительными шаблонами лишний раз).
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Literal

_DATE_FORMATS = ("%d.%m.%Y", "%Y-%m-%d")
_DAYS_RE = re.compile(r"(\d+)\s*(?:дн|день|дней|day|days)", re.IGNORECASE)
_WEEKS_RE = re.compile(r"(\d+)\s*(?:нед|недел|недели|week|weeks)", re.IGNORECASE)

# Сообщение при невозможности распознать ввод (строгие пути: архив, мастер КП)
_PARSE_HINT_RU = (
    "Укажите срок в формате ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, N дней или N недель."
)

DEFAULT_EXECUTION_TERMS_DAYS = 14

ExecutionTermsPolicy = Literal["strict", "default_if_empty"]


def parse_execution_terms_to_datetime(
    raw: str,
    *,
    now: datetime | None = None,
) -> datetime | None:
    """Возвращает дату дедлайна или None, если распознать не удалось."""
    value = (raw or "").strip()
    if not value:
        return None

    clock = now if now is not None else datetime.now()

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue

    match_days = _DAYS_RE.search(value)
    if match_days:
        return clock + timedelta(days=int(match_days.group(1)))

    match_weeks = _WEEKS_RE.search(value)
    if match_weeks:
        return clock + timedelta(weeks=int(match_weeks.group(1)))

    return None


def parse_execution_terms(
    value: str,
    *,
    policy: ExecutionTermsPolicy,
    now: datetime | None = None,
    default_days: int = DEFAULT_EXECUTION_TERMS_DAYS,
) -> tuple[str, bool]:
    """Парсит срок изготовления и возвращает (ДД.ММ.ГГГГ, used_default).

    strict — пустой или нераспознанный ввод → ValueError.
    default_if_empty — при неудаче подставляет now + default_days, used_default=True.
    """
    clock = now if now is not None else datetime.now()
    text = (value or "").strip()
    deadline = parse_execution_terms_to_datetime(text, now=clock)

    if deadline is not None:
        return deadline.strftime("%d.%m.%Y"), False

    if policy == "strict":
        if not text:
            raise ValueError("Укажите срок изготовления.")
        raise ValueError(_PARSE_HINT_RU)

    fallback = clock + timedelta(days=default_days)
    return fallback.strftime("%d.%m.%Y"), True


def normalize_execution_terms_to_ddmmyyyy(
    raw: str,
    *,
    now: datetime | None = None,
) -> str:
    """Нормализует ввод к строке ДД.ММ.ГГГГ или бросает ValueError."""
    formatted, _ = parse_execution_terms(raw, policy="strict", now=now)
    return formatted
