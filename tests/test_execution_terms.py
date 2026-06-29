from __future__ import annotations

from datetime import datetime

import pytest

from app.services.execution_terms_service import ExecutionTermsService
from core.execution_terms import (
    DEFAULT_EXECUTION_TERMS_DAYS,
    normalize_execution_terms_to_ddmmyyyy,
    parse_execution_terms,
    parse_execution_terms_to_datetime,
)


def _patch_execution_terms_clock(monkeypatch: pytest.MonkeyPatch, base: datetime) -> None:
    """Python 3.12+: нельзя monkeypatch-ить datetime.now на классе datetime; подменяем ссылку в модуле."""

    class FixedNowDatetime(datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(base.year, base.month, base.day, base.hour, base.minute, base.second)

    monkeypatch.setattr("core.execution_terms.datetime", FixedNowDatetime)


@pytest.fixture()
def fixed_now(monkeypatch: pytest.MonkeyPatch) -> datetime:
    """Фиксирует «сегодня», чтобы «N дней» и fallback давали предсказуемый дедлайн."""
    base = datetime(2026, 5, 31, 12, 0, 0)
    _patch_execution_terms_clock(monkeypatch, base)
    return base


@pytest.mark.parametrize(
    ("raw", "expected_ddmmyyyy"),
    [
        ("05.06.2026", "05.06.2026"),
        ("2026-06-05", "05.06.2026"),
        ("5 дней", "05.06.2026"),
        ("2 недели", "14.06.2026"),
    ],
)
def test_parse_execution_terms_strict_valid(
    fixed_now: datetime,
    raw: str,
    expected_ddmmyyyy: str,
) -> None:
    formatted, used_default = parse_execution_terms(raw, policy="strict", now=fixed_now)
    assert formatted == expected_ddmmyyyy
    assert used_default is False


@pytest.mark.parametrize(
    "raw",
    [
        "",
        "   ",
        "скоро",
        "неизвестный формат",
    ],
)
def test_parse_execution_terms_strict_invalid(fixed_now: datetime, raw: str) -> None:
    with pytest.raises(ValueError):
        parse_execution_terms(raw, policy="strict", now=fixed_now)


@pytest.mark.parametrize(
    ("raw", "expected_ddmmyyyy", "used_default"),
    [
        ("05.06.2026", "05.06.2026", False),
        ("2026-06-05", "05.06.2026", False),
        ("5 дней", "05.06.2026", False),
        ("", "14.06.2026", True),
        ("   ", "14.06.2026", True),
        ("скоро", "14.06.2026", True),
    ],
)
def test_parse_execution_terms_default_if_empty(
    fixed_now: datetime,
    raw: str,
    expected_ddmmyyyy: str,
    used_default: bool,
) -> None:
    formatted, got_default = parse_execution_terms(
        raw,
        policy="default_if_empty",
        now=fixed_now,
    )
    assert formatted == expected_ddmmyyyy
    assert got_default is used_default


def test_default_if_empty_uses_configured_offset(fixed_now: datetime) -> None:
    formatted, used_default = parse_execution_terms(
        "",
        policy="default_if_empty",
        now=fixed_now,
        default_days=DEFAULT_EXECUTION_TERMS_DAYS,
    )
    assert formatted == "14.06.2026"
    assert used_default is True


@pytest.mark.parametrize(
    ("raw", "expected_ddmmyyyy"),
    [
        ("05.06.2026", "05.06.2026"),
        ("2026-06-05", "05.06.2026"),
    ],
)
def test_normalize_absolute_dates_ddmmyyyy_and_iso(raw: str, expected_ddmmyyyy: str) -> None:
    assert normalize_execution_terms_to_ddmmyyyy(raw) == expected_ddmmyyyy


def test_normalize_five_days_relative(fixed_now: datetime) -> None:
    assert normalize_execution_terms_to_ddmmyyyy("5 дней", now=fixed_now) == "05.06.2026"


def test_parse_execution_terms_to_datetime_matches_normalize(fixed_now: datetime) -> None:
    dt = parse_execution_terms_to_datetime("5 дней", now=fixed_now)
    assert dt is not None
    assert dt.strftime("%d.%m.%Y") == "05.06.2026"


def test_execution_terms_service_absolute_inputs() -> None:
    svc = ExecutionTermsService()
    assert svc.normalize("05.06.2026") == "05.06.2026"
    assert svc.normalize("2026-06-05") == "05.06.2026"


def test_execution_terms_service_five_days_relative(monkeypatch: pytest.MonkeyPatch) -> None:
    base = datetime(2026, 5, 31, 12, 0, 0)
    _patch_execution_terms_clock(monkeypatch, base)
    assert ExecutionTermsService().normalize("5 дней") == "05.06.2026"
