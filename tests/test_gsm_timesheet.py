"""Т-13 parser: personnel numbers, day grid, header guard."""

from __future__ import annotations

from datetime import date

import pytest

from core.gsm.driver_pool import ELIGIBLE_CODES
from core.gsm.timesheet import (
    TimesheetParseError,
    parse_personnel_number,
    parse_timesheet_xlsx,
    period_from_filename,
)
from tests.helpers.gsm_t13 import build_t13_xlsx


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("СЭ00-00143", "143"),
        ("СЭ0В-00001", "1"),
        ("СЭ00-00438", "438"),
        ("СЭ00-00470", "470"),
        ("", None),
        ("   ", None),
        ("без-дефиса-нет", None),
        ("мусор", None),
    ],
)
def test_parse_personnel_number(raw: str, expected: str | None) -> None:
    assert parse_personnel_number(raw) == expected


def test_july_fixture_keeps_31_days_and_skips_hours() -> None:
    payload = build_t13_xlsx(
        [
            {
                "full_name": "Скрябин\nИван Петрович",
                "personnel": "СЭ00-00029",
                "default_code": "Я",
                "codes": {day: "ОТ" for day in range(1, 10)},
                "hours": "8",
            }
        ]
    )
    parsed = parse_timesheet_xlsx(payload, filename="табель от 31.07.2026 (1).xlsx")
    assert parsed.period == (date(2026, 7, 1), date(2026, 7, 31))
    assert len(parsed.rows) == 1
    row = parsed.rows[0]
    assert row.full_name == "Скрябин Иван Петрович"
    assert row.personnel_number == "29"
    by_day = dict(row.codes)
    assert len(by_day) == 31
    assert by_day[date(2026, 7, 1)] == "ОТ"
    assert by_day[date(2026, 7, 9)] == "ОТ"
    assert by_day[date(2026, 7, 10)] == "Я"
    assert by_day[date(2026, 7, 31)] == "Я"
    assert "8" not in by_day.values()


def test_unknown_code_is_stored_and_not_eligible() -> None:
    payload = build_t13_xlsx(
        [
            {
                "full_name": "Кулигин Никита",
                "personnel": "СЭ00-00143",
                "default_code": "Я",
                "codes": {3: "Б"},
            }
        ]
    )
    parsed = parse_timesheet_xlsx(payload, filename="от 31.07.2026.xlsx")
    codes = dict(parsed.rows[0].codes)
    assert codes[date(2026, 7, 3)] == "Б"
    assert "Б" not in ELIGIBLE_CODES
    assert "Я" in ELIGIBLE_CODES
    assert "В" in ELIGIBLE_CODES


def test_period_from_filename() -> None:
    assert period_from_filename("Табель от 31.07.2026 (1).xlsx") == (
        date(2026, 7, 1),
        date(2026, 7, 31),
    )
    assert period_from_filename("notes.xlsx") is None
    assert period_from_filename("от 32.07.2026.xlsx") is None


def test_unreadable_personnel_goes_to_unmatched() -> None:
    payload = build_t13_xlsx(
        [
            {
                "full_name": "Офис Без Номера",
                "personnel": "",
                "default_code": "Я",
            }
        ]
    )
    parsed = parse_timesheet_xlsx(payload, filename="от 31.07.2026.xlsx")
    assert parsed.rows == ()
    assert parsed.unmatched[0].full_name == "Офис Без Номера"


def test_header_without_tabel_is_rejected() -> None:
    payload = build_t13_xlsx(
        [{"full_name": "Кулигин Никита", "personnel": "СЭ00-00143"}],
        include_header=False,
    )
    with pytest.raises(TimesheetParseError, match="ТАБЕЛЬ"):
        parse_timesheet_xlsx(payload, filename="от 31.07.2026.xlsx")


def test_shifted_day_columns_are_rejected() -> None:
    payload = build_t13_xlsx(
        [{"full_name": "Кулигин Никита", "personnel": "СЭ00-00143"}],
        shift_day_columns=True,
    )
    with pytest.raises(TimesheetParseError, match="колонки дней"):
        parse_timesheet_xlsx(payload, filename="от 31.07.2026.xlsx")


def test_february_drops_days_past_month_end() -> None:
    payload = build_t13_xlsx(
        [
            {
                "full_name": "Кулигин Никита",
                "personnel": "СЭ00-00143",
                "default_code": "Я",
            }
        ],
        year=2026,
        month=2,
    )
    parsed = parse_timesheet_xlsx(payload, filename="от 28.02.2026.xlsx")
    days = [day for day, _code in parsed.rows[0].codes]
    assert max(days) == date(2026, 2, 28)
    assert len(days) == 28
