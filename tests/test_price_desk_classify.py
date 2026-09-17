"""PD-001: classify factory price files by filename anchors."""

from __future__ import annotations

import pytest

from core.price_desk_classify import (
    ClassifyError,
    MSG_UNKNOWN_KIND,
    classify_price_filename,
    parse_price_list_date_from_filename,
)


@pytest.mark.parametrize(
    ("filename", "kind"),
    [
        ("Прайс на составные сваи от 07.09.2026.xlsx", "composite"),
        ("Расчет новых цен на ПБ 17.08.2026.xls", "plates"),
        ("расчёт новых цен на пб 01.01.2026.xls", "plates"),
        ("Расчет чего угодно ПБ.xls", "plates"),
        ("7_5 В прайс на ФБС  от 07.09.2026.xlsx", "fbs"),
        ("Прайс на лестничные ступени от 07.09.2026.xlsx", "step"),
        ("Прайс на мостовые сваи от 07.09.2026.xlsx", "bridge_pile"),
        ("Прайс на цельные сваи от 07.09.2026.xlsx", "pile"),
        ("Прайс ЛМ от 07.09.2026.xlsx", "march"),
    ],
)
def test_classify_seven_anchors(filename: str, kind: str) -> None:
    assert classify_price_filename(filename) == kind


def test_composite_wins_over_pile_substring() -> None:
    assert (
        classify_price_filename("Прайс на составные и цельные сваи.xlsx")
        == "composite"
    )


def test_case_and_yo_do_not_matter() -> None:
    assert classify_price_filename("ПРАЙС НА ФБС.XLSX") == "fbs"
    assert classify_price_filename("Расчёт новых цен на ПБ.xls") == "plates"


def test_unknown_filename_raises() -> None:
    with pytest.raises(ClassifyError, match=MSG_UNKNOWN_KIND):
        classify_price_filename("случайный_файл.xlsx")


def test_classify_uses_basename_only() -> None:
    assert (
        classify_price_filename(r"C:\прайсы\Прайс на цельные сваи от 07.09.2026.xlsx")
        == "pile"
    )


def test_parse_date_from_filename() -> None:
    assert (
        parse_price_list_date_from_filename("Прайс ЛМ от 07.09.2026.xlsx")
        == "2026-09-07"
    )
    assert parse_price_list_date_from_filename("Расчет новых цен на ПБ 17.08.2026.xls") == (
        "2026-08-17"
    )
    assert parse_price_list_date_from_filename("Прайс ЛМ.xlsx") is None
