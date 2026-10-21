"""Выбор видимого имени карточки по справочнику 1С. Живую базу не открывает."""

from __future__ import annotations

from app.services.contract_card_directory_name import resolve_card_display_name


def test_inn_named_row_is_dropped_when_another_name_shares_inn() -> None:
    renamed = resolve_card_display_name(
        "УНИВЕРСАЛСТРОЙ",
        "7606117480",
        ["7606117480", "УНИВЕРСАЛСТРОЙ ООО"],
    )
    assert renamed == "УНИВЕРСАЛСТРОЙ ООО"


def test_unknown_inn_keeps_the_card_name() -> None:
    kept = resolve_card_display_name('ООО "Прогресс"', "6700015814", [])
    assert kept == 'ООО "Прогресс"'


def test_two_different_directory_names_keep_the_card_name() -> None:
    kept = resolve_card_display_name(
        "Банк Точка",
        "7743298799",
        ["АРСЕНАЛ ООО", 'Обособленное подразделение "Арсенал"'],
    )
    assert kept == "Банк Точка"


def test_longer_blank_name_on_0001_is_not_replaced_by_shorter_directory_name() -> None:
    blank = "Общество с ограниченной ответственностью «КАПИТАЛСТРОЙ»"
    kept = resolve_card_display_name(
        blank,
        "7612045875",
        ["КАПИТАЛСТРОЙ ООО"],
        number="0001/10/26",
        imported_name=None,
        full_name=blank,
    )
    assert kept == blank
