"""Номер договора, преамбула и нормализация полей (SC-001, CDC-001)."""

from __future__ import annotations

from datetime import date

import pytest

from core.supply_contract import (
    build_preamble,
    contract_date_from_number,
    format_contract_number,
    next_contract_seq,
    normalize_contract_number,
    normalize_digits,
    normalize_email,
)


def test_next_seq_after_1027_uses_contract_date() -> None:
    seq = next_contract_seq(["1027/08/26"])
    assert format_contract_number(seq, date(2026, 9, 25)) == "1028/09/26"


def test_cancelled_number_stays_occupied() -> None:
    occupied = ["1026/09/26", "1027/09/26"]
    assert next_contract_seq(occupied) == 1028
    assert next_contract_seq(occupied) == 1028


def test_holes_are_not_reused() -> None:
    assert next_contract_seq(["0001/01/26", "0005/03/26"]) == 6


def test_empty_registry_starts_at_0001() -> None:
    assert format_contract_number(next_contract_seq([]), date(2026, 1, 2)) == "0001/01/26"


def test_ooo_preamble_uses_face_and_basis() -> None:
    text = build_preamble(
        legal_form="ooo",
        full_name="Ромашка",
        signatory_position="Генерального директора",
        signatory_name="Иванова Мария Петровна",
        signatory_verb="действующего",
        authority_basis="устав",
    )
    assert "Общество с ограниченной ответственностью" in text
    assert "в лице" in text
    assert "на основании" in text
    assert "действующего" in text
    assert "действующей" not in text


def test_ao_preamble_uses_face_and_basis() -> None:
    text = build_preamble(
        legal_form="ao",
        full_name="Север",
        signatory_position="Директора",
        signatory_name="Петров Пётр",
        authority_basis="доверенность",
        poa_number="14",
        poa_date="01.03.2026",
    )
    assert "Акционерное общество" in text
    assert "в лице" in text
    assert "на основании" in text
    assert "действующего" in text
    assert "14" in text


def test_ip_kfh_and_person_skip_ooo_formula() -> None:
    common = {
        "signatory_name": "Сидоров Сидор",
        "signatory_verb": "действующего",
        "authority_basis": "устав",
    }
    ip_text = build_preamble(legal_form="ip", full_name="Сидоров Сидор Сидорович", **common)
    kfh_text = build_preamble(legal_form="kfh", full_name="Заря", **common)
    person_text = build_preamble(
        legal_form="person",
        full_name="Иванова Анна",
        signatory_name="Иванова Анна",
        signatory_verb="действующего",
        authority_basis="устав",
    )
    for text in (ip_text, kfh_text, person_text):
        assert "Общество с ограниченной ответственностью" not in text
        assert "в лице" not in text
    assert "действующего" in person_text
    assert "действующей" not in person_text


def test_verb_default_is_masculine_participle() -> None:
    text = build_preamble(
        legal_form="ip",
        full_name="Смирнова Ольга",
        signatory_name="Смирнова Ольга",
        authority_basis="устав",
    )
    assert "действующего" in text
    assert "действующей" not in text


def test_explicit_feminine_verb_is_kept() -> None:
    text = build_preamble(
        legal_form="person",
        full_name="Смирнов Олег",
        signatory_name="Смирнов Олег",
        signatory_verb="действующей",
        authority_basis="устав",
    )
    assert "действующей" in text


def test_spaced_inn_and_email_are_glued() -> None:
    assert normalize_digits("760 4 01001") == "760401001"
    assert normalize_email("is-ag @mail.ru") == "is-ag@mail.ru"


def test_normalize_contract_number_drops_whitespace() -> None:
    assert normalize_contract_number("  0786 / 1 / 25  ") == "0786/1/25"
    assert normalize_contract_number("0844-06-25") == "0844-06-25"
    assert normalize_contract_number("966/06/26КП") == "966/06/26КП"
    assert normalize_contract_number("ТБ-3/08/26Б") == "ТБ-3/08/26Б"


def test_normalize_contract_number_rejects_empty_and_no_digit() -> None:
    message = "В номере договора нужна хотя бы одна цифра"
    with pytest.raises(ValueError, match=message):
        normalize_contract_number("")
    with pytest.raises(ValueError, match=message):
        normalize_contract_number("   ")
    with pytest.raises(ValueError, match=message):
        normalize_contract_number("без цифр")


def test_normalize_contract_number_rejects_longer_than_32() -> None:
    assert normalize_contract_number("1" * 32) == "1" * 32
    with pytest.raises(ValueError, match="Номер договора длиннее 32 символов"):
        normalize_contract_number("1" * 33)


@pytest.mark.parametrize(
    ("number", "expected"),
    [
        ("0785/01/25", date(2025, 1, 1)),
        ("0786/1/25", date(2025, 1, 1)),
        ("0844-06-25", date(2025, 6, 1)),
        ("966/06/26КП", date(2026, 6, 1)),
        ("ТБ-1/10/25", date(2025, 10, 1)),
        ("ТБ-3/08/26Б", date(2026, 8, 1)),
    ],
)
def test_contract_date_from_number_uses_month_year_and_day_01(
    number: str,
    expected: date,
) -> None:
    assert contract_date_from_number(number) == expected


def test_contract_date_from_number_without_month_and_year_is_none() -> None:
    assert contract_date_from_number("1033") is None


@pytest.mark.parametrize(
    ("number", "expected"),
    [
        ("1/01/00", date(2000, 1, 1)),
        ("1/12/69", date(2069, 12, 1)),
        ("1/01/70", date(1970, 1, 1)),
        ("1/06/99", date(1999, 6, 1)),
    ],
)
def test_contract_date_year_window(number: str, expected: date) -> None:
    assert contract_date_from_number(number) == expected


def test_contract_date_from_number_rejects_unparsed_tail() -> None:
    assert contract_date_from_number("1/13/25") is None
    assert contract_date_from_number("1/00/25") is None
    assert contract_date_from_number("0785/01/2025") is None


def test_loaded_prefixes_do_not_change_next_seq() -> None:
    occupied = ["1033/09/26", "ТБ-1/10/25", "0844-06-25"]
    assert next_contract_seq(occupied) == 1034
