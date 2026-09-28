"""Номер договора, преамбула и нормализация полей (SC-001)."""

from __future__ import annotations

from datetime import date

from core.supply_contract import (
    build_preamble,
    format_contract_number,
    next_contract_seq,
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
