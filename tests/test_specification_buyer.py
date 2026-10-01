"""Имя покупателя для печати спецификации. Договор и build_preamble не меняются."""

from __future__ import annotations

from core.specification_buyer import buyer_preamble_for_print, buyer_signature_heading
from core.supply_contract import build_preamble

_MD = "Общество с ограниченной ответственностью «Специализированный застройщик «МД-Строй»"
_AO = "Акционерное общество «Север»"


def test_full_name_that_already_starts_with_ooo_is_not_wrapped_again() -> None:
    text = buyer_preamble_for_print(
        legal_form="ooo",
        full_name=_MD,
        signatory_name="Либинзон Дмитрий Евгеньевич",
        signatory_position="Директор",
        authority_basis="устав",
    )

    assert text.count("Общество с ограниченной ответственностью") == 1
    assert _MD in text
    assert f"«{_MD}»" not in text


def test_bare_etalon_is_wrapped_once() -> None:
    text = buyer_preamble_for_print(
        legal_form="ooo",
        full_name="ЭТАЛОН",
        signatory_name="Головченко Андрей Анатольевич",
        signatory_position="Генеральный директор",
        authority_basis="устав",
    )

    assert text.count("Общество с ограниченной ответственностью") == 1
    assert "Общество с ограниченной ответственностью «ЭТАЛОН»" in text


def test_joint_stock_name_that_already_has_the_form_is_not_wrapped_again() -> None:
    text = buyer_preamble_for_print(
        legal_form="ao",
        full_name=_AO,
        signatory_name="Петров Пётр Петрович",
        signatory_position="Директор",
        authority_basis="устав",
    )

    assert text.count("Акционерное общество") == 1
    assert _AO in text


def test_other_legal_forms_match_build_preamble() -> None:
    kwargs = dict(
        legal_form="ip",
        full_name="Иванов Иван Иванович",
        signatory_name="Иванов Иван Иванович",
        authority_basis="устав",
        signatory_verb="действующего",
    )

    assert buyer_preamble_for_print(**kwargs) == build_preamble(**kwargs)


def test_position_and_person_stay_in_the_nominative() -> None:
    text = buyer_preamble_for_print(
        legal_form="ooo",
        full_name="Ромашка",
        signatory_name="Либинзон Дмитрий Евгеньевич",
        signatory_position="Директор",
        authority_basis="устав",
    )

    assert "в лице Директор Либинзон Дмитрий Евгеньевич" in text
    assert "Директора" not in text
    assert "Либинзона" not in text
    assert text.count("в лице") == 1


def test_signature_prefix_is_added_only_when_the_short_name_has_no_ooo() -> None:
    assert buyer_signature_heading("ЭТАЛОН") == "ООО ЭТАЛОН"
    assert buyer_signature_heading("Ромашка") == "ООО Ромашка"
    assert buyer_signature_heading("ООО «ЭТАЛОН»") == "ООО «ЭТАЛОН»"
    assert buyer_signature_heading("ООО ЭТАЛОН") == "ООО ЭТАЛОН"
    assert buyer_signature_heading(_MD) == _MD
    assert buyer_signature_heading("  ") == ""
