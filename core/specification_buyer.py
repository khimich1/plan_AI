"""Имя покупателя для печати спецификации. Договор и build_preamble не меняет."""

from __future__ import annotations

from core.supply_contract import build_preamble

_OOO_FORM = "Общество с ограниченной ответственностью"
_AO_FORM = "Акционерное общество"


def buyer_preamble_for_print(
    *,
    legal_form: str,
    full_name: str,
    signatory_name: str,
    authority_basis: str,
    signatory_position: str | None = None,
    signatory_verb: str = "действующего",
    poa_number: str | None = None,
    poa_date: str | None = None,
) -> str:
    """Преамбула покупателя для спецификации. Форма, уже стоящая в имени, не повторяется."""
    text = build_preamble(
        legal_form=legal_form,
        full_name=full_name,
        signatory_name=signatory_name,
        authority_basis=authority_basis,
        signatory_position=signatory_position,
        signatory_verb=signatory_verb,
        poa_number=poa_number,
        poa_date=poa_date,
    )
    name = (full_name or "").strip()
    if legal_form == "ooo" and name.startswith(_OOO_FORM):
        text = text.replace(f"{_OOO_FORM} «{name}»", name, 1)
    elif legal_form == "ao" and name.startswith(_AO_FORM):
        text = text.replace(f"{_AO_FORM} «{name}»", name, 1)
    return text


def buyer_signature_heading(short_name: str) -> str:
    """Краткое имя подписи. «ООО » добавляется, только если формы в имени ещё нет."""
    name = (short_name or "").strip()
    if not name:
        return ""
    if "ООО" in name or _OOO_FORM in name:
        return name
    return f"ООО {name}"
