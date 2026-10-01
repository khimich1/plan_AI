"""Номер договора поставки, преамбула и нормализация реквизитов."""

from __future__ import annotations

import copy
from datetime import date
from typing import Any

CANCELLED_STATUS = "отмена не будем работать"
DEFAULT_STATUS = "нет"
CONTRACT_STATUSES = (
    DEFAULT_STATUS,
    "подписан по ЭДО",
    "оригинал в бухгалтерии",
    "подписан с синей печатью оригинал",
    CANCELLED_STATUS,
)
SCAN_NOTES = ("получен скан", "внесены правки")
DEFAULT_SIGNATORY_VERB = "действующего"
SIGNATORY_VERBS = frozenset({"действующего", "действующей"})

LEGAL_FORMS = ("ooo", "ao", "ip", "kfh", "person")
LEGAL_FORMS_REQUIRING_KPP = frozenset({"ooo", "ao"})
LEGAL_FORMS_REQUIRING_POSITION = frozenset({"ooo", "ao"})

MSG_BIND_COUNTERPARTY = "Сначала занесите контрагента из 1С"
MSG_NOT_ARCHIVED = "Договор можно оформить только у КП в статусе «в архиве»"
MSG_NO_ACTIVE_CONTRACT = "Нет действующего договора поставки. Документ не собран."
KP_DOCUMENT = "КоммерческоеПредложение"
INVOICE_DOCUMENT = "Счёт на оплату"
STAMP_DOCUMENTS = frozenset({KP_DOCUMENT, INVOICE_DOCUMENT})


class MissingSupplyContractError(Exception):
    """Нет действующего договора — шапка JSON не возвращается."""


def active_status(status: str) -> bool:
    return status != CANCELLED_STATUS


def duplicate_contract_message(number: str) -> str:
    return f"У контрагента уже есть договор {number}"


def format_contract_number(seq: int, on: date) -> str:
    return f"{seq:04d}/{on.month:02d}/{on.year % 100:02d}"


def contract_seq(number: str) -> int | None:
    head = (number or "").split("/", 1)[0].strip()
    if not head.isdigit():
        return None
    return int(head)


def next_contract_seq(occupied_numbers: list[str]) -> int:
    """Следующий seq: max первого сегмента по всем номерам, включая отменённые."""
    highest = 0
    for number in occupied_numbers:
        seq = contract_seq(number)
        if seq is not None and seq > highest:
            highest = seq
    return highest + 1


def normalize_digits(value: str | None) -> str:
    if not value:
        return ""
    return "".join(ch for ch in value if ch.isdigit())


def normalize_email(value: str | None) -> str:
    if not value:
        return ""
    return "".join(value.split())


def _basis_phrase(
    authority_basis: str,
    *,
    poa_number: str | None,
    poa_date: str | None,
) -> str:
    if authority_basis == "доверенность":
        number = (poa_number or "").strip()
        issued = (poa_date or "").strip()
        tail = f" № {number}" if number else ""
        if issued:
            tail = f"{tail} от {issued}"
        return f"доверенности{tail}"
    return "Устава"


def _party_phrase(
    legal_form: str,
    full_name: str,
    signatory_verb: str,
    party_title: str = "Покупатель",
) -> str:
    name = full_name.strip()
    titled = f"«{party_title}»"
    if legal_form == "ooo":
        return (
            f"Общество с ограниченной ответственностью «{name}», "
            f"именуемое в дальнейшем {titled}"
        )
    if legal_form == "ao":
        return (
            f"Акционерное общество «{name}», "
            f"именуемое в дальнейшем {titled}"
        )
    named = "именуемая" if signatory_verb == "действующей" else "именуемый"
    if legal_form == "ip":
        return f"Индивидуальный предприниматель {name}, {named} в дальнейшем {titled}"
    if legal_form == "kfh":
        return (
            f"Крестьянское (фермерское) хозяйство «{name}», "
            f"именуемое в дальнейшем {titled}"
        )
    return f"{name}, {named} в дальнейшем {titled}"


def build_preamble(
    *,
    legal_form: str,
    full_name: str,
    signatory_name: str,
    authority_basis: str,
    signatory_position: str | None = None,
    signatory_verb: str = DEFAULT_SIGNATORY_VERB,
    poa_number: str | None = None,
    poa_date: str | None = None,
    party_title: str = "Покупатель",
) -> str:
    """Текст преамбулы. Род глагола берётся только из signatory_verb."""
    verb = signatory_verb or DEFAULT_SIGNATORY_VERB
    party = _party_phrase(legal_form, full_name, verb, party_title)
    basis = _basis_phrase(authority_basis, poa_number=poa_number, poa_date=poa_date)
    if legal_form in LEGAL_FORMS_REQUIRING_POSITION:
        position = (signatory_position or "").strip()
        signatory = signatory_name.strip()
        face = f"{position} {signatory}".strip()
        return f"{party}, в лице {face}, {verb} на основании {basis}"
    return f"{party}, {verb} на основании {basis}"


def attach_contract_number(
    document: dict[str, Any],
    counterparty_id: int,
    active_contract: dict[str, Any] | None,
) -> dict[str, Any]:
    """Ставит НомерДоговора в шапку уже собранного документа, рядом с контрагентом.

    Номер берётся только из действующего договора этого контрагента.
    Нет договора — ошибка, документ не возвращается. Живой отправки нет.
    """
    number = _active_contract_number(counterparty_id, active_contract)
    if document.get("Документ") not in STAMP_DOCUMENTS:
        raise MissingSupplyContractError(
            "Ожидался документ «КоммерческоеПредложение» или «Счёт на оплату»"
        )
    return _stamp_header(document, number)


def _active_contract_number(
    counterparty_id: int,
    active_contract: dict[str, Any] | None,
) -> str:
    if active_contract is None:
        raise MissingSupplyContractError(MSG_NO_ACTIVE_CONTRACT)
    owner = active_contract.get("counterparty_id")
    if owner is None or int(owner) != int(counterparty_id):
        raise MissingSupplyContractError(MSG_NO_ACTIVE_CONTRACT)
    if not active_status(str(active_contract.get("status") or "")):
        raise MissingSupplyContractError(MSG_NO_ACTIVE_CONTRACT)
    number = str(active_contract.get("number") or "").strip()
    if not number:
        raise MissingSupplyContractError(MSG_NO_ACTIVE_CONTRACT)
    return number


def _stamp_header(document: dict[str, Any], number: str) -> dict[str, Any]:
    stamped: dict[str, Any] = {}
    placed = False
    for key, value in document.items():
        if key == "НомерДоговора":
            continue
        if key == "Товары":
            stamped[key] = copy.deepcopy(value)
            continue
        stamped[key] = copy.deepcopy(value) if isinstance(value, (dict, list)) else value
        if key == "Контрагент":
            stamped["НомерДоговора"] = number
            placed = True
    if placed:
        return stamped
    goods = stamped.pop("Товары", None)
    stamped["НомерДоговора"] = number
    if goods is not None:
        stamped["Товары"] = goods
    return stamped
