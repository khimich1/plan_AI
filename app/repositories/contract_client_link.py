"""Какой свободный договор связывается с каким клиентом. Базу не открывает."""

from __future__ import annotations

from datetime import date, datetime
from typing import Mapping

from core.supply_contract import CANCELLED_STATUS

ContractRow = Mapping[str, object]
ClientRow = Mapping[str, object]


def normalize_inn(value: object) -> str | None:
    if value is None:
        return None
    text = "".join(str(value).split())
    return text or None


def select_contract_client_links(
    contracts: list[ContractRow],
    clients: list[ClientRow],
) -> list[tuple[int, int]]:
    """Пары (id договора, id клиента) для ещё пустого ``counterparty_id``.

    ИНН годится, только если в справочнике ровно одна строка и она —
    активный клиент. Из свободных неотменённых договоров этого ИНН
    берётся поздняя ``contract_date``, при равенстве — больший ``id``.
    У клиента уже есть другой действующий договор — пара не строится.
    """
    eligible = _eligible_clients_by_inn(clients)
    busy = _clients_with_active_contract(contracts)
    grouped: dict[str, list[ContractRow]] = {}
    for contract in contracts:
        inn = _candidate_inn(contract, eligible, busy)
        if inn is None:
            continue
        grouped.setdefault(inn, []).append(contract)
    pairs = [
        (int(_latest_contract(group)["id"]), eligible[inn])
        for inn, group in grouped.items()
    ]
    pairs.sort(key=lambda item: item[0])
    return pairs


def unique_client_id_for_inn(clients: list[ClientRow], inn: object) -> int | None:
    key = normalize_inn(inn)
    if key is None:
        return None
    return _eligible_clients_by_inn(clients).get(key)


def _eligible_clients_by_inn(clients: list[ClientRow]) -> dict[str, int]:
    rows_by_inn: dict[str, list[ClientRow]] = {}
    for client in clients:
        inn = normalize_inn(client.get("inn"))
        if inn is None:
            continue
        rows_by_inn.setdefault(inn, []).append(client)
    eligible: dict[str, int] = {}
    for inn, rows in rows_by_inn.items():
        if len(rows) != 1 or not _active_client(rows[0]):
            continue
        eligible[inn] = int(rows[0]["id"])
    return eligible


def _clients_with_active_contract(contracts: list[ContractRow]) -> set[int]:
    busy: set[int] = set()
    for contract in contracts:
        if _cancelled(contract):
            continue
        owner = contract.get("counterparty_id")
        if owner is None:
            continue
        busy.add(int(owner))
    return busy


def _candidate_inn(
    contract: ContractRow,
    eligible: dict[str, int],
    busy: set[int],
) -> str | None:
    if contract.get("counterparty_id") is not None or _cancelled(contract):
        return None
    inn = normalize_inn(contract.get("inn"))
    if inn is None or inn not in eligible:
        return None
    if eligible[inn] in busy:
        return None
    if contract.get("id") is None:
        return None
    return inn


def _latest_contract(group: list[ContractRow]) -> ContractRow:
    return max(group, key=_contract_rank)


def _contract_rank(contract: ContractRow) -> tuple[date, int]:
    return (_contract_day(contract.get("contract_date")), int(contract["id"]))


def _contract_day(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if len(text) < 10 or text[4] != "-" or text[7] != "-":
        return date.min
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        return date.min


def _cancelled(contract: ContractRow) -> bool:
    return str(contract.get("status") or "") == CANCELLED_STATUS


def _active_client(client: ClientRow) -> bool:
    return _flag(client.get("is_active")) and _flag(client.get("is_client"))


def _flag(value: object) -> bool:
    try:
        return int(value or 0) != 0
    except (TypeError, ValueError):
        return False
