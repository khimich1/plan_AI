"""Какой свободный договор связывается с каким клиентом. Базу не открывает."""

from __future__ import annotations

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

    Пара строится, только если на ИНН ровно один свободный неотменённый
    договор и ровно один активный клиент, и у этого клиента ещё нет
    другого действующего договора. Два договора и больше дают пусто.
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
        (int(group[0]["id"]), eligible[inn])
        for inn, group in grouped.items()
        if len(group) == 1
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


def _cancelled(contract: ContractRow) -> bool:
    return str(contract.get("status") or "") == CANCELLED_STATUS


def _active_client(client: ClientRow) -> bool:
    return _flag(client.get("is_active")) and _flag(client.get("is_client"))


def _flag(value: object) -> bool:
    try:
        return int(value or 0) != 0
    except (TypeError, ValueError):
        return False
