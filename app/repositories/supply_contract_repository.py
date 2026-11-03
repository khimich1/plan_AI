"""SQL-доступ к договору поставки. Номер выдаётся в одной транзакции."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from app.repositories.contract_client_link import (
    normalize_inn,
    select_contract_client_links,
    unique_client_id_for_inn,
)

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.supply_contract import (
    CANCELLED_STATUS,
    CONTRACT_STATUSES,
    DEFAULT_STATUS,
    SCAN_NOTES,
    contract_date_from_number,
    format_contract_number,
    next_contract_seq,
    normalize_contract_number,
)

_INSERT_COLUMNS = (
    "counterparty_id",
    "number",
    "contract_date",
    "manager_name",
    "status",
    "legal_form",
    "full_name",
    "short_name",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "inn",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "okved",
    "created_at",
    "created_by_user_id",
)

# Поиск фильтрует эти строки в сервисе. Здесь только колонки:
# ключи имени, ИНН и номера плюс поля, которые отдаёт карточка.
_CARD_SEARCH_COLUMNS = (
    "id",
    "imported_name",
    "full_name",
    "short_name",
    "inn",
    "number",
    "counterparty_id",
    "contract_date",
    "manager_name",
    "status",
    "scan_note",
    "scan_path",
    "legal_form",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "okved",
)

# Поля бланка. Номер сюда не входит. Дата и менеджер пишутся отдельно:
# дату — только если её передали, менеджера — только в пустую строку.
_REQUISITE_COLUMNS = (
    "legal_form",
    "full_name",
    "short_name",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "inn",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "okved",
)


def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    return dict(row)


def _blank(value: object) -> bool:
    if value is None:
        return True
    return isinstance(value, str) and not value.strip()


def _requisite_assignments(
    fields: dict[str, Any],
    *,
    contract_date: date | None,
    current_manager: object,
    manager_name: str,
) -> tuple[list[str], list[Any]]:
    assignments: list[str] = []
    values: list[Any] = []
    for column in _REQUISITE_COLUMNS:
        if column not in fields:
            continue
        assignments.append(f"{column} = ?")
        values.append(fields[column])
    if contract_date is not None:
        assignments.append("contract_date = ?")
        values.append(contract_date.isoformat())
    if _blank(current_manager) and not _blank(manager_name):
        assignments.append("manager_name = ?")
        values.append(manager_name.strip())
    return assignments, values


@dataclass(frozen=True)
class ContractLinkStats:
    """Сколько договоров получили связь и сколько свободных осталось."""

    created: int
    skipped: int


@dataclass(frozen=True)
class DirectoryWrite:
    """Договор клиента либо отказ. Строки справочника этот вызов не создаёт."""

    row: dict[str, Any] | None = None
    refusal: str | None = None
    hint: str | None = None


def _card_inn_matches(card: dict[str, Any], party: dict[str, Any]) -> bool:
    """Пустой ИНН карточки не мешает выбору. Заполненный должен совпасть с клиентом."""
    card_inn = normalize_inn(card.get("inn"))
    if card_inn is None:
        return True
    return card_inn == normalize_inn(party.get("inn"))


def _directory_refusal(party: dict[str, Any] | None, *, clients_only: bool) -> str | None:
    if party is None:
        return "missing"
    if not int(party.get("is_active") or 0):
        return "inactive"
    if clients_only and not int(party.get("is_client") or 0):
        return "missing"
    return None


_LINK_UPDATE = """
UPDATE supply_contract
SET counterparty_id = ?
WHERE id = ?
  AND counterparty_id IS NULL
  AND status != ?
  AND NOT EXISTS (
      SELECT 1 FROM supply_contract AS other
      WHERE other.counterparty_id = ?
        AND other.status != ?
        AND other.id != supply_contract.id
  )
"""


class SupplyContractRepository:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def _connect(self) -> sqlite3.Connection:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.isolation_level = None
        return conn

    def get_active_by_counterparty(self, counterparty_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM supply_contract
                WHERE counterparty_id = ? AND status != ?
                """,
                (int(counterparty_id), CANCELLED_STATUS),
            ).fetchone()
        return _row_to_dict(row)

    def get_by_id(self, contract_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM supply_contract WHERE id = ?",
                (int(contract_id),),
            ).fetchone()
        return _row_to_dict(row)

    def get_by_number(self, number: str) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM supply_contract WHERE number = ?",
                (number,),
            ).fetchone()
        return _row_to_dict(row)

    def insert_card(
        self,
        *,
        imported_name: str,
        number: str,
        user_id: int,
        inn: str | None = None,
        contract_date: date | None = None,
    ) -> dict[str, Any]:
        """Неполная карточка: без контрагента и без реквизитов бланка.

        Занятый номер поднимает IntegrityError как есть, без текста реестра.
        """
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            contract_id = self._insert_card_row(
                conn,
                imported_name=imported_name,
                inn=inn,
                number=number,
                contract_date=contract_date,
                user_id=user_id,
            )
            conn.execute("COMMIT")
            stored = self._by_id(conn, contract_id)
            if stored is None:
                raise sqlite3.OperationalError("карточка не прочитана после вставки")
            return stored
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def insert_missing_cards(
        self,
        rows: list[dict[str, Any]],
        *,
        user_id: int,
    ) -> tuple[int, int]:
        """Вставляет карточки, которых ещё нет. Занятый номер пропускает.

        Одна транзакция: ошибка откатывает все вставки этого вызова.
        Имя пишется только в imported_name.
        """
        if not rows:
            return 0, 0
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            taken = {
                str(row["number"])
                for row in conn.execute("SELECT number FROM supply_contract")
            }
            created, skipped = self._insert_missing(conn, rows, taken, user_id)
            conn.execute("COMMIT")
            return created, skipped
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def list_for_search(self) -> list[dict[str, Any]]:
        """Все строки договора колонками карточки. Подстроку режет сервис."""
        columns = ", ".join(_CARD_SEARCH_COLUMNS)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT {columns} FROM supply_contract ORDER BY id"
            ).fetchall()
        return [dict(row) for row in rows]

    def create_active(
        self,
        fields: dict[str, Any],
        *,
        counterparty_id: int,
        manager_name: str,
        contract_date: date,
        user_id: int,
    ) -> tuple[dict[str, Any], bool]:
        """Возвращает договор и флаг «строка только что создана».

        Повтор при действующем договоре не вставляет строку и не берёт новый seq.
        """
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            existing = self._active_row(conn, counterparty_id)
            if existing is not None:
                conn.execute("COMMIT")
                return existing, False
            number = self._allocate_number(conn, contract_date)
            created_at = datetime.now().isoformat(timespec="seconds")
            row = {
                **fields,
                "counterparty_id": int(counterparty_id),
                "number": number,
                "contract_date": contract_date.isoformat(),
                "manager_name": manager_name,
                "status": DEFAULT_STATUS,
                "created_at": created_at,
                "created_by_user_id": int(user_id),
            }
            placeholders = ", ".join("?" for _ in _INSERT_COLUMNS)
            columns = ", ".join(_INSERT_COLUMNS)
            try:
                cur = conn.execute(
                    f"INSERT INTO supply_contract ({columns}) VALUES ({placeholders})",
                    tuple(row.get(column) for column in _INSERT_COLUMNS),
                )
            except sqlite3.IntegrityError:
                conn.execute("ROLLBACK")
                winner = self.get_active_by_counterparty(counterparty_id)
                if winner is not None:
                    return winner, False
                raise
            contract_id = int(cur.lastrowid)
            self._insert_event(
                conn,
                contract_id=contract_id,
                field="status",
                old_value=None,
                new_value=DEFAULT_STATUS,
                user_id=user_id,
                created_at=created_at,
            )
            conn.execute("COMMIT")
            stored = self._by_id(conn, contract_id)
            if stored is None:
                raise sqlite3.OperationalError("договор не прочитан после вставки")
            return stored, True
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def insert_imported(
        self,
        rows: list[dict[str, Any]],
        *,
        user_id: int,
    ) -> list[dict[str, Any]]:
        """Строки листа без counterparty_id. Номер занимает seq."""
        if not rows:
            return []
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            created_at = datetime.now().isoformat(timespec="seconds")
            stored: list[dict[str, Any]] = []
            for row in rows:
                contract_id = self._insert_imported_row(
                    conn,
                    row,
                    user_id=user_id,
                    created_at=created_at,
                )
                saved = self._by_id(conn, contract_id)
                if saved is None:
                    raise sqlite3.OperationalError("импортированный договор не прочитан")
                stored.append(saved)
            conn.execute("COMMIT")
            return stored
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def link_counterparty(
        self,
        contract_id: int,
        counterparty_id: int,
        *,
        user_id: int,
    ) -> dict[str, Any] | None:
        """Явная связка. Похожее имя само сюда не попадает."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            party = conn.execute(
                "SELECT id, name FROM counterparties WHERE id = ?",
                (int(counterparty_id),),
            ).fetchone()
            if party is None:
                raise ValueError("counterparty")
            if current.get("counterparty_id") not in (None, int(counterparty_id)):
                raise ValueError("already-linked")
            active = self._active_row(conn, int(counterparty_id))
            if active is not None and int(active["id"]) != int(contract_id):
                raise ValueError(str(active["number"]))
            if current.get("counterparty_id") == int(counterparty_id):
                conn.execute("COMMIT")
                return current
            created_at = datetime.now().isoformat(timespec="seconds")
            conn.execute(
                "UPDATE supply_contract SET counterparty_id = ? WHERE id = ?",
                (int(counterparty_id), int(contract_id)),
            )
            self._insert_event(
                conn,
                contract_id=int(contract_id),
                field="counterparty_id",
                old_value=None if current.get("counterparty_id") is None else str(current["counterparty_id"]),
                new_value=str(int(counterparty_id)),
                user_id=user_id,
                created_at=created_at,
            )
            conn.execute("COMMIT")
            return self._by_id(conn, contract_id)
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def link_unique_inn_contracts(self, *, user_id: int | None = None) -> ContractLinkStats:
        """Пишет ``counterparty_id`` выбранным договорам. Справочник не меняет.

        Повторный вызов ничего не добавляет: уже заполненный id и второй
        действующий договор того же клиента условие UPDATE отсекает.
        """
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            contracts, clients = self._fetch_link_inputs(conn)
            client_count = len(clients)
            pairs = select_contract_client_links(contracts, clients)
            created = self._apply_unique_links(conn, pairs, user_id=user_id)
            if self._count_clients(conn) != client_count:
                raise sqlite3.OperationalError("справочник контрагентов изменился")
            conn.execute("COMMIT")
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()
        unlinked = sum(1 for row in contracts if row.get("counterparty_id") is None)
        return ContractLinkStats(created=created, skipped=unlinked - created)

    def planned_contract_client_links(self) -> list[tuple[int, int]]:
        """Та же пара, что и запись, но без UPDATE."""
        with self._connect() as conn:
            contracts, clients = self._fetch_link_inputs(conn)
        return select_contract_client_links(contracts, clients)

    def active_contract_refs(self, client_ids: list[int]) -> dict[int, dict[str, Any]]:
        """Действующий договор каждого клиента: id и номер. Отменённый не входит."""
        if not client_ids:
            return {}
        marks = ", ".join("?" for _ in client_ids)
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT id, number, counterparty_id
                FROM supply_contract
                WHERE counterparty_id IN ({marks})
                  AND status != ?
                """,
                (*[int(item) for item in client_ids], CANCELLED_STATUS),
            ).fetchall()
        return {
            int(row["counterparty_id"]): {"id": int(row["id"]), "number": str(row["number"])}
            for row in rows
        }

    def linkable_contract_refs(self) -> dict[int, dict[str, Any]]:
        """Одна свободная карточка на клиента. Несколько карточек и два кода дают пусто."""
        with self._connect() as conn:
            contracts, clients = self._fetch_link_inputs(conn)
        numbers = {
            int(row["id"]): str(row["number"])
            for row in contracts
            if row.get("id") is not None and row.get("number")
        }
        return {
            int(client_id): {"id": int(contract_id), "number": numbers[int(contract_id)]}
            for contract_id, client_id in select_contract_client_links(contracts, clients)
            if int(contract_id) in numbers
        }

    def free_contract_refs_by_inn(self) -> dict[str, list[dict[str, Any]]]:
        """Свободные неотменённые карточки, сгруппированные по ИНН без пробелов."""
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, number, inn
                FROM supply_contract
                WHERE counterparty_id IS NULL
                  AND status != ?
                ORDER BY id
                """,
                (CANCELLED_STATUS,),
            ).fetchall()
        grouped: dict[str, list[dict[str, Any]]] = {}
        for row in rows:
            inn = normalize_inn(row["inn"])
            if inn is None:
                continue
            grouped.setdefault(inn, []).append(
                {"id": int(row["id"]), "number": str(row["number"])}
            )
        return grouped

    def ensure_directory_contract(
        self,
        counterparty_id: int,
        *,
        user_id: int,
        clients_only: bool,
        number: str | None = None,
        contract_id: int | None = None,
    ) -> DirectoryWrite:
        """Привязать выбранный номер или карточку. Пустой вызов строку не создаёт.

        Новый текст пишется отдельной вставкой, без ``_allocate_number``.
        """
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            party_count = self._count_clients(conn)
            party = self._directory_party(conn, counterparty_id)
            refusal = _directory_refusal(party, clients_only=clients_only)
            if refusal is not None or party is None:
                conn.execute("COMMIT")
                return DirectoryWrite(refusal=refusal or "missing")
            outcome = self._link_directory_choice(
                conn,
                party,
                user_id=user_id,
                number=number,
                contract_id=contract_id,
            )
            if outcome.refusal is not None:
                conn.execute("COMMIT")
                return outcome
            if self._count_clients(conn) != party_count:
                raise sqlite3.OperationalError("справочник контрагентов изменился")
            conn.execute("COMMIT")
            return outcome
        except sqlite3.IntegrityError:
            self._rollback_quietly(conn)
            return self._directory_conflict(counterparty_id, number=number, contract_id=contract_id)
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def _directory_conflict(
        self,
        counterparty_id: int,
        *,
        number: str | None,
        contract_id: int | None,
    ) -> DirectoryWrite:
        winner = self.get_active_by_counterparty(int(counterparty_id))
        if winner is not None and contract_id is not None and int(winner["id"]) == int(contract_id):
            return DirectoryWrite(row=winner)
        if winner is not None and number:
            try:
                if str(winner["number"]) == normalize_contract_number(number):
                    return DirectoryWrite(row=winner)
            except ValueError:
                pass
        if winner is not None:
            return DirectoryWrite(refusal="busy", hint=str(winner["number"]))
        raise sqlite3.IntegrityError("номер договора не записан")

    def save_directory_blank(
        self,
        counterparty_id: int,
        fields: dict[str, Any],
        *,
        user_id: int,
        manager_name: str,
        clients_only: bool,
        number_date: date,
        contract_date: date,
        update_date: date | None,
    ) -> DirectoryWrite:
        """Первый готовый бланк берёт следующий seq. Повтор пишет реквизиты в ту же строку."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            party_count = self._count_clients(conn)
            party = self._directory_party(conn, counterparty_id)
            refusal = _directory_refusal(party, clients_only=clients_only)
            if refusal is not None or party is None:
                conn.execute("COMMIT")
                return DirectoryWrite(refusal=refusal or "missing")
            active = self._active_row(conn, int(party["id"]))
            if active is not None:
                stored = self._write_blank_requisites(
                    conn,
                    active,
                    fields,
                    manager_name=manager_name,
                    contract_date=update_date,
                )
            else:
                stored = self._insert_directory_blank(
                    conn,
                    party,
                    fields,
                    user_id=user_id,
                    manager_name=manager_name,
                    number_date=number_date,
                    contract_date=contract_date,
                )
            if self._count_clients(conn) != party_count:
                raise sqlite3.OperationalError("справочник контрагентов изменился")
            conn.execute("COMMIT")
            return DirectoryWrite(row=stored)
        except sqlite3.IntegrityError:
            self._rollback_quietly(conn)
            winner = self.get_active_by_counterparty(int(counterparty_id))
            if winner is None:
                raise
            updated = self.save_requisites(
                int(winner["id"]),
                fields,
                manager_name=manager_name,
                contract_date=update_date,
            )
            if updated is None:
                raise sqlite3.OperationalError("карточка не прочитана после записи реквизитов")
            return DirectoryWrite(row=updated)
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def unique_active_client_id(self, inn: object) -> int | None:
        with self._connect() as conn:
            _contracts, clients = self._fetch_link_inputs(conn)
        return unique_client_id_for_inn(clients, inn)

    def list_counterparty_names(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, name, name_normalized
                FROM counterparties
                WHERE is_active = 1
                ORDER BY id
                """
            ).fetchall()
        return [dict(row) for row in rows]

    def list_registry(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    c.id,
                    c.number,
                    c.contract_date,
                    c.manager_name,
                    c.status,
                    c.scan_note,
                    c.scan_path,
                    c.counterparty_id,
                    COALESCE(NULLIF(cp.name, ''), c.imported_name, '') AS counterparty_name
                FROM supply_contract c
                LEFT JOIN counterparties cp ON cp.id = c.counterparty_id
                ORDER BY c.id
                """
            ).fetchall()
        items: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            item["has_scan"] = bool(item.pop("scan_path", None))
            items.append(item)
        return items

    def apply_update(
        self,
        contract_id: int,
        *,
        user_id: int,
        status: str | None = None,
        scan_note: str | None = None,
        scan_note_set: bool = False,
        contract_date: str | None = None,
        scan_path: str | None = None,
        scan_replaced: bool = False,
    ) -> dict[str, Any] | None:
        """Меняет статус, отметку скана, дату и путь файла. Номер не трогает."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            created_at = datetime.now().isoformat(timespec="seconds")
            self._apply_status(conn, current, status, user_id, created_at)
            self._apply_scan_note(conn, current, scan_note, scan_note_set, user_id, created_at)
            self._apply_date(conn, current, contract_date, user_id, created_at)
            self._apply_scan_file(conn, current, scan_path, scan_replaced, user_id, created_at)
            conn.execute("COMMIT")
            return self._by_id(conn, contract_id)
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def update_contract_date(
        self,
        contract_id: int,
        contract_date: date,
        *,
        user_id: int,
    ) -> dict[str, Any] | None:
        """Меняет дату. Номер не пересчитывается."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            new_value = contract_date.isoformat()
            if current["contract_date"] == new_value:
                conn.execute("COMMIT")
                return current
            created_at = datetime.now().isoformat(timespec="seconds")
            conn.execute(
                "UPDATE supply_contract SET contract_date = ? WHERE id = ?",
                (new_value, int(contract_id)),
            )
            self._insert_event(
                conn,
                contract_id=int(contract_id),
                field="contract_date",
                old_value=current["contract_date"],
                new_value=new_value,
                user_id=user_id,
                created_at=created_at,
            )
            conn.execute("COMMIT")
            return self._by_id(conn, contract_id)
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def change_number(
        self,
        contract_id: int,
        number: str,
        *,
        user_id: int,
    ) -> dict[str, Any] | None:
        """Меняет номер и пишет событие. Дату не трогает.

        Тот же номер после нормализации — пустая операция.
        Занятый номер поднимает IntegrityError как есть.
        """
        normalized = normalize_contract_number(number)
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            if current["number"] == normalized:
                conn.execute("COMMIT")
                return current
            created_at = datetime.now().isoformat(timespec="seconds")
            conn.execute(
                "UPDATE supply_contract SET number = ? WHERE id = ?",
                (normalized, int(contract_id)),
            )
            self._insert_event(
                conn,
                contract_id=int(contract_id),
                field="number",
                old_value=str(current["number"]),
                new_value=normalized,
                user_id=user_id,
                created_at=created_at,
            )
            conn.execute("COMMIT")
            stored = self._by_id(conn, contract_id)
            if stored is None:
                raise sqlite3.OperationalError("карточка не прочитана после смены номера")
            return stored
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def delete_card(self, contract_id: int) -> dict[str, Any] | None:
        """Удаляет строку договора и её события. Справочник контрагентов не трогает."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, int(contract_id))
            if current is None:
                conn.execute("COMMIT")
                return None
            conn.execute(
                "DELETE FROM supply_contract_event WHERE contract_id = ?",
                (int(contract_id),),
            )
            removed = conn.execute(
                "DELETE FROM supply_contract WHERE id = ?",
                (int(contract_id),),
            )
            if removed.rowcount != 1:
                raise sqlite3.OperationalError("карточка не удалена")
            conn.execute("COMMIT")
            return current
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def update_imported_name_and_inn(
        self,
        contract_id: int,
        *,
        imported_name: str,
        inn: str | None,
    ) -> dict[str, Any] | None:
        """Имя из списка и ИНН на той же строке. Событие не пишет."""
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            conn.execute(
                """
                UPDATE supply_contract
                SET imported_name = ?, inn = ?
                WHERE id = ?
                """,
                (imported_name, inn, int(contract_id)),
            )
            conn.execute("COMMIT")
            stored = self._by_id(conn, contract_id)
            if stored is None:
                raise sqlite3.OperationalError("карточка не прочитана после правки имени")
            return stored
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def save_requisites(
        self,
        contract_id: int,
        fields: dict[str, Any],
        *,
        manager_name: str,
        contract_date: date | None = None,
    ) -> dict[str, Any] | None:
        """Поля бланка на той же строке. Номер не меняет.

        Дату пишет только если её передали. Менеджера — только если в строке пусто.
        """
        conn = self._connect()
        try:
            conn.execute("BEGIN IMMEDIATE")
            current = self._by_id(conn, contract_id)
            if current is None:
                conn.execute("COMMIT")
                return None
            assignments, values = _requisite_assignments(
                fields,
                contract_date=contract_date,
                current_manager=current.get("manager_name"),
                manager_name=manager_name,
            )
            if assignments:
                set_clause = ", ".join(assignments)
                conn.execute(
                    f"UPDATE supply_contract SET {set_clause} WHERE id = ?",
                    (*values, int(contract_id)),
                )
            conn.execute("COMMIT")
            stored = self._by_id(conn, contract_id)
            if stored is None:
                raise sqlite3.OperationalError("карточка не прочитана после записи реквизитов")
            return stored
        except Exception:
            self._rollback_quietly(conn)
            raise
        finally:
            conn.close()

    def _apply_status(
        self,
        conn: sqlite3.Connection,
        current: dict[str, Any],
        status: str | None,
        user_id: int,
        created_at: str,
    ) -> None:
        if status is None or status == current["status"]:
            return
        if status not in CONTRACT_STATUSES:
            raise ValueError(status)
        conn.execute(
            "UPDATE supply_contract SET status = ? WHERE id = ?",
            (status, int(current["id"])),
        )
        self._insert_event(
            conn,
            contract_id=int(current["id"]),
            field="status",
            old_value=current["status"],
            new_value=status,
            user_id=user_id,
            created_at=created_at,
        )
        current["status"] = status

    def _apply_scan_note(
        self,
        conn: sqlite3.Connection,
        current: dict[str, Any],
        scan_note: str | None,
        scan_note_set: bool,
        user_id: int,
        created_at: str,
    ) -> None:
        if not scan_note_set:
            return
        new_value = scan_note or None
        if new_value is not None and new_value not in SCAN_NOTES:
            raise ValueError(new_value)
        old_value = current.get("scan_note")
        if old_value == new_value:
            return
        conn.execute(
            "UPDATE supply_contract SET scan_note = ? WHERE id = ?",
            (new_value, int(current["id"])),
        )
        self._insert_event(
            conn,
            contract_id=int(current["id"]),
            field="scan_note",
            old_value=old_value,
            new_value=new_value or "",
            user_id=user_id,
            created_at=created_at,
        )

    def _apply_date(
        self,
        conn: sqlite3.Connection,
        current: dict[str, Any],
        contract_date: str | None,
        user_id: int,
        created_at: str,
    ) -> None:
        if contract_date is None or contract_date == current["contract_date"]:
            return
        conn.execute(
            "UPDATE supply_contract SET contract_date = ? WHERE id = ?",
            (contract_date, int(current["id"])),
        )
        self._insert_event(
            conn,
            contract_id=int(current["id"]),
            field="contract_date",
            old_value=current["contract_date"],
            new_value=contract_date,
            user_id=user_id,
            created_at=created_at,
        )

    def _apply_scan_file(
        self,
        conn: sqlite3.Connection,
        current: dict[str, Any],
        scan_path: str | None,
        scan_replaced: bool,
        user_id: int,
        created_at: str,
    ) -> None:
        if not scan_replaced or not scan_path:
            return
        old_flag = "есть" if current.get("scan_path") else None
        conn.execute(
            "UPDATE supply_contract SET scan_path = ? WHERE id = ?",
            (scan_path, int(current["id"])),
        )
        self._insert_event(
            conn,
            contract_id=int(current["id"]),
            field="scan",
            old_value=old_flag,
            new_value="есть",
            user_id=user_id,
            created_at=created_at,
        )

    def _directory_party(
        self,
        conn: sqlite3.Connection,
        counterparty_id: int,
    ) -> dict[str, Any] | None:
        row = conn.execute(
            """
            SELECT id, name, inn, is_active, is_client
            FROM counterparties
            WHERE id = ?
            """,
            (int(counterparty_id),),
        ).fetchone()
        return None if row is None else dict(row)

    def _link_directory_choice(
        self,
        conn: sqlite3.Connection,
        party: dict[str, Any],
        *,
        user_id: int,
        number: str | None,
        contract_id: int | None,
    ) -> DirectoryWrite:
        has_number = bool(number and str(number).strip())
        has_id = contract_id is not None
        if has_number and has_id:
            return DirectoryWrite(refusal="both")
        if not has_number and not has_id:
            return DirectoryWrite(refusal="empty")
        if has_id:
            card = self._by_id(conn, int(contract_id))
            if card is None:
                return DirectoryWrite(refusal="missing_card")
            return self._claim_directory_card(conn, party, card, user_id=user_id)
        normalized = normalize_contract_number(str(number))
        card = self._card_by_number(conn, normalized)
        if card is None:
            return self._insert_known_directory_number(
                conn,
                party,
                number=normalized,
                user_id=user_id,
            )
        return self._claim_directory_card(conn, party, card, user_id=user_id)

    def _claim_directory_card(
        self,
        conn: sqlite3.Connection,
        party: dict[str, Any],
        card: dict[str, Any],
        *,
        user_id: int,
    ) -> DirectoryWrite:
        client_id = int(party["id"])
        owner = card.get("counterparty_id")
        shown = str(card["number"])
        if owner is not None and int(owner) != client_id:
            return DirectoryWrite(refusal="foreign", hint=shown)
        if owner is not None:
            return DirectoryWrite(row=card)
        if str(card.get("status") or "") == CANCELLED_STATUS:
            return DirectoryWrite(refusal="occupied", hint=shown)
        active = self._active_row(conn, client_id)
        if active is not None and int(active["id"]) != int(card["id"]):
            return DirectoryWrite(refusal="busy", hint=str(active["number"]))
        if not _card_inn_matches(card, party):
            return DirectoryWrite(refusal="inn")
        stored = self._assign_directory_owner(
            conn,
            card,
            client_id,
            user_id=user_id,
        )
        return DirectoryWrite(row=stored)

    def _assign_directory_owner(
        self,
        conn: sqlite3.Connection,
        card: dict[str, Any],
        client_id: int,
        *,
        user_id: int,
    ) -> dict[str, Any]:
        contract_id = int(card["id"])
        updated = conn.execute(
            """
            UPDATE supply_contract
            SET counterparty_id = ?
            WHERE id = ?
              AND counterparty_id IS NULL
              AND status != ?
              AND NOT EXISTS (
                  SELECT 1 FROM supply_contract AS other
                  WHERE other.counterparty_id = ?
                    AND other.status != ?
                    AND other.id != ?
              )
            """,
            (
                client_id,
                contract_id,
                CANCELLED_STATUS,
                client_id,
                CANCELLED_STATUS,
                contract_id,
            ),
        )
        if updated.rowcount != 1:
            raise sqlite3.IntegrityError("карточка не привязана")
        created_at = datetime.now().isoformat(timespec="seconds")
        self._insert_event(
            conn,
            contract_id=contract_id,
            field="counterparty_id",
            old_value=None,
            new_value=str(client_id),
            user_id=user_id,
            created_at=created_at,
        )
        stored = self._by_id(conn, contract_id)
        if stored is None:
            raise sqlite3.OperationalError("карточка не прочитана после привязки")
        return stored

    def _insert_known_directory_number(
        self,
        conn: sqlite3.Connection,
        party: dict[str, Any],
        *,
        number: str,
        user_id: int,
    ) -> DirectoryWrite:
        """Вписанный номер сразу с id клиента. Seq не считает."""
        client_id = int(party["id"])
        active = self._active_row(conn, client_id)
        if active is not None:
            if str(active["number"]) == number:
                return DirectoryWrite(row=active)
            return DirectoryWrite(refusal="busy", hint=str(active["number"]))
        contract_date = contract_date_from_number(number)
        created_at = datetime.now().isoformat(timespec="seconds")
        cur = conn.execute(
            """
            INSERT INTO supply_contract (
                counterparty_id, imported_name, inn, number, contract_date,
                status, manager_name, created_at, created_by_user_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                client_id,
                str(party["name"]),
                party.get("inn"),
                number,
                None if contract_date is None else contract_date.isoformat(),
                DEFAULT_STATUS,
                None,
                created_at,
                int(user_id),
            ),
        )
        contract_id = int(cur.lastrowid)
        self._insert_event(
            conn,
            contract_id=contract_id,
            field="status",
            old_value=None,
            new_value=DEFAULT_STATUS,
            user_id=user_id,
            created_at=created_at,
        )
        stored = self._by_id(conn, contract_id)
        if stored is None:
            raise sqlite3.OperationalError("договор не прочитан после вставки")
        return DirectoryWrite(row=stored)

    def _write_blank_requisites(
        self,
        conn: sqlite3.Connection,
        current: dict[str, Any],
        fields: dict[str, Any],
        *,
        manager_name: str,
        contract_date: date | None,
    ) -> dict[str, Any]:
        assignments, values = _requisite_assignments(
            fields,
            contract_date=contract_date,
            current_manager=current.get("manager_name"),
            manager_name=manager_name,
        )
        if assignments:
            conn.execute(
                f"UPDATE supply_contract SET {', '.join(assignments)} WHERE id = ?",
                (*values, int(current["id"])),
            )
        stored = self._by_id(conn, int(current["id"]))
        if stored is None:
            raise sqlite3.OperationalError("карточка не прочитана после записи реквизитов")
        return stored

    def _insert_directory_blank(
        self,
        conn: sqlite3.Connection,
        party: dict[str, Any],
        fields: dict[str, Any],
        *,
        user_id: int,
        manager_name: str,
        number_date: date,
        contract_date: date,
    ) -> dict[str, Any]:
        number = self._allocate_number(conn, number_date)
        created_at = datetime.now().isoformat(timespec="seconds")
        columns = (
            "counterparty_id",
            "imported_name",
            "number",
            "contract_date",
            "manager_name",
            "status",
            *_REQUISITE_COLUMNS,
            "created_at",
            "created_by_user_id",
        )
        values = {
            "counterparty_id": int(party["id"]),
            "imported_name": str(party["name"]),
            "number": number,
            "contract_date": contract_date.isoformat(),
            "manager_name": manager_name,
            "status": DEFAULT_STATUS,
            **fields,
            "created_at": created_at,
            "created_by_user_id": int(user_id),
        }
        placeholders = ", ".join("?" for _ in columns)
        names = ", ".join(columns)
        cur = conn.execute(
            f"INSERT INTO supply_contract ({names}) VALUES ({placeholders})",
            tuple(values.get(column) for column in columns),
        )
        contract_id = int(cur.lastrowid)
        self._insert_event(
            conn,
            contract_id=contract_id,
            field="status",
            old_value=None,
            new_value=DEFAULT_STATUS,
            user_id=user_id,
            created_at=created_at,
        )
        stored = self._by_id(conn, contract_id)
        if stored is None:
            raise sqlite3.OperationalError("договор не прочитан после вставки")
        return stored

    def _card_by_number(self, conn: sqlite3.Connection, number: str) -> dict[str, Any] | None:
        row = conn.execute(
            "SELECT * FROM supply_contract WHERE number = ?",
            (number,),
        ).fetchone()
        return _row_to_dict(row)

    def _fetch_link_inputs(
        self, conn: sqlite3.Connection
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        contracts = [
            dict(row)
            for row in conn.execute(
                """
                SELECT id, counterparty_id, status, inn, contract_date, number
                FROM supply_contract
                """
            )
        ]
        clients = [
            dict(row)
            for row in conn.execute(
                "SELECT id, inn, is_active, is_client FROM counterparties"
            )
        ]
        return contracts, clients

    def _count_clients(self, conn: sqlite3.Connection) -> int:
        row = conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()
        return int(row[0])

    def _apply_unique_links(
        self,
        conn: sqlite3.Connection,
        pairs: list[tuple[int, int]],
        *,
        user_id: int | None,
    ) -> int:
        created = 0
        created_at = datetime.now().isoformat(timespec="seconds")
        for contract_id, client_id in pairs:
            cur = conn.execute(
                _LINK_UPDATE,
                (
                    int(client_id),
                    int(contract_id),
                    CANCELLED_STATUS,
                    int(client_id),
                    CANCELLED_STATUS,
                ),
            )
            if cur.rowcount != 1:
                continue
            created += 1
            if user_id is None:
                continue
            self._insert_event(
                conn,
                contract_id=int(contract_id),
                field="counterparty_id",
                old_value=None,
                new_value=str(int(client_id)),
                user_id=int(user_id),
                created_at=created_at,
            )
        return created

    def _active_row(self, conn: sqlite3.Connection, counterparty_id: int) -> dict[str, Any] | None:
        row = conn.execute(
            """
            SELECT * FROM supply_contract
            WHERE counterparty_id = ? AND status != ?
            """,
            (int(counterparty_id), CANCELLED_STATUS),
        ).fetchone()
        return _row_to_dict(row)

    def _by_id(self, conn: sqlite3.Connection, contract_id: int) -> dict[str, Any] | None:
        row = conn.execute(
            "SELECT * FROM supply_contract WHERE id = ?",
            (int(contract_id),),
        ).fetchone()
        return _row_to_dict(row)

    def _insert_missing(
        self,
        conn: sqlite3.Connection,
        rows: list[dict[str, Any]],
        taken: set[str],
        user_id: int,
    ) -> tuple[int, int]:
        created = 0
        skipped = 0
        for row in rows:
            number = str(row["number"])
            if number in taken:
                skipped += 1
                continue
            inn = row.get("inn")
            self._insert_card_row(
                conn,
                imported_name=str(row["imported_name"]),
                inn=None if inn is None else str(inn),
                number=number,
                contract_date=row.get("contract_date"),
                user_id=user_id,
            )
            taken.add(number)
            created += 1
        return created, skipped

    def _insert_card_row(
        self,
        conn: sqlite3.Connection,
        *,
        imported_name: str,
        inn: str | None,
        number: str,
        contract_date: date | None,
        user_id: int,
    ) -> int:
        created_at = datetime.now().isoformat(timespec="seconds")
        stored_date = None if contract_date is None else contract_date.isoformat()
        cur = conn.execute(
            """
            INSERT INTO supply_contract (
                counterparty_id, imported_name, inn, number, contract_date,
                status, manager_name, created_at, created_by_user_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                None,
                imported_name,
                inn,
                number,
                stored_date,
                DEFAULT_STATUS,
                None,
                created_at,
                int(user_id),
            ),
        )
        return int(cur.lastrowid)

    def _insert_imported_row(
        self,
        conn: sqlite3.Connection,
        row: dict[str, Any],
        *,
        user_id: int,
        created_at: str,
    ) -> int:
        name = str(row["imported_name"])
        values = (
            None,
            name,
            str(row["number"]),
            str(row["contract_date"]),
            str(row["manager_name"]),
            str(row["status"]),
            row.get("scan_note"),
            "",
            name,
            name,
            "",
            "",
            "",
            "",
            None,
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            created_at,
            int(user_id),
        )
        columns = (
            "counterparty_id, imported_name, number, contract_date, manager_name, status, "
            "scan_note, legal_form, full_name, short_name, signatory_name, signatory_verb, "
            "authority_basis, inn, kpp, ogrn, legal_address, email, bank_name, account, "
            "corr_account, bik, created_at, created_by_user_id"
        )
        placeholders = ", ".join("?" for _ in values)
        try:
            cur = conn.execute(
                f"INSERT INTO supply_contract ({columns}) VALUES ({placeholders})",
                values,
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"duplicate:{row['number']}") from exc
        contract_id = int(cur.lastrowid)
        self._insert_event(
            conn,
            contract_id=contract_id,
            field="import",
            old_value=None,
            new_value=str(row["number"]),
            user_id=user_id,
            created_at=created_at,
        )
        return contract_id

    def _allocate_number(self, conn: sqlite3.Connection, contract_date: date) -> str:
        rows = conn.execute("SELECT number FROM supply_contract").fetchall()
        occupied = [str(row["number"]) for row in rows]
        seq = next_contract_seq(occupied)
        return format_contract_number(seq, contract_date)

    def _insert_event(
        self,
        conn: sqlite3.Connection,
        *,
        contract_id: int,
        field: str,
        old_value: str | None,
        new_value: str,
        user_id: int,
        created_at: str,
    ) -> None:
        conn.execute(
            """
            INSERT INTO supply_contract_event (
                contract_id, field, old_value, new_value, user_id, created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (contract_id, field, old_value, new_value, int(user_id), created_at),
        )

    def _rollback_quietly(self, conn: sqlite3.Connection) -> None:
        try:
            conn.execute("ROLLBACK")
        except sqlite3.OperationalError:
            pass
