"""SQL-доступ к договору поставки. Номер выдаётся в одной транзакции."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from typing import Any

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.supply_contract import (
    CANCELLED_STATUS,
    CONTRACT_STATUSES,
    DEFAULT_STATUS,
    SCAN_NOTES,
    format_contract_number,
    next_contract_seq,
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


def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    return dict(row)


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
