"""Схема supply_contract: пустые реквизиты и дата (CDC-002)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema
from core.supply_contract import CANCELLED_STATUS

NULLABLE_COLUMNS = (
    "contract_date",
    "manager_name",
    "legal_form",
    "full_name",
    "short_name",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "inn",
    "ogrn",
    "legal_address",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
)

REQUIRED_COLUMNS = (
    "number",
    "status",
    "created_at",
    "created_by_user_id",
)

_CONTRACT_SNAPSHOT = """
SELECT
    id, counterparty_id, imported_name, number, contract_date, manager_name,
    status, scan_note, scan_path, legal_form, full_name, short_name,
    signatory_position, signatory_name, signatory_verb, authority_basis,
    poa_number, poa_date, inn, kpp, ogrn, legal_address, postal_address,
    phone, email, bank_name, account, corr_account, bik, edo_operator,
    edo_id, okved, created_at, created_by_user_id
FROM supply_contract
ORDER BY id
"""

_FILLED_CONTRACT = (
    41,
    5,
    None,
    "0001/10/26",
    "2026-10-01",
    "Иван",
    "есть",
    "на полке",
    "scans/0001.pdf",
    "ooo",
    "ООО Ромашка",
    "Ромашка",
    "директор",
    "Иванов Иван",
    "действующего",
    "устав",
    None,
    None,
    "7701000001",
    "770101001",
    "1027700000001",
    "Москва",
    "Москва, а/я",
    "+79990000000",
    "a@b.ru",
    "Банк",
    "40702810000000000001",
    "30101810000000000001",
    "044525225",
    None,
    None,
    "23.61",
    "2026-10-01T09:00:00",
    7,
)

_LAWYER_CONTRACT = (
    42,
    None,
    "Секлюцкий",
    "0785/01/25",
    "",
    "",
    "нет",
    None,
    None,
    "",
    "",
    "",
    None,
    "",
    "",
    "",
    None,
    None,
    "",
    None,
    "",
    "",
    None,
    None,
    "",
    "",
    "",
    "",
    "",
    None,
    None,
    None,
    "2026-01-15T00:00:00",
    3,
)

_EVENT = (17, 41, "number", "0000/09/26", "0001/10/26", 7, "2026-10-02T10:00:00")


def _ensure(db_path: Path) -> None:
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(str(db_path))


def _notnull_by_name(conn: sqlite3.Connection) -> dict[str, int]:
    return {
        str(row[1]): int(row[3])
        for row in conn.execute("PRAGMA table_info(supply_contract)")
    }


def _assert_requisites_nullable(conn: sqlite3.Connection) -> None:
    flags = _notnull_by_name(conn)
    for name in NULLABLE_COLUMNS:
        assert flags[name] == 0
    for name in REQUIRED_COLUMNS:
        assert flags[name] == 1
    defaults = {
        str(row[1]): row[4]
        for row in conn.execute("PRAGMA table_info(supply_contract)")
    }
    assert defaults["status"] == "'нет'"


def _assert_partial_active_index(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT sql FROM sqlite_master
        WHERE type = 'index' AND name = 'idx_supply_contract_one_active'
        """
    ).fetchall()
    assert len(rows) == 1
    sql = rows[0][0]
    assert "counterparty_id IS NOT NULL" in sql
    assert CANCELLED_STATUS in sql


def _legacy_db(tmp_path: Path) -> Path:
    """Старая supply_contract: реквизиты и дата ещё NOT NULL."""
    db_path = tmp_path / "legacy-supply-contract.db"
    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute(
            """
            CREATE TABLE counterparties (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code_1c TEXT NOT NULL UNIQUE,
                guid_1c TEXT UNIQUE,
                name TEXT NOT NULL,
                name_normalized TEXT NOT NULL,
                inn TEXT,
                kpp TEXT,
                is_client INTEGER NOT NULL DEFAULT 1,
                is_active INTEGER NOT NULL DEFAULT 1,
                source TEXT NOT NULL DEFAULT 'import',
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                updated_at TEXT NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        conn.execute(
            """
            INSERT INTO counterparties (
                id, code_1c, name, name_normalized, inn, kpp
            ) VALUES (
                5, '00-0001', 'ООО Ромашка', 'ооо ромашка', '7701000001', '770101001'
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE supply_contract (
                id INTEGER PRIMARY KEY,
                counterparty_id INTEGER NULL REFERENCES counterparties(id),
                imported_name TEXT,
                number TEXT NOT NULL UNIQUE,
                contract_date TEXT NOT NULL,
                manager_name TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'нет',
                scan_note TEXT,
                scan_path TEXT,
                legal_form TEXT NOT NULL,
                full_name TEXT NOT NULL,
                short_name TEXT NOT NULL,
                signatory_position TEXT,
                signatory_name TEXT NOT NULL,
                signatory_verb TEXT NOT NULL,
                authority_basis TEXT NOT NULL,
                poa_number TEXT,
                poa_date TEXT,
                inn TEXT NOT NULL,
                kpp TEXT,
                ogrn TEXT NOT NULL,
                legal_address TEXT NOT NULL,
                postal_address TEXT,
                phone TEXT,
                email TEXT NOT NULL,
                bank_name TEXT NOT NULL,
                account TEXT NOT NULL,
                corr_account TEXT NOT NULL,
                bik TEXT NOT NULL,
                edo_operator TEXT,
                edo_id TEXT,
                okved TEXT,
                created_at TEXT NOT NULL,
                created_by_user_id INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            INSERT INTO supply_contract (
                id, counterparty_id, imported_name, number, contract_date,
                manager_name, status, scan_note, scan_path, legal_form, full_name,
                short_name, signatory_position, signatory_name, signatory_verb,
                authority_basis, poa_number, poa_date, inn, kpp, ogrn,
                legal_address, postal_address, phone, email, bank_name, account,
                corr_account, bik, edo_operator, edo_id, okved, created_at,
                created_by_user_id
            ) VALUES (
                41, 5, NULL, '0001/10/26', '2026-10-01',
                'Иван', 'есть', 'на полке', 'scans/0001.pdf', 'ooo', 'ООО Ромашка',
                'Ромашка', 'директор', 'Иванов Иван', 'действующего',
                'устав', NULL, NULL, '7701000001', '770101001', '1027700000001',
                'Москва', 'Москва, а/я', '+79990000000', 'a@b.ru', 'Банк',
                '40702810000000000001',
                '30101810000000000001', '044525225', NULL, NULL, '23.61',
                '2026-10-01T09:00:00',
                7
            )
            """
        )
        conn.execute(
            """
            INSERT INTO supply_contract (
                id, counterparty_id, imported_name, number, contract_date,
                manager_name, status, scan_note, scan_path, legal_form, full_name,
                short_name, signatory_position, signatory_name, signatory_verb,
                authority_basis, poa_number, poa_date, inn, kpp, ogrn,
                legal_address, postal_address, phone, email, bank_name, account,
                corr_account, bik, edo_operator, edo_id, okved, created_at,
                created_by_user_id
            ) VALUES (
                42, NULL, 'Секлюцкий', '0785/01/25', '',
                '', 'нет', NULL, NULL, '', '',
                '', NULL, '', '',
                '', NULL, NULL, '', NULL, '',
                '', NULL, NULL, '', '',
                '',
                '', '', NULL, NULL, NULL,
                '2026-01-15T00:00:00',
                3
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE supply_contract_event (
                id INTEGER PRIMARY KEY,
                contract_id INTEGER NOT NULL REFERENCES supply_contract(id),
                field TEXT NOT NULL,
                old_value TEXT,
                new_value TEXT,
                user_id INTEGER NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            INSERT INTO supply_contract_event (
                id, contract_id, field, old_value, new_value, user_id, created_at
            ) VALUES (17, 41, 'number', '0000/09/26', '0001/10/26', 7, '2026-10-02T10:00:00')
            """
        )
        conn.commit()
    return db_path


def _assert_copied_legacy_rows(conn: sqlite3.Connection) -> None:
    assert conn.execute(_CONTRACT_SNAPSHOT).fetchall() == [
        _FILLED_CONTRACT,
        _LAWYER_CONTRACT,
    ]
    event = conn.execute(
        """
        SELECT id, contract_id, field, old_value, new_value, user_id, created_at
        FROM supply_contract_event
        """
    ).fetchone()
    assert event == _EVENT


def _insert_bare(
    conn: sqlite3.Connection,
    *,
    number: str,
    counterparty_id: int | None,
    status: str = "нет",
) -> None:
    conn.execute(
        """
        INSERT INTO supply_contract (
            counterparty_id, number, status, created_at, created_by_user_id
        ) VALUES (?, ?, ?, '2026-10-20T00:00:00', 1)
        """,
        (counterparty_id, number, status),
    )


def test_fresh_supply_contract_allows_empty_requisites_and_date(tmp_path: Path) -> None:
    db_path = tmp_path / "fresh-supply-contract.db"
    _ensure(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        _assert_requisites_nullable(conn)
        _assert_partial_active_index(conn)

        _insert_bare(conn, number="2001/10/26", counterparty_id=None)
        _insert_bare(conn, number="2002/10/26", counterparty_id=None)
        conn.commit()
        stored = conn.execute(
            """
            SELECT number, contract_date, legal_form, inn, manager_name, counterparty_id
            FROM supply_contract
            ORDER BY number
            """
        ).fetchall()
        assert stored == [
            ("2001/10/26", None, None, None, None, None),
            ("2002/10/26", None, None, None, None, None),
        ]

        with pytest.raises(sqlite3.IntegrityError):
            _insert_bare(conn, number="2001/10/26", counterparty_id=None)
        conn.rollback()
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """
                INSERT INTO supply_contract (
                    number, status, created_at, created_by_user_id
                ) VALUES (NULL, 'нет', '2026-10-20T00:00:00', 1)
                """
            )
        conn.rollback()


def test_legacy_supply_contract_keeps_rows_when_requisites_become_nullable(
    tmp_path: Path,
) -> None:
    db_path = _legacy_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        counterparties_sql = conn.execute(
            "SELECT sql FROM sqlite_master WHERE name = 'counterparties'"
        ).fetchone()[0]
        counterparty = conn.execute(
            """
            SELECT id, code_1c, name, name_normalized, inn, kpp
            FROM counterparties
            """
        ).fetchone()

    _ensure(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        _assert_requisites_nullable(conn)
        _assert_partial_active_index(conn)
        _assert_copied_legacy_rows(conn)
        assert conn.execute("PRAGMA foreign_key_check(supply_contract)").fetchall() == []
        assert (
            conn.execute("PRAGMA foreign_key_check(supply_contract_event)").fetchall()
            == []
        )
        assert (
            conn.execute(
                "SELECT sql FROM sqlite_master WHERE name = 'counterparties'"
            ).fetchone()[0]
            == counterparties_sql
        )
        assert (
            conn.execute(
                """
                SELECT id, code_1c, name, name_normalized, inn, kpp
                FROM counterparties
                """
            ).fetchone()
            == counterparty
        )
        assert (
            conn.execute(
                """
                SELECT name FROM sqlite_master
                WHERE type = 'table' AND name = 'supply_contract_nullable'
                """
            ).fetchone()
            is None
        )

        _insert_bare(conn, number="2001/10/26", counterparty_id=None)
        _insert_bare(conn, number="2002/10/26", counterparty_id=None)
        conn.commit()
        with pytest.raises(sqlite3.IntegrityError):
            _insert_bare(conn, number="0001/10/26", counterparty_id=None)
        conn.rollback()

        _insert_bare(
            conn,
            number="0002/10/26",
            counterparty_id=5,
            status=CANCELLED_STATUS,
        )
        conn.commit()
        with pytest.raises(sqlite3.IntegrityError):
            _insert_bare(conn, number="0003/10/26", counterparty_id=5, status="нет")
        conn.rollback()

        numbers = {
            row[0]
            for row in conn.execute("SELECT number FROM supply_contract")
        }
        assert numbers == {
            "0001/10/26",
            "0785/01/25",
            "2001/10/26",
            "2002/10/26",
            "0002/10/26",
        }


def test_supply_contract_nullability_migration_is_idempotent(tmp_path: Path) -> None:
    db_path = _legacy_db(tmp_path)
    _ensure(db_path)
    with sqlite3.connect(db_path) as conn:
        _insert_bare(conn, number="2001/10/26", counterparty_id=None)
        conn.commit()
        flags = _notnull_by_name(conn)
        rows = conn.execute(_CONTRACT_SNAPSHOT).fetchall()
        event = conn.execute(
            """
            SELECT id, contract_id, field, old_value, new_value, user_id, created_at
            FROM supply_contract_event
            """
        ).fetchone()

    _ensure(db_path)

    with sqlite3.connect(db_path) as conn:
        assert _notnull_by_name(conn) == flags
        assert conn.execute(_CONTRACT_SNAPSHOT).fetchall() == rows
        assert (
            conn.execute(
                """
                SELECT id, contract_id, field, old_value, new_value, user_id, created_at
                FROM supply_contract_event
                """
            ).fetchone()
            == event
        )
        added = conn.execute(
            """
            SELECT contract_date, legal_form, full_name, inn
            FROM supply_contract
            WHERE number = '2001/10/26'
            """
        ).fetchone()
        assert added == (None, None, None, None)
        _assert_partial_active_index(conn)


@pytest.mark.parametrize("trigger", ["contract_date", "legal_form"])
def test_either_notnull_trigger_rebuilds_supply_contract(
    tmp_path: Path,
    trigger: str,
) -> None:
    db_path = tmp_path / f"trigger-{trigger}.db"
    contract_date_sql = "TEXT NOT NULL" if trigger == "contract_date" else "TEXT"
    legal_form_sql = "TEXT NOT NULL" if trigger == "legal_form" else "TEXT"
    contract_date_value = "2026-10-01" if trigger == "contract_date" else None
    legal_form_value = "ooo" if trigger == "legal_form" else None
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            f"""
            CREATE TABLE supply_contract (
                id INTEGER PRIMARY KEY,
                counterparty_id INTEGER NULL,
                number TEXT NOT NULL UNIQUE,
                contract_date {contract_date_sql},
                legal_form {legal_form_sql},
                status TEXT NOT NULL DEFAULT 'нет',
                created_at TEXT NOT NULL,
                created_by_user_id INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            INSERT INTO supply_contract (
                id, number, contract_date, legal_form, status,
                created_at, created_by_user_id
            ) VALUES (7, '0001/10/26', ?, ?, 'есть', '2026-10-01T09:00:00', 7)
            """,
            (contract_date_value, legal_form_value),
        )
        conn.commit()

    _ensure(db_path)

    with sqlite3.connect(db_path) as conn:
        flags = _notnull_by_name(conn)
        assert flags["contract_date"] == 0
        assert flags["legal_form"] == 0
        assert flags["number"] == 1
        assert flags["status"] == 1
        assert flags["created_at"] == 1
        assert flags["created_by_user_id"] == 1
        assert "manager_name" not in flags
        assert "okved" in flags
        row = conn.execute(
            """
            SELECT id, number, contract_date, legal_form, status, okved,
                   created_by_user_id
            FROM supply_contract
            """
        ).fetchone()
        assert row == (
            7,
            "0001/10/26",
            contract_date_value,
            legal_form_value,
            "есть",
            None,
            7,
        )
        _assert_partial_active_index(conn)
