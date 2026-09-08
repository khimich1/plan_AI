"""CTR-001: counterparties table + KP_offers snapshot columns."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema

_REQUIRED_COLUMNS = {
    "id",
    "code_1c",
    "guid_1c",
    "name",
    "name_normalized",
    "inn",
    "kpp",
    "is_client",
    "is_active",
    "source",
    "created_at",
    "updated_at",
}

_KP_SNAPSHOT_COLUMNS = ("counterparty_id", "customer_inn", "customer_kpp")


def _table_columns(cur: sqlite3.Cursor, table: str) -> dict[str, tuple]:
    cur.execute(f"PRAGMA table_info({table})")
    return {row[1]: row for row in cur.fetchall()}


def _index_names(cur: sqlite3.Cursor, table: str) -> set[str]:
    cur.execute(f"PRAGMA index_list({table})")
    return {row[1] for row in cur.fetchall()}


def test_fresh_schema_has_counterparties_columns_and_indexes(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cols = _table_columns(cur, "counterparties")
        assert _REQUIRED_COLUMNS <= cols.keys()

        assert cols["code_1c"][3] == 1  # NOT NULL
        assert cols["name"][3] == 1
        assert cols["name_normalized"][3] == 1
        assert cols["is_client"][3] == 1
        assert cols["is_active"][3] == 1
        assert cols["source"][3] == 1
        assert cols["guid_1c"][3] == 0  # nullable
        assert cols["inn"][3] == 0
        assert cols["kpp"][3] == 0
        assert cols["is_client"][4] == "1"
        assert cols["is_active"][4] == "1"
        assert cols["source"][4] in ("'import'", "import")

        indexes = _index_names(cur, "counterparties")
        assert "idx_counterparties_name_norm" in indexes
        assert "idx_counterparties_inn" in indexes

        cur.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized)
            VALUES ('00-00000001', 'ООО Ромашка', 'ооо ромашка')
            """
        )
        with pytest.raises(sqlite3.IntegrityError):
            cur.execute(
                """
                INSERT INTO counterparties (code_1c, name, name_normalized)
                VALUES ('00-00000001', 'Другое имя', 'другое имя')
                """
            )

        cur.execute(
            """
            INSERT INTO counterparties (code_1c, guid_1c, name, name_normalized)
            VALUES ('00-00000002', 'guid-a', 'Бета', 'бета')
            """
        )
        with pytest.raises(sqlite3.IntegrityError):
            cur.execute(
                """
                INSERT INTO counterparties (code_1c, guid_1c, name, name_normalized)
                VALUES ('00-00000003', 'guid-a', 'Гамма', 'гамма')
                """
            )


def test_fresh_schema_adds_kp_offers_snapshot_columns(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh_kp.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cols = _table_columns(cur, "KP_offers")
        for name in _KP_SNAPSHOT_COLUMNS:
            assert name in cols, f"KP_offers missing {name}"
            assert cols[name][3] == 0  # nullable — старые КП не ломаем


def test_legacy_db_migrates_counterparties_and_preserves_kp_rows(
    tmp_path: Path,
) -> None:
    db_path = str(tmp_path / "legacy.db")
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL,
                customer_name TEXT
            )
            """
        )
        cur.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) "
            "VALUES ('2026-01-01', 'Старый клиент')"
        )
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cols = _table_columns(cur, "counterparties")
        assert _REQUIRED_COLUMNS <= cols.keys()
        kp_cols = _table_columns(cur, "KP_offers")
        for name in _KP_SNAPSHOT_COLUMNS:
            assert name in kp_cols
        cur.execute("SELECT customer_name FROM KP_offers")
        assert cur.fetchone()[0] == "Старый клиент"
        cur.execute("SELECT COUNT(*) FROM counterparties")
        assert cur.fetchone()[0] == 0


def test_ensure_schema_idempotent_for_counterparties(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='counterparties'"
        )
        assert cur.fetchone() is not None
        indexes = _index_names(cur, "counterparties")
        assert "idx_counterparties_name_norm" in indexes
        assert "idx_counterparties_inn" in indexes
