"""Guard: schema init only in kp_db_schema + documented entrypoints (A4).

Also MNA-301: ``line_id`` on kp_* line tables + ``kp_meta.product_type = mixed``.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema

_ROOT = Path(__file__).resolve().parent.parent / "core"

_SLICES = (
    "kp_db_offers.py",
    "kp_db_managers.py",
    "kp_db_plates.py",
    "kp_db_rests.py",
)

_FORBIDDEN_PATTERNS = ("ensure_schema(", "_init_schema(")

_LINE_TABLES = (
    "kp_plates",
    "kp_piles",
    "kp_steps",
    "kp_marches",
    "kp_bridge_piles",
    "kp_fbs",
)


def _table_columns(cur: sqlite3.Cursor, table: str) -> dict[str, tuple]:
    cur.execute(f"PRAGMA table_info({table})")
    return {row[1]: row for row in cur.fetchall()}


def _assert_line_id_text(cur: sqlite3.Cursor, table: str) -> None:
    cols = _table_columns(cur, table)
    assert "line_id" in cols, f"{table} must have line_id column"
    assert cols["line_id"][2].upper() == "TEXT", f"{table}.line_id must be TEXT"


# --- A4 boundary -------------------------------------------------------------


@pytest.mark.parametrize("module_name", _SLICES)
def test_persistence_slices_do_not_call_schema_init(module_name: str) -> None:
    text = (_ROOT / module_name).read_text(encoding="utf-8")
    for pattern in _FORBIDDEN_PATTERNS:
        assert pattern not in text, (
            f"{module_name} must not call {pattern!r}; "
            "use startup ensure_schema or test fixtures (make_iso_db)."
        )


def test_schema_module_defines_ensure_schema() -> None:
    text = (_ROOT / "kp_db_schema.py").read_text(encoding="utf-8")
    assert "def ensure_schema(" in text
    assert "def _init_schema_impl(" in text


# --- MNA-301: line_id + mixed ------------------------------------------------


def test_fresh_schema_has_line_id_on_all_kp_line_tables(tmp_path: Path) -> None:
    """Fresh ensure_schema creates line_id TEXT on every kp_* line table."""
    db_path = str(tmp_path / "fresh_line_id.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        for table in _LINE_TABLES:
            _assert_line_id_text(cur, table)


def test_migrate_existing_db_adds_line_id_to_all_kp_line_tables(
    tmp_path: Path,
) -> None:
    """Legacy DB without line_id gains the column on all six line tables."""
    db_path = str(tmp_path / "legacy_line_id.db")
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL UNIQUE,
                status TEXT DEFAULT 'в работе',
                product_type TEXT DEFAULT 'plates',
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        # Pre-MNA-301 shapes: no line_id
        cur.execute(
            """
            CREATE TABLE kp_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER,
                plate_name TEXT NOT NULL,
                qty INTEGER NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_piles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_marches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_bridge_piles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE kp_fbs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-01-01')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) "
            "VALUES (1, 'в работе', 'plates')"
        )
        conn.commit()

    for table in _LINE_TABLES:
        with sqlite3.connect(db_path) as conn:
            cols = _table_columns(conn.cursor(), table)
            assert "line_id" not in cols, f"precondition: {table} has no line_id"

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        for table in _LINE_TABLES:
            _assert_line_id_text(cur, table)


def test_ensure_schema_idempotent_with_line_id(tmp_path: Path) -> None:
    """Second ensure_schema must not fail after line_id migration."""
    db_path = str(tmp_path / "idem_line_id.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        for table in _LINE_TABLES:
            _assert_line_id_text(cur, table)


def test_kp_meta_accepts_product_type_mixed(tmp_path: Path) -> None:
    """kp_meta.product_type may store the multi-nomenclature value ``mixed``."""
    db_path = str(tmp_path / "mixed_meta.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-08-12')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) "
            "VALUES (1, 'в работе', 'mixed')"
        )
        conn.commit()

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "mixed"


def test_line_id_column_accepts_text_values(tmp_path: Path) -> None:
    """After schema init, line_id can be written and read on each line table."""
    db_path = str(tmp_path / "line_id_write.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-08-12')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) "
            "VALUES (1, 'в работе', 'mixed')"
        )

        cur.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, qty, line_id
            ) VALUES (1, 1, 'ПБ 60-12-8п', 2, 'ln_plates_1')
            """
        )
        cur.execute(
            """
            INSERT INTO kp_piles (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price, line_id
            ) VALUES (1, 2, 'С120.35-12', 'B25', 1, 100.0, 100.0, 'ln_piles_1')
            """
        )
        cur.execute(
            """
            INSERT INTO kp_steps (
                kp_id, position_number, mark,
                qty, unit_price, discounted_price, line_id
            ) VALUES (1, 3, 'ЛС11', 1, 10.0, 10.0, 'ln_steps_1')
            """
        )
        cur.execute(
            """
            INSERT INTO kp_marches (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price, line_id
            ) VALUES (1, 4, '1ЛМ 27-11-14-4', 'B25', 1, 20.0, 20.0, 'ln_marches_1')
            """
        )
        cur.execute(
            """
            INSERT INTO kp_bridge_piles (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price, line_id
            ) VALUES (1, 5, 'C8-35В4', 'B25', 1, 30.0, 30.0, 'ln_bridge_1')
            """
        )
        cur.execute(
            """
            INSERT INTO kp_fbs (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price, line_id
            ) VALUES (1, 6, 'ФБС 9.3.6-Т', 'B25', 1, 40.0, 40.0, 'ln_fbs_1')
            """
        )
        conn.commit()

        expected = {
            "kp_plates": "ln_plates_1",
            "kp_piles": "ln_piles_1",
            "kp_steps": "ln_steps_1",
            "kp_marches": "ln_marches_1",
            "kp_bridge_piles": "ln_bridge_1",
            "kp_fbs": "ln_fbs_1",
        }
        for table, line_id in expected.items():
            cur.execute(f"SELECT line_id FROM {table} WHERE kp_id = 1")
            assert cur.fetchone()[0] == line_id


def test_fresh_schema_has_pile_logistics_columns(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh_pile_logistics.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        offers = _table_columns(cur, "KP_offers")
        meta = _table_columns(cur, "kp_meta")
        assert "pile_logistics_cost" in offers
        assert "pile_trip_overrides_json" in meta
        cur.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) VALUES ('2026-01-01', 'x')"
        )
        kp_id = cur.lastrowid
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (?, 'в работе')",
            (kp_id,),
        )
        conn.commit()
        pile_cost = cur.execute(
            "SELECT pile_logistics_cost FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
        overrides = cur.execute(
            "SELECT pile_trip_overrides_json FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()[0]
    assert pile_cost in (0, 0.0, None)
    assert overrides is None


def test_migrate_existing_db_adds_pile_logistics_columns(tmp_path: Path) -> None:
    db_path = str(tmp_path / "legacy_pile_logistics.db")
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL UNIQUE,
                status TEXT DEFAULT 'в работе',
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        conn.execute("INSERT INTO KP_offers (creation_date) VALUES ('2026-01-01')")
        conn.execute("INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')")
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        offers = _table_columns(cur, "KP_offers")
        meta = _table_columns(cur, "kp_meta")
        assert "pile_logistics_cost" in offers
        assert "pile_trip_overrides_json" in meta
        row = cur.execute(
            "SELECT COALESCE(pile_logistics_cost, 0), pile_trip_overrides_json "
            "FROM KP_offers o JOIN kp_meta m ON m.kp_id = o.kp_id WHERE o.kp_id = 1"
        ).fetchone()
    assert row[0] == 0
    assert row[1] is None
