"""STEP-002: kp_steps table (no concrete_grade)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core import kp_db_schema


def test_fresh_schema_has_kp_steps(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(kp_steps)")
        step_cols = {row[1]: row for row in cur.fetchall()}
        assert step_cols.keys() >= {
            "id",
            "kp_id",
            "position_number",
            "mark",
            "qty",
            "unit_price",
            "discounted_price",
        }
        assert "concrete_grade" not in step_cols

        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-08-05')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) VALUES (1, 'в архиве', 'steps')"
        )
        cur.execute(
            """
            INSERT INTO kp_steps (
                kp_id, position_number, mark,
                qty, unit_price, discounted_price
            ) VALUES (1, 1, 'ЛС11', 5, 1409.91, 1409.91)
            """
        )
        conn.commit()

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "steps"

        cur.execute("SELECT mark, qty FROM kp_steps WHERE kp_id = 1")
        assert cur.fetchone() == ("ЛС11", 5)


def test_migrate_existing_db_adds_kp_steps(tmp_path: Path) -> None:
    db_path = str(tmp_path / "legacy.db")
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
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-01-01')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) VALUES (1, 'в работе', 'plates')"
        )
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='kp_steps'"
        )
        assert cur.fetchone() is not None


def test_ensure_schema_idempotent_with_kp_steps(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
