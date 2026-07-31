"""PILE-001: kp_meta.product_type + kp_piles table."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core import kp_db_schema


def test_fresh_schema_has_product_type_and_kp_piles(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(kp_meta)")
        meta_cols = {row[1]: row for row in cur.fetchall()}
        assert "product_type" in meta_cols
        assert meta_cols["product_type"][2].upper() == "TEXT"

        cur.execute("PRAGMA table_info(kp_piles)")
        pile_cols = {row[1]: row for row in cur.fetchall()}
        assert pile_cols.keys() >= {
            "id",
            "kp_id",
            "position_number",
            "mark",
            "concrete_grade",
            "qty",
            "unit_price",
            "discounted_price",
        }

        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-07-30')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) VALUES (1, 'в архиве', 'piles')"
        )
        cur.execute(
            """
            INSERT INTO kp_piles (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price
            ) VALUES (1, 1, 'С120.35-12', 'B25', 5, 44634.03, 44634.03)
            """
        )
        conn.commit()

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "piles"

        cur.execute("SELECT mark, qty FROM kp_piles WHERE kp_id = 1")
        assert cur.fetchone() == ("С120.35-12", 5)


def test_migrate_existing_db_adds_product_type_and_kp_piles(tmp_path: Path) -> None:
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
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-01-01')"
        )
        cur.execute("INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')")
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(kp_meta)")
        assert "product_type" in {row[1] for row in cur.fetchall()}

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        row = cur.fetchone()
        assert row is not None
        assert row[0] in (None, "plates")

        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='kp_piles'"
        )
        assert cur.fetchone() is not None


def test_ensure_schema_idempotent_with_kp_piles(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
