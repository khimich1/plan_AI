"""BP-002: kp_bridge_piles table (with concrete_grade)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from core import kp_db_schema


def test_fresh_schema_has_kp_bridge_piles(tmp_path: Path) -> None:
    db_path = str(tmp_path / "fresh.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(kp_bridge_piles)")
        cols = {row[1]: row for row in cur.fetchall()}
        assert cols.keys() >= {
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
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (1, '2026-08-05')"
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, product_type) "
            "VALUES (1, 'в архиве', 'bridge_piles')"
        )
        cur.execute(
            """
            INSERT INTO kp_bridge_piles (
                kp_id, position_number, mark, concrete_grade,
                qty, unit_price, discounted_price
            ) VALUES (1, 1, 'C8-35В4', 'B25', 2, 49813.83, 49813.83)
            """
        )
        conn.commit()

        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "bridge_piles"

        cur.execute(
            "SELECT mark, concrete_grade, qty FROM kp_bridge_piles WHERE kp_id = 1"
        )
        assert cur.fetchone() == ("C8-35В4", "B25", 2)

        # Must not write into kp_piles
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = 1")
        assert cur.fetchone()[0] == 0


def test_ensure_schema_idempotent_with_kp_bridge_piles(tmp_path: Path) -> None:
    db_path = str(tmp_path / "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='kp_bridge_piles'"
        )
        assert cur.fetchone() is not None
