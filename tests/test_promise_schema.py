"""Promise journal schema (Task 1): tables, defaults, idempotent ensure_schema."""

from __future__ import annotations

import shutil
import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema
from core.kp_db_common import DEFAULT_DB, _connect

_TS = "2026-09-03T12:00:00"

PROMISE_TABLES = (
    "kp_promise",
    "kp_promise_alloc",
    "kp_promise_exclusion",
    "notifications",
    "kp_setting",
)

PROMISE_COLS = {
    "id",
    "kp_id",
    "tracks_total",
    "promised_date",
    "kind",
    "status",
    "created_by",
    "created_at",
    "expires_at",
}

ALLOC_COLS = {"id", "promise_id", "week_start", "tracks", "status"}

EXCLUSION_COLS = {
    "id",
    "kp_id",
    "plan_id",
    "week_start",
    "reason",
    "excluded_by",
    "created_at",
}

NOTIFICATION_COLS = {
    "id",
    "user_id",
    "kind",
    "payload_json",
    "read_at",
    "created_at",
}

SETTING_COLS = {"key", "value", "updated_by", "updated_at"}


def _fresh_db(tmp_path: Path, name: str = "promise.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _table_cols(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}


def _seed_kp(conn: sqlite3.Connection, kp_id: int = 1) -> None:
    conn.execute(
        "INSERT INTO KP_offers (kp_id, creation_date) VALUES (?, '2026-09-03')",
        (kp_id,),
    )


def test_ensure_schema_creates_promise_tables(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)

    with _connect(db_path) as conn:
        names = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }
        for table in PROMISE_TABLES:
            assert table in names, f"missing table {table}"


def test_promise_tables_have_expected_columns(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cols.db")

    with _connect(db_path) as conn:
        assert _table_cols(conn, "kp_promise") == PROMISE_COLS
        assert _table_cols(conn, "kp_promise_alloc") == ALLOC_COLS
        assert _table_cols(conn, "kp_promise_exclusion") == EXCLUSION_COLS
        assert _table_cols(conn, "notifications") == NOTIFICATION_COLS
        assert _table_cols(conn, "kp_setting") == SETTING_COLS


def test_kp_setting_seeds_defaults(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "defaults.db")

    with _connect(db_path) as conn:
        rows = {
            row[0]: row[1]
            for row in conn.execute("SELECT key, value FROM kp_setting")
        }
    assert rows["promise_tracks_per_day"] == "3"
    assert rows["promise_buffer"] == "1.0"


def test_ensure_schema_idempotent_with_promise_tables(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)

    with _connect(db_path) as conn:
        assert _table_cols(conn, "kp_promise") == PROMISE_COLS
        rows = {
            row[0]: row[1]
            for row in conn.execute("SELECT key, value FROM kp_setting")
        }
    assert rows["promise_tracks_per_day"] == "3"
    assert rows["promise_buffer"] == "1.0"


def test_kp_setting_defaults_do_not_overwrite_existing(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "no_clobber.db")

    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_setting SET value = '4' WHERE key = 'promise_tracks_per_day'"
        )
        conn.execute(
            "UPDATE kp_setting SET value = '1.15' WHERE key = 'promise_buffer'"
        )
        conn.commit()

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)

    with _connect(db_path) as conn:
        rows = {
            row[0]: row[1]
            for row in conn.execute("SELECT key, value FROM kp_setting")
        }
    assert rows["promise_tracks_per_day"] == "4"
    assert rows["promise_buffer"] == "1.15"


def test_kp_promise_fk_requires_existing_kp(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "fk.db")

    with _connect(db_path) as conn:
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """
                INSERT INTO kp_promise (
                    kp_id, tracks_total, promised_date, kind, status,
                    created_by, created_at
                ) VALUES (999, 2, '2026-09-25', 'promise', 'active', 'alice', ?)
                """,
                (_TS,),
            )


def test_kp_promise_alloc_fk_and_delete_cascade(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "alloc_fk.db")

    with _connect(db_path) as conn:
        _seed_kp(conn)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_promise (
                kp_id, tracks_total, promised_date, kind, status,
                created_by, created_at
            ) VALUES (1, 4, '2026-09-25', 'hold', 'active', 'alice', ?)
            """,
            (_TS,),
        )
        promise_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
            VALUES (?, '2026-09-21', 4, 'active')
            """,
            (promise_id,),
        )
        conn.commit()

        conn.execute("DELETE FROM kp_promise WHERE id = ?", (promise_id,))
        conn.commit()

        leftover = conn.execute(
            "SELECT COUNT(*) FROM kp_promise_alloc WHERE promise_id = ?",
            (promise_id,),
        ).fetchone()[0]
        assert leftover == 0


def test_notifications_user_id_is_soft_reference(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "notif.db")

    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO notifications (user_id, kind, payload_json, created_at)
            VALUES (4242, 'promise_excluded', '{"kp_id": 1}', ?)
            """,
            (_TS,),
        )
        conn.commit()
        row = conn.execute(
            "SELECT user_id, kind, read_at FROM notifications"
        ).fetchone()
    assert row == (4242, "promise_excluded", None)


def test_ensure_schema_on_copied_existing_plita(tmp_path: Path) -> None:
    src = Path(DEFAULT_DB)
    if not src.is_file():
        pytest.skip("plita.db is not present")

    dest = tmp_path / "plita_copy.db"
    shutil.copy2(src, dest)
    db_path = str(dest)

    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)

    with _connect(db_path) as conn:
        names = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }
        for table in PROMISE_TABLES:
            assert table in names
        rows = {
            row[0]: row[1]
            for row in conn.execute("SELECT key, value FROM kp_setting")
        }
    assert "promise_tracks_per_day" in rows
    assert "promise_buffer" in rows
