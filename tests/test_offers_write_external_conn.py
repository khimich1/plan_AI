"""MTP-100: _external_conn on update_kp_status / update_kp_execution_date."""

from __future__ import annotations

import sqlite3

from core.kp import offers_write
from core.kp_db_common import _connect
from tests.helpers.kp_db_fixtures import make_iso_db, seed_kp_offer


def _read_state(db_path: str, kp_id: int) -> tuple[str | None, str | None]:
    with sqlite3.connect(db_path) as conn:
        status = conn.execute(
            "SELECT status FROM kp_meta WHERE kp_id = ?", (kp_id,)
        ).fetchone()
        terms = conn.execute(
            "SELECT execution_terms FROM KP_offers WHERE kp_id = ?", (kp_id,)
        ).fetchone()
    return (status[0] if status else None, terms[0] if terms else None)


def test_update_with_external_conn_does_not_commit(tmp_path) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 1, status="в архиве")

    conn = _connect(db_path)
    try:
        assert offers_write.update_kp_execution_date(
            1, "15.08.2026", db_path, _external_conn=conn
        )
        assert offers_write.update_kp_status(
            1, "в работе", db_path, _external_conn=conn
        )
        # Uncommitted writes on this connection; other connections see old state.
        other_status, other_terms = _read_state(db_path, 1)
        assert other_status == "в архиве"
        assert other_terms == "21.04.2026"

        conn.rollback()
    finally:
        conn.close()

    status, terms = _read_state(db_path, 1)
    assert status == "в архиве"
    assert terms == "21.04.2026"


def test_update_without_external_conn_commits(tmp_path) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 2, status="в архиве")

    assert offers_write.update_kp_execution_date(2, "01.09.2026", db_path)
    assert offers_write.update_kp_status(2, "в работе", db_path)

    status, terms = _read_state(db_path, 2)
    assert status == "в работе"
    assert terms == "01.09.2026"
