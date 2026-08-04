"""Archive shipped_progress read-path (audit Q4/Q5)."""

from __future__ import annotations

import sqlite3

from app.repositories.kp_archive_repository import KpArchiveRepository
from app.services.archive_service import ArchiveService
from tests.helpers import kp_db_fixtures as fx


def test_shipped_progress_null_when_ordered_qty_missing(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1, status="в работе")
    repo = KpArchiveRepository(db_path=db)
    service = ArchiveService(repository=repo, outputs_dir=tmp_path)

    assert service._shipped_progress(1) is None

    with sqlite3.connect(db) as conn:
        ordered = conn.execute(
            "SELECT ordered_qty FROM kp_meta WHERE kp_id = 1"
        ).fetchone()[0]
    assert ordered is None


def test_shipped_progress_uses_frozen_m(tmp_path) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1, status="в работе")
    with sqlite3.connect(db) as conn:
        conn.execute("UPDATE kp_meta SET ordered_qty = 10 WHERE kp_id = 1")
        conn.commit()

    repo = KpArchiveRepository(db_path=db)
    service = ArchiveService(repository=repo, outputs_dir=tmp_path)
    progress = service._shipped_progress(1)
    assert progress == {"x": 0, "m": 10}
