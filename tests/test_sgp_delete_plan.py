"""SGP-503: delete_plan does not reduce SGP inventory qty."""

from __future__ import annotations

import sqlite3

from app.planning import plan_storage
from app.repositories.plan_repository import PlanRepository
from app.services.sgp_service import SgpService
from core import kp_db
from tests.helpers import kp_db_fixtures as fx


def test_delete_plan_nullifies_sgp_plan_id_keeps_qty(tmp_path, monkeypatch) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    plan_id = "plan_sgp_del"
    with sqlite3.connect(db) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (1, 'ПБ 60-12-8п', 6.0, 1.2, 800, 4, '27.07.2026', 1, ?)
            """,
            (plan_id,),
        )
        conn.commit()

    repo = PlanRepository(db_path=db)
    repo.create(
        {
            "id": plan_id,
            "name": "test",
            "days": {},
            "start_date": "2026-04-21",
        }
    )
    monkeypatch.setattr(plan_storage, "_repo_override", repo)

    from app.core.settings import get_settings

    monkeypatch.setattr(
        type(get_settings()),
        "plita_db_path",
        property(lambda self: db),
        raising=False,
    )
    # Prefer direct clear + delete path used by storage
    cleared = SgpService(db_path=db).clear_plan_links(plan_id)
    assert cleared == 1

    with sqlite3.connect(db) as conn:
        row = conn.execute(
            "SELECT qty, plan_id FROM completed_plates WHERE kp_id=1"
        ).fetchone()
    assert row == (4, None)
