"""SGP-103: orphan pre-flight rejects send when kp_plates qty > day_view."""

from __future__ import annotations

import sqlite3

import pytest

from app.services.production_completion_service import (
    ProductionCompletionError,
    ProductionCompletionService,
)
from core import kp_db
from tests.helpers import kp_db_fixtures as fx


def test_orphan_preflight_blocks_send_without_db_changes(tmp_path, monkeypatch) -> None:
    db = fx.make_iso_db(tmp_path)
    fx.seed_kp_offer(db, 1)
    fx.seed_plate(
        db,
        kp_id=1,
        plate_name="ПБ 60-12-8п",
        length_m=6.0,
        width_m=1.2,
        load_class=800,
        qty=5,
        status="в плане",
        plan_id="plan-orphan",
        day_number=1,
    )

    def fake_day_view(_date: str, **_kwargs):
        return {
            "date": "2026-04-21",
            "plans": [
                {
                    "plan_id": "plan-orphan",
                    "plan_name": "orphan",
                    "completed": False,
                    "tracks": [
                        {
                            "track_number": 1,
                            "plates_info": [
                                {
                                    "plate_name": "ПБ 60-12-8п",
                                    "length_m": 6.0,
                                    "width_mm": 1200,
                                    "qty": 3,  # view shows 3, DB has 5 → orphan 2
                                    "load_code": 8,
                                    "load_class": 800,
                                    "kp_id": 1,
                                }
                            ],
                        }
                    ],
                }
            ],
            "plans_count": 1,
            "total_tracks": 1,
        }

    monkeypatch.setattr(
        "app.services.production_completion_service.build_day_view_detail",
        fake_day_view,
    )

    class _Repo:
        def load_plan(self, plan_id: str):
            return {
                "id": plan_id,
                "days": {"2026-04-21": {"day_number": 1, "completed": False}},
            }

    svc = ProductionCompletionService(db_path=db, plan_repository=_Repo())  # type: ignore[arg-type]
    before = fx.completed_snapshot(db)

    with pytest.raises(ProductionCompletionError, match="orphan"):
        svc.send_to_sgp(plan_id="plan-orphan", target_date="2026-04-21")

    assert fx.completed_snapshot(db) == before
    with sqlite3.connect(db) as conn:
        qty = conn.execute(
            "SELECT COALESCE(SUM(qty),0) FROM kp_plates WHERE plan_id='plan-orphan'"
        ).fetchone()[0]
    assert int(qty) == 5
