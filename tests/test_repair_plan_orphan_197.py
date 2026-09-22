"""PLI-011: разовый ремонт сироты id=197 + призрака в JSON плана."""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pytest

from core import kp_db
from core.plan_commit import _verify_plan_integrity
from scripts.repair_plan_orphan_197 import (
    DEFAULT_PLAN_ID,
    RepairConfig,
    parse_args,
    repair_orphan_197,
    tracks_by_day_from_plan,
)

PLATE_NAME = "ПБ 20,6-7,2-8п"


def _ghost_item() -> dict:
    return {
        "length": 2.06,
        "mode": "solid",
        "width": 0.72,
        "load_code": 8,
        "kp_id": 5,
        "plate_name": PLATE_NAME,
        "kp_plate_id": None,
        "concrete_grade": "М400",
    }


def _linked_item(*, kp_id: int, plate_id: int) -> dict:
    return {
        "length": 2.06,
        "mode": "solid",
        "width": 0.72,
        "load_code": 8,
        "kp_id": kp_id,
        "plate_name": PLATE_NAME,
        "kp_plate_id": plate_id,
        "concrete_grade": "М400",
    }


def _plan_payload() -> dict:
    return {
        "id": DEFAULT_PLAN_ID,
        "name": "Фикстура KP4/KP5",
        "start_date": "2026-09-13",
        "days": {
            "2026-09-14": {
                "date": "2026-09-14",
                "day_number": 2,
                "tracks": [
                    {
                        "track_number": 1,
                        "production_day": 2,
                        "items": [
                            _linked_item(kp_id=4, plate_id=160),
                            _linked_item(kp_id=4, plate_id=160),
                            _linked_item(kp_id=5, plate_id=179),
                            _ghost_item(),
                        ],
                    }
                ],
            }
        },
    }


def _seed(db_path: str) -> None:
    kp_db.init_schema(db_path)
    payload = _plan_payload()
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (4, '2026-09-01')"
        )
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date) VALUES (5, '2026-09-01')"
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                id, kp_id, position_number, plate_name, qty, status, plan_id, day_number
            ) VALUES
                (160, 4, 1, ?, 2, 'в плане', ?, 2),
                (197, 4, 2, ?, 1, 'в плане', ?, NULL),
                (179, 5, 1, ?, 1, 'в плане', ?, 2)
            """,
            (PLATE_NAME, DEFAULT_PLAN_ID, PLATE_NAME, DEFAULT_PLAN_ID, PLATE_NAME, DEFAULT_PLAN_ID),
        )
        conn.execute(
            """
            INSERT INTO production_plans (id, payload_json, version, is_active)
            VALUES (?, ?, 1, 1)
            """,
            (DEFAULT_PLAN_ID, json.dumps(payload, ensure_ascii=False)),
        )
        conn.commit()


def _load_plan(db_path: str) -> dict:
    with sqlite3.connect(db_path) as conn:
        raw = conn.execute(
            "SELECT payload_json FROM production_plans WHERE id = ?",
            (DEFAULT_PLAN_ID,),
        ).fetchone()[0]
    return json.loads(raw)


def _row_197(db_path: str) -> tuple:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT kp_id, day_number, plan_id, status FROM kp_plates WHERE id = 197"
        ).fetchone()


def _ghosts(plan: dict) -> list[dict]:
    found = []
    for day in (plan.get("days") or {}).values():
        for track in (day or {}).get("tracks") or []:
            for item in track.get("items") or []:
                if item.get("kp_plate_id") is None:
                    found.append(item)
    return found


@pytest.fixture
def fixture_db(tmp_path: Path) -> str:
    db_path = str(tmp_path / "plita_repair.db")
    _seed(db_path)
    return db_path


def test_cli_default_is_dry_run() -> None:
    args = parse_args([])
    assert args.apply is False


def test_dry_run_prints_plan_and_does_not_mutate(fixture_db: str) -> None:
    report = repair_orphan_197(fixture_db, apply=False)
    assert report.dry_run is True
    assert report.applied is False
    assert report.errors == []
    assert any("197" in line for line in report.changes)
    assert any("kp_id" in line for line in report.changes)

    row = _row_197(fixture_db)
    assert row[1] is None
    ghosts = _ghosts(_load_plan(fixture_db))
    assert len(ghosts) == 1
    assert ghosts[0]["kp_id"] == 5
    assert ghosts[0]["kp_plate_id"] is None


def test_apply_on_copy_fixes_orphan_and_ghost(fixture_db: str) -> None:
    report = repair_orphan_197(fixture_db, apply=True)
    assert report.dry_run is False
    assert report.applied is True
    assert report.errors == []

    row = _row_197(fixture_db)
    assert row[0] == 4
    assert row[1] == 2
    assert _ghosts(_load_plan(fixture_db)) == []
    plan = _load_plan(fixture_db)
    items = plan["days"]["2026-09-14"]["tracks"][0]["items"]
    fixed = [item for item in items if item.get("kp_plate_id") == 197]
    assert len(fixed) == 1
    assert fixed[0]["kp_id"] == 4


def test_gate_observe_clean_after_apply(fixture_db: str) -> None:
    repair_orphan_197(fixture_db, apply=True)
    plan = _load_plan(fixture_db)
    problems = _verify_plan_integrity(
        tracks_by_day=tracks_by_day_from_plan(plan),
        rescue_leftovers={},
        db_path=fixture_db,
        plan_id=DEFAULT_PLAN_ID,
    )
    assert problems == []


def test_apply_is_idempotent(fixture_db: str) -> None:
    repair_orphan_197(fixture_db, apply=True)
    report = repair_orphan_197(fixture_db, apply=True, config=RepairConfig())
    assert report.already_fixed is True
    assert report.applied is False
    row = _row_197(fixture_db)
    assert row[1] == 2
