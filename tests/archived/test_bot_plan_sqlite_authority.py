"""Cross-surface parity: API ↔ bot.handlers.plan_manager share SQLite authority (WP1 / A1)."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.planning import plan_storage
from app.repositories.plan_repository import PlanRepository
from bot.handlers import plan_manager as bot_plan_manager
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX


def _wire_shared_repository(monkeypatch: pytest.MonkeyPatch, db_path: str) -> PlanRepository:
    """Bot shim and API must use the same isolated SQLite DB."""
    repo = PlanRepository(db_path=db_path)
    monkeypatch.setattr(plan_storage, "_repo_override", repo)
    return repo


def _sample_bot_plan(plan_id: str = "plan_bot_cross_surface") -> dict[str, Any]:
    return {
        "id": plan_id,
        "name": "Bot cross-surface test",
        "created_at": "2026-06-20 12:00:00",
        "start_date": "2026-06-21",
        "tracks_count": 5,
        "days": {
            "2026-06-21": {
                "date": "2026-06-21",
                "day_number": 1,
                "tracks": [{"track_number": 1, "items": []}],
                "saved_tracks_count": 1,
            }
        },
    }


def test_api_write_readable_via_bot_plan_manager(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """API build → bot.handlers.plan_manager.load_plan returns same plan_id + version."""
    _wire_shared_repository(monkeypatch, production_api_db)
    plan_id = production_built_plan["plan"]["id"]
    expected_version = production_built_plan["plan"]["version"]

    loaded = bot_plan_manager.load_plan(plan_id)
    assert loaded is not None
    assert loaded["id"] == plan_id
    assert loaded["days"] == production_built_plan["plan"]["days"]

    repo_record = PlanRepository(db_path=production_api_db).get(plan_id)
    assert repo_record is not None
    assert repo_record["version"] == expected_version


def test_bot_plan_manager_save_readable_via_api(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """bot.handlers.plan_manager.save_plan → API GET returns the same payload + version."""
    _wire_shared_repository(monkeypatch, production_api_db)
    plan = _sample_bot_plan()

    assert bot_plan_manager.save_plan(plan) is True

    response = production_api_client.get(
        f"{API_PREFIX}/plans/{plan['id']}",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    api_plan = response.json()
    assert api_plan["id"] == plan["id"]
    assert api_plan["name"] == plan["name"]
    assert api_plan["days"] == plan["days"]
    assert api_plan["version"] == 1


def test_bot_save_does_not_write_json_plan_files(
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Hot path: bot save must not create bot/data/plans/*.json files."""
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    monkeypatch.setattr(bot_plan_manager, "PLANS_DIR", plans_dir)
    _wire_shared_repository(monkeypatch, production_api_db)

    plan = _sample_bot_plan("plan_no_json_sidecar")
    assert bot_plan_manager.save_plan(plan) is True

    assert list(plans_dir.glob("*.json")) == []

    repo_record = PlanRepository(db_path=production_api_db).get(plan["id"])
    assert repo_record is not None
    assert repo_record["payload"]["name"] == plan["name"]
