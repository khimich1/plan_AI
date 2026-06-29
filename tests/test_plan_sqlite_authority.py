"""Cross-surface integration: API and bot adapter share SQLite plan authority (WP1 / A1)."""
from __future__ import annotations

from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.planning import plan_manager, plan_storage
from app.repositories.plan_repository import PlanRepository
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX


def test_api_write_readable_via_plan_storage(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    """Запись через API → чтение через plan_storage (bot/web shared layer)."""
    plan_id = production_built_plan["plan"]["id"]

    loaded = plan_storage.load_plan(plan_id)
    assert loaded is not None
    assert loaded["id"] == plan_id
    assert loaded["days"] == production_built_plan["plan"]["days"]


def test_plan_manager_save_readable_via_repository(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Bot path: plan_manager.save_plan → PlanRepository.get."""
    settings = get_settings()
    db_path = str(settings.plita_db_path)
    repo = PlanRepository(db_path=db_path)
    monkeypatch.setattr(plan_storage, "_repo_override", repo)

    plan = {
        "id": "plan_bot_adapter_test",
        "name": "Bot adapter test",
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

    assert plan_manager.save_plan(plan) is True

    record = repo.get("plan_bot_adapter_test")
    assert record is not None
    assert record["payload"]["name"] == "Bot adapter test"
    assert record["version"] == 1


def test_api_list_metadata_matches_repository(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    """Метаданные API и plan_storage.load_plans_metadata согласованы."""
    plan_id = production_built_plan["plan"]["id"]

    response = production_api_client.get(
        f"{API_PREFIX}/plans",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    api_ids = {item["id"] for item in response.json()["plans"]}

    storage_meta = plan_storage.load_plans_metadata()
    storage_ids = {item["id"] for item in storage_meta["plans"]}

    assert plan_id in api_ids
    assert api_ids == storage_ids
