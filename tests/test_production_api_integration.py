"""Integration tests for ``/api/v1/production`` (WP3 / Q-M9).

Covers happy-path flows through FastAPI TestClient with session auth and
failure modes: 401 without cookie, 403 for disallowed roles, 409 plan_version_conflict.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.repositories.plan_errors import PlanVersionConflict
from app.repositories.plan_repository import PlanRepository
from app.schemas.errors import ERROR_CODE_PLAN_VERSION_CONFLICT
from app.services.production_completion_service import ProductionCompletionError
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_service import ProductionService
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX
DATE_KEY = paf.DATE_KEY


# ---------------------------------------------------------------------------
# Auth / RBAC failure modes
# ---------------------------------------------------------------------------


def test_list_plans_requires_auth(production_api_client: TestClient) -> None:
    response = production_api_client.get(f"{API_PREFIX}/plans")
    assert response.status_code == 401


def test_manager_role_forbidden(
    production_api_client: TestClient,
    production_manager_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/plans",
        cookies=production_manager_cookie,
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_production_role_allowed(
    production_api_client: TestClient,
    production_user_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/plans",
        cookies=production_user_cookie,
    )
    assert response.status_code == 200


# ---------------------------------------------------------------------------
# Happy path — minimum 8 production routes
# ---------------------------------------------------------------------------


def test_get_plans(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    response = production_api_client.get(f"{API_PREFIX}/plans", cookies=production_admin_cookie)
    assert response.status_code == 200
    payload = response.json()
    assert any(item["id"] == plan_id for item in payload["plans"])


def test_get_plan_by_id(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    response = production_api_client.get(
        f"{API_PREFIX}/plans/{plan_id}",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["id"] == plan_id
    assert payload["version"] >= 1
    assert DATE_KEY in payload.get("days", {})


def test_post_plans_build(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan = production_built_plan["plan"]
    assert plan["id"]
    assert plan.get("version", 0) >= 1
    assert production_built_plan["summary"]["selected_plates_count"] == 3


def test_post_plan_activate(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    response = production_api_client.post(
        f"{API_PREFIX}/plans/{plan_id}/activate",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    assert response.json() == {"plan_id": plan_id, "active": True}

    listed = production_api_client.get(
        f"{API_PREFIX}/plans",
        cookies=production_admin_cookie,
    ).json()
    assert listed["active_plan_id"] == plan_id


def test_get_kp_candidates(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/kp-candidates",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    assert "items" in response.json()


def test_get_day_view(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    response = production_api_client.get(
        f"{API_PREFIX}/days/{DATE_KEY}",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["date"] == DATE_KEY
    plan_blocks = [block for block in payload["plans"] if block["plan_id"] == plan_id]
    assert plan_blocks, "built plan should appear in day view"


def test_get_work_calendar(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/work-calendar",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    assert "extra_holidays" in response.json()


def test_put_work_calendar(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    save = production_api_client.put(
        f"{API_PREFIX}/work-calendar",
        json={"extra_holidays": ["2026-05-01"], "extra_workdays": []},
        cookies=production_admin_cookie,
    )
    assert save.status_code == 200
    assert save.json()["extra_holidays"] == ["2026-05-01"]

    load = production_api_client.get(
        f"{API_PREFIX}/work-calendar",
        cookies=production_admin_cookie,
    )
    assert load.status_code == 200
    assert load.json()["extra_holidays"] == ["2026-05-01"]


# ---------------------------------------------------------------------------
# Failure modes — 409 plan version conflict (+ validation / business errors)
# ---------------------------------------------------------------------------


def test_build_plan_version_conflict(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def raise_on_build(self, **kwargs):
        raise PlanVersionConflict("plan_stale", 2)

    monkeypatch.setattr(ProductionPlanningService, "build_plan", raise_on_build)

    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "all",
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 409
    assert response.json()["detail"]["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT


def test_complete_day_plan_version_conflict(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = production_built_plan["plan"]["id"]
    original_mark = PlanRepository.mark_day_completed

    def raise_conflict(self, pid: str, target_date: str, expected_version: int | None = None):
        raise PlanVersionConflict(pid, expected_version or 1)

    monkeypatch.setattr(PlanRepository, "mark_day_completed", raise_conflict)

    response = production_api_client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=production_admin_cookie,
    )

    monkeypatch.setattr(PlanRepository, "mark_day_completed", original_mark)

    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT
    assert detail["details"]["plan_id"] == plan_id


def test_build_plan_validation_error(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 0,
            "filter_method": "all",
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422


def test_build_plan_business_error_when_no_plates(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    production_admin_cookie: dict[str, str],
) -> None:
    db_path = str(tmp_path / "empty.db")
    from core import kp_db

    kp_db.init_schema(db_path)
    monkeypatch.setenv("APP_SECRET_KEY", paf.VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, list(paf.TEST_USERS))
    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        paf.fake_optimize,
    )
    monkeypatch.setattr("core.work_calendar.load_holidays", lambda: set())
    monkeypatch.setattr("core.work_calendar.load_extra_workdays", lambda: set())

    client = TestClient(create_app())
    response = client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "all",
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422
    assert isinstance(response.json()["detail"], str)


def test_get_plan_not_found(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/plans/nonexistent_plan",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 404


def test_activate_and_delete_plan(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    deleted = production_api_client.delete(
        f"{API_PREFIX}/plans/{plan_id}",
        cookies=production_admin_cookie,
    )
    assert deleted.status_code == 200
    assert deleted.json() == {"plan_id": plan_id, "deleted": True}

    missing = production_api_client.get(
        f"{API_PREFIX}/plans/{plan_id}",
        cookies=production_admin_cookie,
    )
    assert missing.status_code == 404


def test_remove_track_plan_version_conflict(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    def raise_conflict(self, pid: str, date: str, track_index: int, **kwargs):
        raise PlanVersionConflict(pid, 1)

    monkeypatch.setattr(PlanRepository, "remove_track_from_plan", raise_conflict)

    response = production_api_client.delete(
        f"{API_PREFIX}/plans/{plan_id}/days/{DATE_KEY}/tracks/0",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT


def test_complete_day_business_error_maps_to_422(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = production_built_plan["plan"]["id"]

    def raise_completion_error(self, **kwargs):
        raise ProductionCompletionError("не списано 3 плит")

    monkeypatch.setattr(ProductionService, "complete_day", raise_completion_error)

    response = production_api_client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422
    assert response.json()["detail"] == "Не удалось выполнить операцию. Проверьте введённые данные."
