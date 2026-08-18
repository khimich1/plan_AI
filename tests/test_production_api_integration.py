"""Integration tests for ``/api/v1/production`` (WP3 / Q-M9).

Covers happy-path flows through FastAPI TestClient with session auth and
failure modes: 401 without cookie, 403 for disallowed roles, 409 plan_version_conflict.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.repositories.plan_errors import PlanVersionConflict
from app.repositories.plan_repository import PlanRepository
from app.schemas.errors import ERROR_CODE_PLAN_VERSION_CONFLICT
from app.services.production_completion_service import ProductionCompletionError
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_service import ProductionService
from core.production_capacity import TRACKS_PER_DAY_DEFAULT
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


def test_plan_lifecycle_create_build_activate_orchestration(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """WP3/A10: build (create) → activate проходит через ProductionPlanningService."""
    build_calls: list[str] = []
    original_build = ProductionPlanningService.build_plan

    def tracked_build(self, **kwargs):
        build_calls.append("build_plan")
        return original_build(self, **kwargs)

    monkeypatch.setattr(ProductionPlanningService, "build_plan", tracked_build)

    built = paf.build_plan_via_api(production_api_client, production_admin_cookie)
    assert build_calls == ["build_plan"]
    assert built["stats"]["is_new_plan"] is True

    plan_id = built["plan"]["id"]
    activate = production_api_client.post(
        f"{API_PREFIX}/plans/{plan_id}/activate",
        cookies=production_admin_cookie,
    )
    assert activate.status_code == 200

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


def test_get_day_capacity_range_defaults(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/day-capacity",
        params={"from": "2026-08-01", "to": "2026-08-03"},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    capacity = response.json()["capacity"]
    assert set(capacity) == {"2026-08-01", "2026-08-02", "2026-08-03"}
    expected = int(TRACKS_PER_DAY_DEFAULT)
    assert all(v == expected for v in capacity.values())


def test_put_day_capacity_and_get(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    before = production_api_client.get(
        f"{API_PREFIX}/day-capacity",
        params={"from": "2026-08-12", "to": "2026-08-12"},
        cookies=production_admin_cookie,
    )
    assert before.status_code == 200
    assert before.json()["capacity"]["2026-08-12"] == int(TRACKS_PER_DAY_DEFAULT)

    save = production_api_client.put(
        f"{API_PREFIX}/day-capacity",
        json={"date": "2026-08-12", "max_tracks": 3},
        cookies=production_admin_cookie,
    )
    assert save.status_code == 200
    assert save.json() == {"date": "2026-08-12", "max_tracks": 3}

    load = production_api_client.get(
        f"{API_PREFIX}/day-capacity",
        params={"from": "2026-08-12", "to": "2026-08-12"},
        cookies=production_admin_cookie,
    )
    assert load.status_code == 200
    assert load.json()["capacity"]["2026-08-12"] == 3


def test_put_day_capacity_rejects_negative(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.put(
        f"{API_PREFIX}/day-capacity",
        json={"date": "2026-08-12", "max_tracks": -1},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422


def test_put_day_capacity_rejects_above_hard_cap(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.put(
        f"{API_PREFIX}/day-capacity",
        json={"date": "2026-08-12", "max_tracks": 6},
        cookies=production_admin_cookie,
    )
    # Schema le=TRACKS_PER_DAY_HARD_CAP → 422 at Pydantic boundary (was 400 from service).
    assert response.status_code == 422
    assert "5" in response.text


def test_get_day_occupancy_max_by_day_respects_override(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    """Override max_tracks=3 → max_by_day[date]=3; max_per_day stays default (5)."""
    from datetime import datetime, timedelta
    from zoneinfo import ZoneInfo

    target = (
        datetime.now(ZoneInfo("Europe/Moscow")).date() + timedelta(days=3)
    ).isoformat()

    save = production_api_client.put(
        f"{API_PREFIX}/day-capacity",
        json={"date": target, "max_tracks": 3},
        cookies=production_admin_cookie,
    )
    assert save.status_code == 200

    response = production_api_client.get(
        f"{API_PREFIX}/day-occupancy",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["max_per_day"] == int(TRACKS_PER_DAY_DEFAULT)
    assert "max_by_day" in payload
    assert payload["max_by_day"][target] == 3
    assert payload["max_by_day"][target] != payload["max_per_day"]


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

    def raise_conflict(
        self,
        pid: str,
        target_date: str,
        expected_version: int | None = None,
        **_kwargs,
    ):
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

    client = CsrfAwareTestClient(create_app())
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

    from app.services.plan_distribution_service import PlanDistributionService

    def raise_conflict(self, repo, pid: str, date: str, track_index: int, **kwargs):
        raise PlanVersionConflict(pid, 1)

    monkeypatch.setattr(PlanDistributionService, "remove_track_from_plan", raise_conflict)

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


# ---------------------------------------------------------------------------
# POST /analyze-substrates (orch-2026-08-12-podlozhki Task 8)
# ---------------------------------------------------------------------------


def test_analyze_substrates_manager_forbidden(
    production_api_client: TestClient,
    production_manager_cookie: dict[str, str],
) -> None:
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_manager_cookie,
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_analyze_substrates_deadline_before_first_fill(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": "2026-04-21", "tracks": 1}],
            "deadline_until": "2026-04-20",
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422
    assert "deadline_until" in response.json()["detail"]


def test_analyze_substrates_empty_backlog(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    production_admin_cookie: dict[str, str],
) -> None:
    db_path = str(tmp_path / "empty_analyze.db")
    from core import kp_db

    kp_db.init_schema(db_path)
    monkeypatch.setenv("APP_SECRET_KEY", paf.VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, list(paf.TEST_USERS))
    monkeypatch.setattr("core.work_calendar.load_holidays", lambda: set())
    monkeypatch.setattr("core.work_calendar.load_extra_workdays", lambda: set())

    client = CsrfAwareTestClient(create_app())
    response = client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422
    assert "производств" in response.json()["detail"].lower()


def test_analyze_substrates_happy_path(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sqlite3

    with sqlite3.connect(production_api_db) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (2, '2026-01-02', '05.09.2026', 'ПозднийКлиент')"
        )
        conn.execute("INSERT INTO kp_meta (kp_id, status) VALUES (2, 'в работе')")
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (2, 1, 'ПБ 57-4,8', 5.7, 0.48, 800, 2, 'в производстве')
            """
        )
        # Align urgent plate name/size with mocked optimizer primary cut.
        conn.execute(
            "UPDATE kp_plates SET plate_name = ?, length_m = ?, width_m = ? "
            "WHERE kp_id = 1",
            ("ПБ 57-7,2", 5.7, 0.72),
        )
        conn.commit()

        urgent_plate_id = int(
            conn.execute("SELECT id FROM kp_plates WHERE kp_id = 1").fetchone()[0]
        )
        late_plate_id = int(
            conn.execute("SELECT id FROM kp_plates WHERE kp_id = 2").fetchone()[0]
        )

    def fake_optimize(*, orders_2d, **kwargs):
        return {
            "_opt_status": "ok",
            "primary_cuts": [
                {
                    "primary_instance_id": "prim-1",
                    "kp_id": 1,
                    "plate_name": "ПБ 57-7,2",
                    "rest": 480,
                    "lengths": [5.7],
                    "width": 720,
                }
            ],
            "secondary_cuts": [
                {
                    "parent_instance_id": "prim-1",
                    "kp_id": 2,
                    "plate_name": "ПБ 57-4,8",
                    "cuts": [480],
                    "qty": 1,
                    "lengths": [5.7],
                }
            ],
        }

    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        fake_optimize,
    )

    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [
                {"date": "2026-04-20", "tracks": 5},
                {"date": "2026-04-21", "tracks": 3},
            ],
            "deadline_until": "2026-04-25",
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    payload = response.json()
    assert set(payload) == {
        "urgent_positions",
        "substrate_recommendations",
        "capacity_deficit",
        "analysis_meta",
    }
    assert payload["analysis_meta"]["optimization_status"] == "ok"
    assert payload["analysis_meta"]["orders_count"] == 2
    assert isinstance(payload["analysis_meta"]["analysis_duration_ms"], int)
    assert payload["analysis_meta"]["error_message"] is None
    # 5+3 tracks cover urgent length (~17 m) → no deficit
    assert payload["capacity_deficit"] is None

    urgent_ids = {item["plate_id"] for item in payload["urgent_positions"]}
    assert urgent_plate_id in urgent_ids

    assert len(payload["substrate_recommendations"]) == 1
    rec = payload["substrate_recommendations"][0]
    assert rec["plate_id"] == late_plate_id
    assert rec["kp_id"] == 2
    assert rec["under_plate_id"] == urgent_plate_id
    assert rec["under_kp_id"] == 1
    assert rec["saving_mm"] == 480
    assert rec["saving_m"] == pytest.approx(480 * 5.7 / 1000.0)


def test_analyze_substrates_capacity_deficit_present(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from core.production.capacity import CapacityDeficit, CapacityOption

    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        lambda **kwargs: {"_opt_status": "ok", "primary_cuts": [], "secondary_cuts": []},
    )
    monkeypatch.setattr(
        "app.services.production_service.calculate_capacity_deficit",
        lambda *args, **kwargs: CapacityDeficit(
            tracks_needed=5,
            tracks_available=1,
            tracks_missing=4,
            deficit_until=DATE_KEY,
            options=(
                CapacityOption(
                    action="bump_fill",
                    date=DATE_KEY,
                    add_tracks=4,
                    free=5,
                ),
            ),
        ),
    )
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    deficit = response.json()["capacity_deficit"]
    assert deficit is not None
    assert set(deficit) == {
        "tracks_needed",
        "tracks_available",
        "tracks_missing",
        "deficit_until",
        "options",
    }
    assert deficit["tracks_missing"] == 4
    assert deficit["options"] == [
        {"action": "bump_fill", "date": DATE_KEY, "add_tracks": 4, "free": 5}
    ]

def test_analyze_substrates_uses_run_cpu_bound(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import app.api.v1.endpoints.production as production_endpoint

    calls: list[object] = []
    real_run = production_endpoint.run_cpu_bound

    async def spy_run_cpu_bound(fn, **kwargs):
        calls.append(fn)
        return await real_run(fn, **kwargs)

    monkeypatch.setattr(production_endpoint, "run_cpu_bound", spy_run_cpu_bound)
    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        lambda **kwargs: {"_opt_status": "ok", "primary_cuts": [], "secondary_cuts": []},
    )
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    assert len(calls) == 1


def test_analyze_substrates_optimizer_error_returns_error_message(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """ProductionSubstrateError → HTTP 200 + analysis_meta.error_message."""

    def fake_optimize(*, orders_2d, **kwargs):
        return {
            "_opt_status": "error",
            "_opt_error_message": "infeasible",
            "primary_cuts": [],
            "secondary_cuts": [],
        }

    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        fake_optimize,
    )
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    payload = response.json()
    assert payload["substrate_recommendations"] == []
    assert payload["analysis_meta"]["optimization_status"] == "error"
    assert payload["analysis_meta"]["error_message"]
    assert "подлож" in payload["analysis_meta"]["error_message"].lower()


def test_analyze_substrates_production_role_allowed(
    production_api_client: TestClient,
    production_user_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        lambda **kwargs: {"_opt_status": "ok", "primary_cuts": [], "secondary_cuts": []},
    )
    response = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": DATE_KEY, "tracks": 1}],
            "deadline_until": DATE_KEY,
        },
        cookies=production_user_cookie,
    )
    assert response.status_code == 200, response.text
    assert "urgent_positions" in response.json()
