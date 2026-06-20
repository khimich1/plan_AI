"""Integration tests for ``/api/v1/production`` (Q5).

Covers happy-path flows through FastAPI TestClient with session auth and
failure modes: 401/403, 404, 422 validation, 409 plan_version_conflict,
day_already_completed, and structured business errors.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.repositories.plan_errors import PlanVersionConflict
from app.repositories.plan_repository import PlanRepository
from app.schemas.errors import ERROR_CODE_PLAN_VERSION_CONFLICT
from app.security.session import create_session_token
from app.services.production_completion_service import ProductionCompletionError
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_service import ProductionService
from core import kp_db

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
API_PREFIX = "/api/v1/production"
PLATE_NAME = "ПБ 60-12-8п"
KP_ID = 1
DATE_KEY = "2026-04-21"

USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
]


def _session_cookie(user_id: int, role: str, username: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


def _fake_optimize(self, *, orders_2d, **kwargs):
    if not orders_2d:
        return [], {}
    order = orders_2d[0]
    items = [
        {
            "kp_id": order["kp_id"],
            "plate_name": order["plate_name"],
            "length": order["length"],
            "width": order["width"],
            "load_code": order["load_code"],
        }
        for _ in range(int(order.get("qty") or 0))
    ]
    tracks = [{"label": "ОСНОВНАЯ", "items": items}]
    optimization_result = {
        "total_plates": len(items),
        "plate_assignments": [
            {
                "source": "primary",
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            for _ in items
        ],
    }
    return tracks, optimization_result


def _seed_production_db(db_path: str, *, plate_qty: int = 3) -> None:
    kp_db.init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (?, '2026-01-01', '21.04.2026', 'ТестКлиент')",
            (KP_ID,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (?, 'в работе')",
            (KP_ID,),
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (?, 1, ?, 6.0, 1.2, 800, ?, 'в производстве')
            """,
            (KP_ID, PLATE_NAME, plate_qty),
        )
        conn.commit()


def _build_plan_via_api(
    client: TestClient,
    cookies: dict[str, str],
    *,
    tracks_count: int = 1,
) -> dict[str, Any]:
    response = client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": tracks_count,
            "filter_method": "all",
        },
        cookies=cookies,
    )
    assert response.status_code == 200, response.text
    return response.json()


@pytest.fixture()
def production_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> str:
    """Isolated plita DB, work calendar, mocked optimizer."""
    db_path = str(tmp_path / "plita.db")
    _seed_production_db(db_path)

    calendar_path = tmp_path / "work_calendar.json"
    calendar_path.write_text(
        '{"extra_holidays": [], "extra_workdays": []}',
        encoding="utf-8",
    )

    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    monkeypatch.setenv("WORK_CALENDAR_PATH", str(calendar_path))
    get_settings.cache_clear()

    monkeypatch.setattr(AuthRepository, "list_users", lambda self: list(USERS))
    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        _fake_optimize,
    )
    monkeypatch.setattr("core.production.planning.get_reinforcement", lambda **kwargs: 999.0)
    monkeypatch.setattr("core.work_calendar.load_holidays", lambda: set())
    monkeypatch.setattr("core.work_calendar.load_extra_workdays", lambda: set())

    return db_path


@pytest.fixture()
def client(production_env: str) -> TestClient:
    del production_env
    return TestClient(create_app())


@pytest.fixture()
def admin_cookie() -> dict[str, str]:
    return _session_cookie(1, "admin", "admin")


@pytest.fixture()
def production_cookie() -> dict[str, str]:
    return _session_cookie(2, "production", "prod_user")


@pytest.fixture()
def manager_cookie() -> dict[str, str]:
    return _session_cookie(3, "manager", "manager_a")


@pytest.fixture()
def built_plan(client: TestClient, admin_cookie: dict[str, str]) -> dict[str, Any]:
    return _build_plan_via_api(client, admin_cookie)


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


def test_list_plans_requires_auth(client: TestClient) -> None:
    response = client.get(f"{API_PREFIX}/plans")
    assert response.status_code == 401


def test_manager_role_forbidden(client: TestClient, manager_cookie: dict[str, str]) -> None:
    response = client.get(f"{API_PREFIX}/plans", cookies=manager_cookie)
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_production_role_allowed(client: TestClient, production_cookie: dict[str, str]) -> None:
    response = client.get(f"{API_PREFIX}/plans", cookies=production_cookie)
    assert response.status_code == 200


# ---------------------------------------------------------------------------
# Happy path — plans CRUD / build
# ---------------------------------------------------------------------------


def test_list_plans_empty_then_after_build(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]

    response = client.get(f"{API_PREFIX}/plans", cookies=admin_cookie)
    assert response.status_code == 200
    payload = response.json()
    assert any(item["id"] == plan_id for item in payload["plans"])


def test_build_plan_returns_version(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan = built_plan["plan"]
    assert plan["id"]
    assert plan.get("version", 0) >= 1
    assert built_plan["summary"]["selected_plates_count"] == 3


def test_get_plan_includes_version(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]

    response = client.get(f"{API_PREFIX}/plans/{plan_id}", cookies=admin_cookie)
    assert response.status_code == 200
    payload = response.json()
    assert payload["id"] == plan_id
    assert payload["version"] >= 1
    assert DATE_KEY in payload.get("days", {})


def test_activate_and_delete_plan(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]

    activate = client.post(f"{API_PREFIX}/plans/{plan_id}/activate", cookies=admin_cookie)
    assert activate.status_code == 200
    assert activate.json() == {"plan_id": plan_id, "active": True}

    listed = client.get(f"{API_PREFIX}/plans", cookies=admin_cookie).json()
    assert listed["active_plan_id"] == plan_id

    deleted = client.delete(f"{API_PREFIX}/plans/{plan_id}", cookies=admin_cookie)
    assert deleted.status_code == 200
    assert deleted.json() == {"plan_id": plan_id, "deleted": True}

    missing = client.get(f"{API_PREFIX}/plans/{plan_id}", cookies=admin_cookie)
    assert missing.status_code == 404


def test_get_plan_not_found(client: TestClient, admin_cookie: dict[str, str]) -> None:
    response = client.get(f"{API_PREFIX}/plans/nonexistent_plan", cookies=admin_cookie)
    assert response.status_code == 404


def test_day_occupancy_and_kp_candidates(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]

    occupancy = client.get(f"{API_PREFIX}/day-occupancy", cookies=admin_cookie)
    assert occupancy.status_code == 200
    occ_payload = occupancy.json()
    assert occ_payload["max_per_day"] >= 1
    assert DATE_KEY in occ_payload["occupancy"]

    excluded = client.get(
        f"{API_PREFIX}/day-occupancy",
        params={"exclude_plan_id": plan_id},
        cookies=admin_cookie,
    )
    assert excluded.status_code == 200

    candidates = client.get(f"{API_PREFIX}/kp-candidates", cookies=admin_cookie)
    assert candidates.status_code == 200
    assert "items" in candidates.json()


def test_get_day_view_after_build(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]

    response = client.get(f"{API_PREFIX}/days/{DATE_KEY}", cookies=admin_cookie)
    assert response.status_code == 200
    payload = response.json()
    assert payload["date"] == DATE_KEY
    plan_blocks = [block for block in payload["plans"] if block["plan_id"] == plan_id]
    assert plan_blocks, "built plan should appear in day view"


def test_remove_track_happy_path(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    production_env: str,
) -> None:
    plan_id = built_plan["plan"]["id"]

    response = client.delete(
        f"{API_PREFIX}/plans/{plan_id}/days/{DATE_KEY}/tracks/0",
        cookies=admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["plan_id"] == plan_id
    assert payload["date"] == DATE_KEY
    assert payload["track_index"] == 0
    assert payload["plates_returned"] >= 1

    plan = PlanRepository(db_path=production_env).get(plan_id)
    assert plan is not None
    day = plan["payload"]["days"].get(DATE_KEY, {})
    assert len(day.get("tracks") or []) == 0


def test_complete_day_happy_path(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    production_env: str,
) -> None:
    plan_id = built_plan["plan"]["id"]

    response = client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["completed"] is True
    assert payload["moved_plates"] == 3

    with sqlite3.connect(production_env) as conn:
        completed_qty = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
            (KP_ID,),
        ).fetchone()[0]
    assert completed_qty == 3


def test_work_calendar_round_trip(
    client: TestClient,
    admin_cookie: dict[str, str],
) -> None:
    save = client.put(
        f"{API_PREFIX}/work-calendar",
        json={"extra_holidays": ["2026-05-01"], "extra_workdays": []},
        cookies=admin_cookie,
    )
    assert save.status_code == 200
    assert save.json()["extra_holidays"] == ["2026-05-01"]

    load = client.get(f"{API_PREFIX}/work-calendar", cookies=admin_cookie)
    assert load.status_code == 200
    assert load.json()["extra_holidays"] == ["2026-05-01"]


# ---------------------------------------------------------------------------
# Failure modes
# ---------------------------------------------------------------------------


def test_build_plan_validation_error(client: TestClient, admin_cookie: dict[str, str]) -> None:
    response = client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 0,
            "filter_method": "all",
        },
        cookies=admin_cookie,
    )
    assert response.status_code == 422


def test_build_plan_business_error_when_no_plates(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    admin_cookie: dict[str, str],
) -> None:
    db_path = str(tmp_path / "empty.db")
    kp_db.init_schema(db_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    monkeypatch.setattr(AuthRepository, "list_users", lambda self: list(USERS))
    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        _fake_optimize,
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
        cookies=admin_cookie,
    )
    assert response.status_code == 422
    assert isinstance(response.json()["detail"], str)


def test_complete_day_validation_error(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
) -> None:
    plan_id = built_plan["plan"]["id"]
    response = client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={
            "plan_id": plan_id,
            "rejected_plates": [{"track_number": 0, "plate_index": 0, "qty": 1}],
        },
        cookies=admin_cookie,
    )
    assert response.status_code == 422


def test_remove_track_day_already_completed(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    production_env: str,
) -> None:
    plan_id = built_plan["plan"]["id"]
    repo = PlanRepository(db_path=production_env)
    record = repo.get(plan_id)
    assert record is not None
    plan = dict(record["payload"])
    plan["days"][DATE_KEY]["completed"] = True
    repo.save(plan, expected_version=record["version"])

    response = client.delete(
        f"{API_PREFIX}/plans/{plan_id}/days/{DATE_KEY}/tracks/0",
        cookies=admin_cookie,
    )
    assert response.status_code == 409
    assert "заверш" in response.json()["detail"].lower()


def test_complete_day_plan_version_conflict(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = built_plan["plan"]["id"]
    original_mark = PlanRepository.mark_day_completed

    def raise_conflict(self, pid: str, target_date: str, expected_version: int | None = None):
        raise PlanVersionConflict(pid, expected_version or 1)

    monkeypatch.setattr(PlanRepository, "mark_day_completed", raise_conflict)

    response = client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=admin_cookie,
    )

    monkeypatch.setattr(PlanRepository, "mark_day_completed", original_mark)

    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT
    assert detail["details"]["plan_id"] == plan_id


def test_remove_track_plan_version_conflict(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = built_plan["plan"]["id"]

    def raise_conflict(self, pid: str, date: str, track_index: int, **kwargs):
        raise PlanVersionConflict(pid, 1)

    monkeypatch.setattr(PlanRepository, "remove_track_from_plan", raise_conflict)

    response = client.delete(
        f"{API_PREFIX}/plans/{plan_id}/days/{DATE_KEY}/tracks/0",
        cookies=admin_cookie,
    )
    assert response.status_code == 409
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT


def test_build_plan_version_conflict(
    client: TestClient,
    admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def raise_on_build(self, **kwargs):
        raise PlanVersionConflict("plan_stale", 2)

    monkeypatch.setattr(ProductionPlanningService, "build_plan", raise_on_build)

    response = client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "all",
        },
        cookies=admin_cookie,
    )
    assert response.status_code == 409
    assert response.json()["detail"]["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT


def test_complete_day_business_error_maps_to_422(
    client: TestClient,
    admin_cookie: dict[str, str],
    built_plan: dict[str, Any],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan_id = built_plan["plan"]["id"]

    def raise_completion_error(self, **kwargs):
        raise ProductionCompletionError("не списано 3 плит")

    monkeypatch.setattr(ProductionService, "complete_day", raise_completion_error)

    response = client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=admin_cookie,
    )
    assert response.status_code == 422
    assert response.json()["detail"] == "Не удалось выполнить операцию. Проверьте введённые данные."


def test_get_day_view_not_found(client: TestClient, admin_cookie: dict[str, str]) -> None:
    response = client.get(f"{API_PREFIX}/days/2099-01-01", cookies=admin_cookie)
    assert response.status_code == 404
