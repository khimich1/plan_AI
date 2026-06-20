"""Shared helpers for ``/api/v1/production`` integration tests."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient

from app.core.settings import get_settings
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token
from app.services.production_planning_service import ProductionPlanningService
from core import kp_db

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"
API_PREFIX = "/api/v1/production"
PLATE_NAME = "ПБ 60-12-8п"
KP_ID = 1
DATE_KEY = "2026-04-21"

TEST_USERS = [
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


def session_cookie(user_id: int, role: str, username: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


def fake_optimize(self, *, orders_2d, **kwargs):
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


def seed_production_db(db_path: str, *, plate_qty: int = 3) -> None:
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


def configure_production_api_env(
    tmp_path: Path,
    monkeypatch: Any,
    *,
    plate_qty: int = 3,
) -> str:
    """Isolated plita DB, work calendar, mocked optimizer. Returns db_path."""
    db_path = str(tmp_path / "plita.db")
    seed_production_db(db_path, plate_qty=plate_qty)

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

    patch_auth_users(monkeypatch, TEST_USERS)
    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        fake_optimize,
    )
    monkeypatch.setattr("core.production.planning.get_reinforcement", lambda **kwargs: 999.0)
    monkeypatch.setattr("core.work_calendar.load_holidays", lambda: set())
    monkeypatch.setattr("core.work_calendar.load_extra_workdays", lambda: set())

    return db_path


def build_plan_via_api(
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
