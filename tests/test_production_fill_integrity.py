"""Red / contract tests for production fill integrity (audit A1 + API guards + W2).

Wave 1 TDD:
- Occupancy-aware analyze (red until T4): free slots = max − occupied.
- API guards span / hard-cap / candidates limit / ISO path dates (green after T2).

Wave 2 TDD (T5):
- Repeat complete → 409 ``day_already_completed``; ``completed_plates`` do not grow.
- Stale ``expected_version`` on complete → 409; KP not written off.
- Stale ``expected_version`` on DELETE track → 409; track remains.
- SGP ``reserve_on_conn`` fail after build → compensate: no plan, plates not «в плане».
"""
from __future__ import annotations

import sqlite3
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.repositories.plan_repository import PlanRepository
from app.schemas.errors import ERROR_CODE_PLAN_VERSION_CONFLICT
from app.services.sgp_service import SgpError, SgpService
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX
DATE_KEY = paf.DATE_KEY
KP_ID = paf.KP_ID
PLATE_NAME = paf.PLATE_NAME


def _mock_substrate_optimizer(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        lambda **kwargs: {"_opt_status": "ok", "primary_cuts": [], "secondary_cuts": []},
    )


def _analyze(
    client: TestClient,
    cookies: dict[str, str],
    *,
    tracks: int,
    day: str = DATE_KEY,
) -> Any:
    return client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={
            "fill_targets": [{"date": day, "tracks": tracks}],
            "deadline_until": day,
        },
        cookies=cookies,
    )


def _completed_plates_qty(db_path: str, *, kp_id: int = KP_ID) -> int:
    with sqlite3.connect(db_path) as conn:
        return int(
            conn.execute(
                "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
                (kp_id,),
            ).fetchone()[0]
        )


def _kp_plates_by_status(db_path: str, *, kp_id: int = KP_ID) -> dict[str, int]:
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            "SELECT status, COALESCE(SUM(qty), 0) FROM kp_plates "
            "WHERE kp_id = ? GROUP BY status",
            (kp_id,),
        ).fetchall()
    return {str(status): int(qty) for status, qty in rows}


def _seed_free_sgp_row(db_path: str, *, qty: int = 1) -> int:
    """Insert free (kp_id IS NULL) completed_plates row matching fixture plate."""
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (NULL, ?, 6.0, 1.2, 800, ?, '20.04.2026')
            """,
            (PLATE_NAME, qty),
        )
        conn.commit()
        return int(cur.lastrowid)


# ---------------------------------------------------------------------------
# A1 — occupancy-aware analyze (red until T4)
# ---------------------------------------------------------------------------


def test_analyze_rejects_fill_over_free_slots_when_day_occupied(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """occupancy=3, max=5 → free=2; fill tracks=4 must be 422 (сейчас analyze игнорирует occupancy → 200)."""
    _mock_substrate_optimizer(monkeypatch)
    monkeypatch.setattr(
        PlanRepository,
        "get_global_occupancy",
        lambda self, exclude_plan_id=None: {DATE_KEY: 3},
    )

    response = _analyze(
        production_api_client,
        production_admin_cookie,
        tracks=4,
    )
    assert response.status_code == 422, response.text
    detail = response.json().get("detail", "")
    detail_text = detail if isinstance(detail, str) else str(detail)
    assert "свободн" in detail_text.lower()


def test_analyze_allows_fill_within_free_slots_when_day_occupied(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """occupancy=3, max=5 → free=2; fill tracks=2 → 200."""
    _mock_substrate_optimizer(monkeypatch)
    monkeypatch.setattr(
        PlanRepository,
        "get_global_occupancy",
        lambda self, exclude_plan_id=None: {DATE_KEY: 3},
    )

    response = _analyze(
        production_api_client,
        production_admin_cookie,
        tracks=2,
    )
    assert response.status_code == 200, response.text
    assert "urgent_positions" in response.json()


# ---------------------------------------------------------------------------
# API guards (green after T2)
# ---------------------------------------------------------------------------


def test_day_capacity_span_over_366_days_returns_422(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    """(to − from).days > 365 → 422 (D5: span ≤ 366 inclusive)."""
    response = production_api_client.get(
        f"{API_PREFIX}/day-capacity",
        params={"from": "2026-01-01", "to": "2027-01-02"},  # days = 366 > 365
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422, response.text


def test_fill_tracks_above_hard_cap_returns_422(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    """tracks=6 > TRACKS_PER_DAY_HARD_CAP(5) → 422 at schema boundary."""
    response = _analyze(
        production_api_client,
        production_admin_cookie,
        tracks=6,
    )
    assert response.status_code == 422, response.text


def test_candidates_limit_too_large_returns_422(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    """limit=10**9 → 422 (Query le=500)."""
    response = production_api_client.get(
        f"{API_PREFIX}/candidates",
        params={"limit": 10**9},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422, response.text


def test_days_path_non_iso_date_returns_422(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    """Non-ISO path date ``/days/foo`` → 422 (D5)."""
    response = production_api_client.get(
        f"{API_PREFIX}/days/foo",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422, response.text


# ---------------------------------------------------------------------------
# Wave 2 — complete 409 / expected_version / SGP compensate (red until T6–T10)
# ---------------------------------------------------------------------------


def test_repeat_complete_day_returns_409_completed_plates_unchanged(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    production_api_db: str,
) -> None:
    """D2: second complete → 409 ``day_already_completed``; completed_plates qty stable."""
    plan_id = production_built_plan["plan"]["id"]

    first = production_api_client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=production_admin_cookie,
    )
    assert first.status_code == 200, first.text
    qty_after_first = _completed_plates_qty(production_api_db)
    assert qty_after_first > 0

    second = production_api_client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={"plan_id": plan_id, "rejected_plates": []},
        cookies=production_admin_cookie,
    )
    assert second.status_code == 409, second.text
    detail = second.json()["detail"]
    assert isinstance(detail, dict), detail
    assert detail["code"] == "day_already_completed"
    assert _completed_plates_qty(production_api_db) == qty_after_first


def test_complete_day_stale_expected_version_returns_409_kp_not_written_off(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
    production_api_db: str,
) -> None:
    """D3: stale expected_version on complete → 409; KP stay «в плане», no write-off."""
    plan_id = production_built_plan["plan"]["id"]
    plan_version = int(production_built_plan["plan"].get("version") or 1)
    stale_version = plan_version + 100

    before_status = _kp_plates_by_status(production_api_db)
    assert before_status.get("в плане", 0) > 0
    assert _completed_plates_qty(production_api_db) == 0

    response = production_api_client.post(
        f"{API_PREFIX}/days/{DATE_KEY}/complete",
        json={
            "plan_id": plan_id,
            "rejected_plates": [],
            "expected_version": stale_version,
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 409, response.text
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT
    assert detail["details"]["plan_id"] == plan_id

    assert _completed_plates_qty(production_api_db) == 0
    assert _kp_plates_by_status(production_api_db) == before_status


def test_delete_track_stale_expected_version_returns_409_track_remains(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_built_plan: dict[str, Any],
) -> None:
    """D3: stale expected_version on DELETE track → 409; track still on the day."""
    plan_id = production_built_plan["plan"]["id"]
    plan_version = int(production_built_plan["plan"].get("version") or 1)
    stale_version = plan_version + 100

    before = production_api_client.get(
        f"{API_PREFIX}/plans/{plan_id}",
        cookies=production_admin_cookie,
    )
    assert before.status_code == 200, before.text
    tracks_before = before.json()["days"][DATE_KEY]["tracks"]
    assert len(tracks_before) >= 1

    response = production_api_client.delete(
        f"{API_PREFIX}/plans/{plan_id}/days/{DATE_KEY}/tracks/0",
        params={"expected_version": stale_version},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 409, response.text
    detail = response.json()["detail"]
    assert detail["code"] == ERROR_CODE_PLAN_VERSION_CONFLICT

    after = production_api_client.get(
        f"{API_PREFIX}/plans/{plan_id}",
        cookies=production_admin_cookie,
    )
    assert after.status_code == 200, after.text
    tracks_after = after.json()["days"][DATE_KEY]["tracks"]
    assert len(tracks_after) == len(tracks_before)


def test_sgp_reserve_fail_after_build_compensates_no_plan_plates_not_in_plan(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """D4: reserve_on_conn raises after persist → compensate: no plan, plates not «в плане»."""
    sgp_id = _seed_free_sgp_row(production_api_db, qty=1)

    def boom(self, cur, conn, **kwargs):  # noqa: ANN001
        raise SgpError("simulated SGP reserve failure", code="sgp_reserve_failed")

    monkeypatch.setattr(SgpService, "reserve_on_conn", boom)

    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "all",
            "sgp_reservations": [
                {"sgp_id": sgp_id, "target_kp_id": KP_ID, "qty": 1},
            ],
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code != 200, response.text

    listed = production_api_client.get(
        f"{API_PREFIX}/plans",
        cookies=production_admin_cookie,
    )
    assert listed.status_code == 200, listed.text
    assert listed.json()["plans"] == []

    by_status = _kp_plates_by_status(production_api_db)
    assert by_status.get("в плане", 0) == 0
    assert by_status.get("в производстве", 0) == 3

    with sqlite3.connect(production_api_db) as conn:
        free_sgp = conn.execute(
            "SELECT qty, kp_id FROM completed_plates WHERE id = ?",
            (sgp_id,),
        ).fetchone()
    assert free_sgp is not None
    assert int(free_sgp[0]) == 1
    assert free_sgp[1] is None
