"""Task 11: plan-build exclusions journal + in-web notification (level 2)."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest
from pydantic import ValidationError

from app.repositories.promise_repository import PromiseRepository
from app.schemas.production import BuildPlanRequest, PromiseExclusionItem
from app.services.promise_service import (
    PromiseExclusionError,
    PromiseService,
)
from core import kp_db_schema
from core.kp_db_common import _connect

WEEK = date(2026, 9, 7)
_TS = "2026-09-03T12:00:00"
PLANNER = {"id": 2, "username": "prod_user", "role": "production"}
OWNER_ID = 7


def _fresh_db(tmp_path: Path, name: str = "exclusion.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed_kp(conn, kp_id: int, *, owner_user_id: int | None = OWNER_ID) -> None:
    conn.execute(
        "INSERT INTO KP_offers (kp_id, creation_date) VALUES (?, '2026-09-03')",
        (kp_id,),
    )
    if owner_user_id is None:
        return
    conn.execute(
        """
        INSERT INTO kp_meta (kp_id, status, owner_user_id, product_type)
        VALUES (?, 'в работе', ?, 'plates')
        """,
        (kp_id, owner_user_id),
    )


def _insert_promise(conn, *, kp_id: int, week_start: date = WEEK, created_by: str = "alice") -> int:
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO kp_promise (
            kp_id, tracks_total, promised_date, kind, status,
            created_by, created_at, expires_at
        ) VALUES (?, 4, '2026-09-25', 'promise', 'active', ?, ?, NULL)
        """,
        (kp_id, created_by, _TS),
    )
    promise_id = int(cur.lastrowid)
    cur.execute(
        """
        INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
        VALUES (?, ?, 4, 'active')
        """,
        (promise_id, week_start.isoformat()),
    )
    return promise_id


def _service(db_path: str) -> PromiseService:
    return PromiseService(
        db_path=db_path,
        occupancy_loader=lambda: {},
        today=date(2026, 9, 3),
        is_workday=lambda day: day.weekday() < 5,
        week_count=3,
    )


def test_schema_rejects_exclusion_without_reason() -> None:
    with pytest.raises(ValidationError):
        PromiseExclusionItem(kp_id=1, week_start=WEEK, reason="")
    with pytest.raises(ValidationError):
        PromiseExclusionItem(kp_id=1, week_start=WEEK, reason="   ")
    with pytest.raises(ValidationError):
        BuildPlanRequest(
            start_date="2026-09-07",
            tracks_count=3,
            filter_method="kp",
            exclusions=[{"kp_id": 1, "week_start": "2026-09-07"}],
        )


def test_schema_accepts_exclusions_on_build_request() -> None:
    payload = BuildPlanRequest(
        start_date="2026-09-07",
        tracks_count=3,
        filter_method="kp",
        selected_kp_ids=[2],
        exclusions=[
            {"kp_id": 1, "week_start": "2026-09-07", "reason": "нет арматуры"},
        ],
    )
    assert payload.exclusions is not None
    assert payload.exclusions[0].kp_id == 1
    assert payload.exclusions[0].week_start == WEEK
    assert payload.exclusions[0].reason == "нет арматуры"


def test_record_writes_exclusion_and_notification_same_tx(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _insert_promise(conn, kp_id=1)
        conn.commit()

    written = _service(db_path).record_plan_exclusions(
        plan_id="plan-11",
        exclusions=[
            {"kp_id": 1, "week_start": WEEK, "reason": "сняли из-за окна"},
        ],
        excluded_by="prod_user",
        user=PLANNER,
    )

    assert len(written) == 1
    assert written[0].notification_id is not None

    journal = _service(db_path).list_plan_exclusions(plan_id="plan-11")
    assert [(row["kp_id"], row["week_start"], row["reason"], row["excluded_by"]) for row in journal] == [
        (1, WEEK, "сняли из-за окна", "prod_user"),
    ]

    notes = PromiseRepository(db_path=db_path).list_notifications(user_id=OWNER_ID)
    assert len(notes) == 1
    assert notes[0]["kind"] == "promise_excluded"
    payload = json.loads(notes[0]["payload_json"])
    assert payload == {
        "kp_id": 1,
        "week_start": WEEK.isoformat(),
        "reason": "сняли из-за окна",
    }


def test_exclusion_and_notification_roll_back_together(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _insert_promise(conn, kp_id=1)
        conn.commit()

    conn = _connect(db_path)
    try:
        _service(db_path).record_plan_exclusions(
            plan_id="plan-rb",
            exclusions=[{"kp_id": 1, "week_start": WEEK, "reason": "откат"}],
            excluded_by="prod_user",
            user=PLANNER,
            _external_conn=conn,
        )
        conn.rollback()
    finally:
        conn.close()

    assert _service(db_path).list_plan_exclusions(plan_id="plan-rb") == []
    assert PromiseRepository(db_path=db_path).list_notifications() == []


def test_record_without_reason_is_rejected(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with pytest.raises(PromiseExclusionError, match="Причина"):
        _service(db_path).record_plan_exclusions(
            plan_id="plan-x",
            exclusions=[{"kp_id": 1, "week_start": WEEK, "reason": "  "}],
            excluded_by="prod_user",
            user=PLANNER,
        )


def test_missing_owner_still_writes_journal_without_blocking(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1, owner_user_id=None)
        _insert_promise(conn, kp_id=1)
        conn.commit()

    written = _service(db_path).record_plan_exclusions(
        plan_id="plan-no-owner",
        exclusions=[{"kp_id": 1, "week_start": WEEK, "reason": "смена приоритета"}],
        excluded_by="prod_user",
        user=PLANNER,
    )

    assert written[0].notification_id is None
    journal = _service(db_path).list_plan_exclusions(plan_id="plan-no-owner")
    assert len(journal) == 1
    assert PromiseRepository(db_path=db_path).list_notifications() == []


def test_build_request_without_exclusions_stays_optional() -> None:
    payload = BuildPlanRequest(
        start_date="2026-09-07",
        tracks_count=2,
        filter_method="all",
    )
    assert payload.exclusions == []


def test_build_endpoint_rejects_exclusion_without_reason(
    production_api_client,
    production_admin_cookie,
) -> None:
    from tests.helpers.production_api_fixtures import API_PREFIX, DATE_KEY

    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "kp",
            "selected_kp_ids": [1],
            "exclusions": [
                {"kp_id": 1, "week_start": "2026-09-07", "reason": ""},
            ],
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422


def test_build_endpoint_writes_exclusion_journal_without_blocking(
    production_api_client,
    production_admin_cookie,
    production_api_db: str,
) -> None:
    from tests.helpers.production_api_fixtures import API_PREFIX, DATE_KEY, KP_ID

    with _connect(production_api_db) as conn:
        conn.execute(
            "UPDATE kp_meta SET owner_user_id = ? WHERE kp_id = ?",
            (OWNER_ID, KP_ID),
        )
        _insert_promise(conn, kp_id=KP_ID, week_start=date(2026, 4, 20))
        conn.commit()

    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "all",
            "exclusions": [
                {
                    "kp_id": KP_ID,
                    "week_start": "2026-04-20",
                    "reason": "нет форм",
                },
            ],
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200, response.text
    plan_id = response.json()["plan"]["id"]

    journal = _service(production_api_db).list_plan_exclusions(plan_id=plan_id)
    assert len(journal) == 1
    assert journal[0]["kp_id"] == KP_ID
    assert journal[0]["reason"] == "нет форм"
    assert journal[0]["excluded_by"] == "admin"

    notes = PromiseRepository(db_path=production_api_db).list_notifications(
        user_id=OWNER_ID, kind="promise_excluded"
    )
    assert len(notes) == 1
    payload = json.loads(notes[0]["payload_json"])
    assert payload["kp_id"] == KP_ID
    assert payload["reason"] == "нет форм"
