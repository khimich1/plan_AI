"""Task 9: GET /production/kp-candidates carries promise meta for the wizard."""

from __future__ import annotations

from datetime import date

import pytest
from fastapi.testclient import TestClient

from app.repositories.kp_repository import KpRepository
from app.repositories.promise_repository import PromiseRepository
from app.schemas.production import KpCandidatesResponse
from app.services.production_service import ProductionService
from core.kp_db_common import _connect
from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService
from tests.helpers.production_api_fixtures import API_PREFIX

PLATE = "ПБ 60-12-8п"
_TS = "2026-09-03T12:00:00"
WEEK_0 = date(2026, 8, 31)  # Monday
WEEK_1 = date(2026, 9, 7)
WEEK_2 = date(2026, 9, 14)


@pytest.fixture()
def db_path(tmp_path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _plate_line(*, qty: int, line_id: str) -> dict:
    return {
        "line_id": line_id,
        "name": PLATE,
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": qty,
        "unit_price": 1000.0,
        "weight": 500.0,
    }


def _save_kp(db_path: str, *, customer: str, qty: int = 2) -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.09.2026",
        [_plate_line(qty=qty, line_id=f"ln_{customer}")],
        customer_name=customer,
        execution_terms="25.09.2026",
        status="в работе",
        db_path=db_path,
    )


def _insert_alloc(
    db_path: str,
    *,
    kp_id: int,
    week_start: date,
    tracks: int,
    kind: str = "promise",
    promise_status: str = "active",
    alloc_status: str = "active",
    promised_date: date = date(2026, 9, 25),
) -> int:
    with _connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_promise (
                kp_id, tracks_total, promised_date, kind, status,
                created_by, created_at, expires_at
            ) VALUES (?, ?, ?, ?, ?, 'alice', ?, NULL)
            """,
            (
                kp_id,
                tracks,
                promised_date.isoformat(),
                kind,
                promise_status,
                _TS,
            ),
        )
        promise_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
            VALUES (?, ?, ?, ?)
            """,
            (promise_id, week_start.isoformat(), tracks, alloc_status),
        )
        conn.commit()
    return promise_id


def _insert_window(
    db_path: str,
    *,
    kp_id: int,
    allocations: tuple[tuple[date, int, str], ...],
    promised_date: date = date(2026, 9, 25),
) -> int:
    tracks_total = sum(tracks for _week, tracks, _status in allocations)
    with _connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_promise (
                kp_id, tracks_total, promised_date, kind, status,
                created_by, created_at, expires_at
            ) VALUES (?, ?, ?, 'promise', 'active', 'alice', ?, NULL)
            """,
            (kp_id, tracks_total, promised_date.isoformat(), _TS),
        )
        promise_id = int(cur.lastrowid)
        for week_start, tracks, alloc_status in allocations:
            cur.execute(
                """
                INSERT INTO kp_promise_alloc (
                    promise_id, week_start, tracks, status
                ) VALUES (?, ?, ?, ?)
                """,
                (promise_id, week_start.isoformat(), tracks, alloc_status),
            )
        conn.commit()
    return promise_id


def _service(db_path: str) -> ProductionService:
    return ProductionService(kp_repository=KpRepository(db_path=db_path))


def test_candidates_without_promises_have_empty_week_summary(db_path: str) -> None:
    kp_id = _save_kp(db_path, customer="Plain")

    payload = _service(db_path).list_kp_candidates()
    parsed = KpCandidatesResponse(**payload)

    assert parsed.count == 1
    assert parsed.items[0].kp_id == kp_id
    assert parsed.items[0].promise is None
    assert parsed.promised_weeks == []


def test_active_promise_is_attached_and_listed_for_week(db_path: str) -> None:
    kp_id = _save_kp(db_path, customer="Promised")
    _insert_alloc(db_path, kp_id=kp_id, week_start=WEEK_1, tracks=4)

    payload = _service(db_path).list_kp_candidates()
    parsed = KpCandidatesResponse(**payload)

    meta = parsed.items[0].promise
    assert meta is not None
    assert meta.promised_date == date(2026, 9, 25)
    assert meta.week_start == WEEK_1
    assert meta.status == "active"
    assert meta.tracks == 4

    assert len(parsed.promised_weeks) == 1
    week = parsed.promised_weeks[0]
    assert week.week_start == WEEK_1
    assert week.items[0].kp_id == kp_id
    assert week.items[0].status == "active"
    assert week.items[0].tracks == 4


def test_overdue_alloc_is_marked_on_candidate_and_week(db_path: str) -> None:
    kp_id = _save_kp(db_path, customer="Overdue")
    _insert_alloc(
        db_path,
        kp_id=kp_id,
        week_start=WEEK_0,
        tracks=3,
        alloc_status="overdue",
    )

    payload = _service(db_path).list_kp_candidates()
    parsed = KpCandidatesResponse(**payload)

    assert parsed.items[0].promise is not None
    assert parsed.items[0].promise.status == "overdue"
    assert parsed.items[0].promise.week_start == WEEK_0
    assert parsed.promised_weeks[0].items[0].status == "overdue"


def test_date_range_keeps_selected_week_and_outside_overdue(db_path: str) -> None:
    selected = _save_kp(db_path, customer="This week")
    other = _save_kp(db_path, customer="Next week")
    missed = _save_kp(db_path, customer="Missed")
    _insert_alloc(db_path, kp_id=selected, week_start=WEEK_1, tracks=2)
    _insert_alloc(db_path, kp_id=other, week_start=WEEK_2, tracks=5)
    _insert_alloc(
        db_path,
        kp_id=missed,
        week_start=WEEK_0,
        tracks=1,
        alloc_status="overdue",
    )

    payload = _service(db_path).list_kp_candidates(
        from_date=date(2026, 9, 7),
        to_date=date(2026, 9, 11),
    )
    parsed = KpCandidatesResponse(**payload)
    weeks = {row.week_start: row for row in parsed.promised_weeks}

    assert WEEK_1 in weeks
    assert WEEK_2 not in weeks
    assert WEEK_0 in weeks
    assert weeks[WEEK_1].items[0].kp_id == selected
    assert weeks[WEEK_0].items[0].kp_id == missed
    assert weeks[WEEK_0].items[0].status == "overdue"

    by_id = {item.kp_id: item for item in parsed.items}
    assert by_id[other].promise is None
    assert by_id[selected].promise is not None
    assert by_id[missed].promise is not None
    assert by_id[missed].promise.status == "overdue"


def test_window_kp_appears_in_both_weeks_candidate_overdue_if_any(db_path: str) -> None:
    kp_id = _save_kp(db_path, customer="Window")
    _insert_window(
        db_path,
        kp_id=kp_id,
        allocations=(
            (WEEK_1, 15, "overdue"),
            (WEEK_2, 5, "active"),
        ),
    )

    payload = _service(db_path).list_kp_candidates()
    parsed = KpCandidatesResponse(**payload)

    assert parsed.items[0].promise is not None
    assert parsed.items[0].promise.status == "overdue"
    assert parsed.items[0].promise.week_start == WEEK_1
    assert parsed.items[0].promise.tracks == 20

    weeks = {row.week_start: row for row in parsed.promised_weeks}
    assert weeks[WEEK_1].items[0].status == "overdue"
    assert weeks[WEEK_1].items[0].tracks == 15
    assert weeks[WEEK_2].items[0].status == "active"
    assert weeks[WEEK_2].items[0].tracks == 5


def test_holds_and_consumed_are_excluded(db_path: str) -> None:
    hold_kp = _save_kp(db_path, customer="Hold")
    done_kp = _save_kp(db_path, customer="Consumed")
    live_kp = _save_kp(db_path, customer="Live")
    _insert_alloc(db_path, kp_id=hold_kp, week_start=WEEK_1, tracks=2, kind="hold")
    _insert_alloc(
        db_path,
        kp_id=done_kp,
        week_start=WEEK_1,
        tracks=2,
        promise_status="consumed",
        alloc_status="consumed",
    )
    _insert_alloc(db_path, kp_id=live_kp, week_start=WEEK_1, tracks=3)

    payload = _service(db_path).list_kp_candidates()
    parsed = KpCandidatesResponse(**payload)
    by_id = {item.kp_id: item for item in parsed.items}

    assert by_id[hold_kp].promise is None
    assert by_id[done_kp].promise is None
    assert by_id[live_kp].promise is not None
    assert [item.kp_id for week in parsed.promised_weeks for item in week.items] == [
        live_kp
    ]


def test_repository_lists_overdue_outside_week_filter(db_path: str) -> None:
    kp_id = _save_kp(db_path, customer="Repo")
    _insert_alloc(
        db_path,
        kp_id=kp_id,
        week_start=WEEK_0,
        tracks=2,
        alloc_status="overdue",
    )
    _insert_alloc(db_path, kp_id=kp_id, week_start=WEEK_2, tracks=4)

    rows = PromiseRepository(db_path=db_path).list_wizard_promise_allocs(
        week_starts=[WEEK_1],
    )
    weeks = {row["week_start"] for row in rows}
    assert weeks == {WEEK_0}
    assert rows[0]["status"] == "overdue"


def test_get_kp_candidates_includes_promised_weeks(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/kp-candidates",
        cookies=production_admin_cookie,
    )
    assert response.status_code == 200
    payload = response.json()
    assert "promised_weeks" in payload
    assert payload["promised_weeks"] == []


def test_get_kp_candidates_rejects_inverted_range(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/kp-candidates",
        params={"from": "2026-09-14", "to": "2026-09-07"},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 400


def test_get_kp_candidates_requires_both_range_params(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
) -> None:
    response = production_api_client.get(
        f"{API_PREFIX}/kp-candidates",
        params={"from": "2026-09-07"},
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422
