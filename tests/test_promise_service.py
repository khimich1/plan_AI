"""Promise quote service (Task 3): occupancy fail-closed, accumulation, holds."""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

import pytest

from app.repositories.promise_repository import PromiseRepository
from app.schemas.archive import PromiseHoldAllocation
from app.services.promise_service import (
    PromiseHoldForbiddenError,
    PromiseHoldNotFoundError,
    PromiseHoldUnavailableError,
    PromiseKnobInvalidError,
    PromiseNotFoundError,
    PromiseService,
    load_plan_occupancy,
)
from core import kp_db_schema
from core.kp_db_common import _connect
from core.production.promise_buckets import OccupancyUnavailableError
from core.production_capacity import MAX_TRACK_LENGTH_M

ADMIN = {"id": 1, "role": "admin"}
TODAY = date(2026, 9, 2)  # Wednesday — partial week Thu–Fri
WEEK_0 = date(2026, 8, 31)
_TS = "2026-09-03T12:00:00"


def _fresh_db(tmp_path: Path, name: str = "promise.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed_kp(conn, kp_id: int) -> None:
    conn.execute(
        "INSERT INTO KP_offers (kp_id, creation_date) VALUES (?, '2026-09-03')",
        (kp_id,),
    )


def _seed_plate(conn, kp_id: int, *, length_m: float, qty: int) -> None:
    conn.execute(
        """
        INSERT INTO kp_plates (
            kp_id, position_number, plate_name, length_m, width_m, qty
        ) VALUES (?, 1, 'ПБ', ?, 1.2, ?)
        """,
        (kp_id, length_m, qty),
    )


def _insert_alloc(
    conn,
    *,
    kp_id: int,
    week_start: date,
    tracks: int,
    kind: str = "promise",
    promise_status: str = "active",
    alloc_status: str = "active",
    expires_at: str | None = None,
) -> int:
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO kp_promise (
            kp_id, tracks_total, promised_date, kind, status,
            created_by, created_at, expires_at
        ) VALUES (?, ?, '2026-09-25', ?, ?, 'alice', ?, ?)
        """,
        (kp_id, tracks, kind, promise_status, _TS, expires_at),
    )
    promise_id = int(cur.lastrowid)
    cur.execute(
        """
        INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
        VALUES (?, ?, ?, ?)
        """,
        (promise_id, week_start.isoformat(), tracks, alloc_status),
    )
    return promise_id


def _service(
    db_path: str,
    *,
    occupancy_loader=None,
    **kwargs,
) -> PromiseService:
    loader = occupancy_loader if occupancy_loader is not None else (lambda: {})
    return PromiseService(
        db_path=db_path,
        occupancy_loader=loader,
        today=TODAY,
        is_workday=lambda day: day.weekday() < 5,
        week_count=3,
        **kwargs,
    )


def test_repository_reads_knob_and_buffer_defaults(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    repo = PromiseRepository(db_path=db_path)

    assert repo.get_promise_tracks_per_day() == 3
    assert repo.get_promise_buffer() == 1.0


def test_repository_sums_promised_not_held(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_kp(conn, 2)
        _insert_alloc(conn, kp_id=1, week_start=WEEK_0, tracks=5)
        _insert_alloc(conn, kp_id=2, week_start=WEEK_0, tracks=3, kind="hold")
        conn.commit()

    repo = PromiseRepository(db_path=db_path)
    assert repo.sum_promised_by_week() == {WEEK_0: 5}
    assert repo.sum_held_by_week(today=TODAY) == {WEEK_0: 3}


def test_repository_skips_expired_holds_and_inactive(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_kp(conn, 2)
        _seed_kp(conn, 3)
        _insert_alloc(
            conn, kp_id=1, week_start=WEEK_0, tracks=2, kind="hold",
            expires_at="2026-09-01T23:59:59",
        )
        _insert_alloc(
            conn, kp_id=2, week_start=WEEK_0, tracks=4, kind="hold",
            expires_at="2026-09-02T23:59:59",
        )
        _insert_alloc(
            conn, kp_id=3, week_start=WEEK_0, tracks=9,
            promise_status="consumed",
        )
        conn.commit()

    repo = PromiseRepository(db_path=db_path)
    assert repo.sum_held_by_week(today=TODAY) == {WEEK_0: 4}
    assert repo.sum_promised_by_week() == {}


def test_load_plan_occupancy_exception_is_fail_closed() -> None:
    def boom() -> dict:
        raise RuntimeError("calendar down")

    with pytest.raises(OccupancyUnavailableError, match="занятость"):
        load_plan_occupancy(calendar_loader=boom)


def test_load_plan_occupancy_none_calendar_is_empty_not_error() -> None:
    assert load_plan_occupancy(calendar_loader=lambda: None) == {}


def test_load_plan_occupancy_malformed_days_info_is_fail_closed() -> None:
    with pytest.raises(OccupancyUnavailableError):
        load_plan_occupancy(calendar_loader=lambda: {"days_info": "broken"})


def test_load_plan_occupancy_extracts_occupied() -> None:
    calendar = {
        "days_info": {
            "2026-09-03": {"occupied": 2, "max": 5},
            "2026-09-04": {"occupied": 0, "max": 5},
        }
    }
    assert load_plan_occupancy(calendar_loader=lambda: calendar) == {
        "2026-09-03": 2,
        "2026-09-04": 0,
    }


def test_quote_matches_spec_contract(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.commit()

    quote = _service(db_path).get_quote(1, user=ADMIN)

    assert quote.tracks == 1
    assert quote.solo_days == 1
    assert quote.solo_date == date(2026, 9, 3)
    assert quote.solo_week_end_date == date(2026, 9, 4)
    assert quote.knob == 3
    assert quote.window is not None
    assert quote.window.from_week == WEEK_0
    assert quote.window.to_week == WEEK_0
    assert quote.window.promised_date == date(2026, 9, 4)
    assert quote.earliest_start_week == WEEK_0
    assert len(quote.weeks) == 3
    week0 = quote.weeks[0]
    assert week0.week_start == WEEK_0
    assert week0.workdays == 2
    assert week0.capacity == 6
    assert week0.planned == 0
    assert week0.promised == 0
    assert week0.held == 0
    assert week0.free == 6


def test_quote_second_kp_sees_first_promise(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        _seed_kp(conn, 2)
        _seed_plate(conn, 2, length_m=MAX_TRACK_LENGTH_M, qty=1)
        _insert_alloc(conn, kp_id=1, week_start=WEEK_0, tracks=5)
        conn.commit()

    quote = _service(db_path).get_quote(2, user=ADMIN)

    assert quote.tracks == 1
    assert quote.weeks[0].promised == 5
    assert quote.weeks[0].held == 0
    assert quote.weeks[0].planned == 0
    assert quote.weeks[0].free == 1
    assert quote.weeks[0].free == max(
        0, quote.weeks[0].capacity - quote.weeks[0].planned - quote.weeks[0].promised
    )


def test_quote_holds_do_not_reduce_free(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        _seed_kp(conn, 2)
        _insert_alloc(conn, kp_id=2, week_start=WEEK_0, tracks=4, kind="hold")
        conn.commit()

    quote = _service(db_path).get_quote(1, user=ADMIN)

    assert quote.weeks[0].held == 4
    assert quote.weeks[0].promised == 0
    assert quote.weeks[0].free == quote.weeks[0].capacity


def test_quote_occupancy_error_is_not_all_free(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.commit()

    def boom() -> dict:
        raise RuntimeError("plan calendar unavailable")

    service = _service(db_path, occupancy_loader=boom)
    with pytest.raises(OccupancyUnavailableError, match="занятость"):
        service.get_quote(1, user=ADMIN)


def test_quote_rejects_none_occupancy_map(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.commit()

    service = _service(db_path, occupancy_loader=lambda: None)
    with pytest.raises(OccupancyUnavailableError):
        service.get_quote(1, user=ADMIN)


def test_quote_counts_planned_from_occupancy(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.commit()

    occupancy = {date(2026, 9, 3): 2, date(2026, 9, 4): 1}
    quote = _service(db_path, occupancy_loader=lambda: occupancy).get_quote(
        1, user=ADMIN
    )

    assert quote.weeks[0].planned == 3
    assert quote.weeks[0].free == 3


def test_quote_uses_knob_from_setting(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.execute(
            "UPDATE kp_setting SET value = '2' WHERE key = 'promise_tracks_per_day'"
        )
        conn.commit()

    quote = _service(db_path).get_quote(1, user=ADMIN)

    assert quote.knob == 2
    assert quote.weeks[0].capacity == 4


def test_quote_missing_kp_raises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with pytest.raises(PromiseNotFoundError):
        _service(db_path).get_quote(999, user=ADMIN)


# --- Task 5: holds -----------------------------------------------------------

MANAGER_ALICE = {"id": 2, "username": "alice", "role": "manager"}
MANAGER_BOB = {"id": 3, "username": "bob", "role": "manager"}
NOW = datetime(2026, 9, 2, 12, 0, 0)
END_OF_DAY = datetime(2026, 9, 2, 23, 59, 59)
AFTER_MIDNIGHT = datetime(2026, 9, 3, 0, 0, 1)


def _hold_service(db_path: str, *, now: datetime = NOW, **kwargs) -> PromiseService:
    return _service(db_path, now=now, **kwargs)


def _seed_quote_kp(conn, kp_id: int, *, tracks: int = 1) -> None:
    _seed_kp(conn, kp_id)
    _seed_plate(conn, kp_id, length_m=MAX_TRACK_LENGTH_M, qty=tracks)


def test_create_hold_uses_fresh_quote_and_expires_end_of_day(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    hold = _hold_service(db_path).create_hold(1, user=MANAGER_ALICE)

    assert hold.kp_id == 1
    assert hold.kind == "hold"
    assert hold.status == "active"
    assert hold.tracks_total == 1
    assert hold.promised_date == date(2026, 9, 4)
    assert hold.expires_at == END_OF_DAY
    assert hold.created_by == "alice"
    assert hold.allocations == [
        PromiseHoldAllocation(week_start=WEEK_0, tracks=1),
    ]

    repo = PromiseRepository(db_path=db_path)
    row = repo.get_active_hold(1, now=NOW)
    assert row is not None
    assert row["kind"] == "hold"
    assert row["status"] == "active"
    assert row["expires_at"].startswith("2026-09-02T23:59:59")


def test_hold_shows_as_held_in_other_quote_and_does_not_reduce_free(
    tmp_path: Path,
) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1, tracks=4)
        _seed_quote_kp(conn, 2)
        conn.commit()

    _hold_service(db_path).create_hold(1, user=MANAGER_ALICE)
    quote = _hold_service(db_path).get_quote(2, user=ADMIN)

    week0 = quote.weeks[0]
    assert week0.held == 4
    assert week0.promised == 0
    assert week0.free == week0.capacity
    assert week0.free == max(0, week0.capacity - week0.planned - week0.promised)


def test_expired_hold_reads_as_expired_after_midnight(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    created = service.create_hold(1, user=MANAGER_ALICE)
    assert created.status == "active"

    later = PromiseService(
        db_path=db_path,
        occupancy_loader=lambda: {},
        today=AFTER_MIDNIGHT.date(),
        now=AFTER_MIDNIGHT,
        is_workday=lambda day: day.weekday() < 5,
        week_count=3,
    )
    expired = later.get_hold(1, user=ADMIN)
    assert expired is not None
    assert expired.status == "expired"
    assert later.repository.get_active_hold(1, now=AFTER_MIDNIGHT) is None

    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 2)
        conn.commit()
    quote = later.get_quote(2, user=ADMIN)
    assert quote.weeks[0].held == 0


def test_repeat_hold_replaces_old_for_same_kp(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    first = service.create_hold(1, user=MANAGER_ALICE)
    second = service.create_hold(1, user=MANAGER_ALICE)

    assert second.id != first.id
    assert second.status == "active"
    assert service.repository.get_active_hold(1, now=NOW)["id"] == second.id

    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT id, status FROM kp_promise WHERE kp_id = 1 AND kind = 'hold' ORDER BY id"
        ).fetchall()
    assert [row[1] for row in rows] == ["released", "active"]
    assert rows[0][0] == first.id
    assert rows[1][0] == second.id


def test_owner_can_release_hold(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    service.create_hold(1, user=MANAGER_ALICE)
    released = service.release_hold(1, user=MANAGER_ALICE)

    assert released.status == "released"
    assert service.repository.get_active_hold(1, now=NOW) is None


def test_admin_can_release_others_hold(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    service.create_hold(1, user=MANAGER_ALICE)
    released = service.release_hold(1, user=ADMIN)

    assert released.status == "released"


def test_other_manager_cannot_release_hold(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    service.create_hold(1, user=MANAGER_ALICE)

    with pytest.raises(PromiseHoldForbiddenError):
        service.release_hold(1, user=MANAGER_BOB)

    assert service.repository.get_active_hold(1, now=NOW) is not None


def test_release_missing_hold_raises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    with pytest.raises(PromiseHoldNotFoundError):
        _hold_service(db_path).release_hold(1, user=ADMIN)


def test_create_hold_without_window_raises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1, tracks=20)
        _seed_kp(conn, 99)
        for offset, week in enumerate((WEEK_0, date(2026, 9, 7), date(2026, 9, 14))):
            _insert_alloc(
                conn,
                kp_id=99,
                week_start=week,
                tracks=30 + offset,
            )
        conn.commit()

    with pytest.raises(PromiseHoldUnavailableError):
        _hold_service(db_path).create_hold(1, user=MANAGER_ALICE)


def test_overnight_expire_does_not_touch_promises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _insert_alloc(
            conn,
            kp_id=1,
            week_start=WEEK_0,
            tracks=5,
            kind="promise",
            expires_at="2026-09-01T23:59:59",
        )
        conn.commit()

    repo = PromiseRepository(db_path=db_path)
    repo.expire_stale_holds(now=AFTER_MIDNIGHT)
    assert repo.sum_promised_by_week() == {WEEK_0: 5}

    with _connect(db_path) as conn:
        status = conn.execute(
            "SELECT status FROM kp_promise WHERE kp_id = 1"
        ).fetchone()[0]
    assert status == "active"


# --- Task 6: move-to-production gate ----------------------------------------

def test_evaluate_move_gate_rejects_date_before_promised(tmp_path: Path) -> None:
    from app.services.promise_service import PromiseGateError

    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    with pytest.raises(PromiseGateError, match="04.09.2026") as exc_info:
        service.evaluate_move_gate(1, date(2026, 9, 3), user=ADMIN)

    assert exc_info.value.earliest == date(2026, 9, 4)


def test_evaluate_move_gate_accepts_promised_date_and_builds_payload(
    tmp_path: Path,
) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    payload = _hold_service(db_path).evaluate_move_gate(
        1, date(2026, 9, 4), user=MANAGER_ALICE
    )

    assert payload is not None
    assert payload.convert_hold_id is None
    assert payload.tracks_total == 1
    assert payload.promised_date == "2026-09-04"
    assert payload.allocations == ((WEEK_0.isoformat(), 1),)
    assert payload.created_by == "alice"


def test_evaluate_move_gate_converts_active_hold(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path)
    hold = service.create_hold(1, user=MANAGER_ALICE)
    payload = service.evaluate_move_gate(1, date(2026, 9, 4), user=MANAGER_ALICE)

    assert payload is not None
    assert payload.convert_hold_id == hold.id


def test_evaluate_move_gate_occupancy_error_is_not_all_free(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_quote_kp(conn, 1)
        conn.commit()

    service = _hold_service(db_path, occupancy_loader=lambda: (_ for _ in ()).throw(
        RuntimeError("calendar down")
    ))
    with pytest.raises(OccupancyUnavailableError, match="занятость"):
        service.evaluate_move_gate(1, date(2026, 10, 15), user=ADMIN)


def test_evaluate_move_gate_skips_promise_when_no_tracks(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        conn.commit()

    payload = _hold_service(db_path).evaluate_move_gate(
        1, date(2026, 9, 4), user=ADMIN
    )
    assert payload is None


# --- Task 8: settle on plan commit (consume / overdue) ----------------------

WEEK_1 = date(2026, 9, 7)


def _insert_promise_window(
    conn,
    *,
    kp_id: int,
    allocations: tuple[tuple[date, int], ...],
) -> int:
    tracks_total = sum(tracks for _week, tracks in allocations)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO kp_promise (
            kp_id, tracks_total, promised_date, kind, status,
            created_by, created_at, expires_at
        ) VALUES (?, ?, '2026-09-25', 'promise', 'active', 'alice', ?, NULL)
        """,
        (kp_id, tracks_total, _TS),
    )
    promise_id = int(cur.lastrowid)
    for week_start, tracks in allocations:
        cur.execute(
            """
            INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
            VALUES (?, ?, ?, 'active')
            """,
            (promise_id, week_start.isoformat(), tracks),
        )
    return promise_id


def _alloc_statuses(db_path: str, promise_id: int) -> list[tuple[date, str]]:
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT week_start, status FROM kp_promise_alloc "
            "WHERE promise_id = ? ORDER BY week_start",
            (promise_id,),
        ).fetchall()
    return [(date.fromisoformat(str(week)), str(status)) for week, status in rows]


def _promise_status(db_path: str, promise_id: int) -> str:
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT status FROM kp_promise WHERE id = ?",
            (promise_id,),
        ).fetchone()
    assert row is not None
    return str(row[0])


def test_settle_consumes_entered_kp_and_drops_promised(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.execute(
            "UPDATE kp_plates SET status = 'в плане', plan_id = 'p1' WHERE kp_id = 1"
        )
        promise_id = _insert_promise_window(
            conn, kp_id=1, allocations=((WEEK_0, 5),)
        )
        conn.commit()

    repo = PromiseRepository(db_path=db_path)
    assert repo.sum_promised_by_week() == {WEEK_0: 5}

    settlement = _service(db_path).settle_plan_commit(
        entered_kp_ids={1},
        covered_weeks=(WEEK_0,),
    )

    assert promise_id in settlement.consumed_promise_ids
    assert _promise_status(db_path, promise_id) == "consumed"
    assert _alloc_statuses(db_path, promise_id) == [(WEEK_0, "consumed")]
    assert repo.sum_promised_by_week() == {}
    assert _service(db_path).list_overdue_allocations() == []


def test_settle_marks_missed_promise_overdue_without_blocking(
    tmp_path: Path,
) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.execute(
            "UPDATE kp_plates SET status = 'в плане', plan_id = 'p1' WHERE kp_id = 1"
        )
        _seed_kp(conn, 2)
        entered_id = _insert_promise_window(
            conn, kp_id=1, allocations=((WEEK_0, 5),)
        )
        missed_id = _insert_promise_window(
            conn, kp_id=2, allocations=((WEEK_0, 3),)
        )
        conn.commit()

    settlement = _service(db_path).settle_plan_commit(
        entered_kp_ids={1},
        covered_weeks=(WEEK_0,),
    )

    assert entered_id in settlement.consumed_promise_ids
    assert missed_id not in settlement.consumed_promise_ids
    assert _promise_status(db_path, missed_id) == "active"
    assert _alloc_statuses(db_path, missed_id) == [(WEEK_0, "overdue")]
    overdue = _service(db_path).list_overdue_allocations()
    assert [(row["kp_id"], row["week_start"], row["tracks"]) for row in overdue] == [
        (2, WEEK_0, 3)
    ]
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {}


def test_settle_keeps_promise_active_while_unplanned_plates_remain(
    tmp_path: Path,
) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=2)
        promise_id = _insert_promise_window(
            conn, kp_id=1, allocations=((WEEK_0, 5), (WEEK_1, 5))
        )
        conn.commit()

    settlement = _service(db_path).settle_plan_commit(
        entered_kp_ids={1},
        covered_weeks=(WEEK_0,),
    )

    assert promise_id not in settlement.consumed_promise_ids
    assert _promise_status(db_path, promise_id) == "active"
    assert _alloc_statuses(db_path, promise_id) == [
        (WEEK_0, "consumed"),
        (WEEK_1, "active"),
    ]
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {WEEK_1: 5}


def test_settle_on_external_conn_rolls_back_with_caller(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.execute(
            "UPDATE kp_plates SET status = 'в плане', plan_id = 'p1' WHERE kp_id = 1"
        )
        promise_id = _insert_promise_window(
            conn, kp_id=1, allocations=((WEEK_0, 5),)
        )
        conn.commit()

    conn = _connect(db_path)
    try:
        _service(db_path).settle_plan_commit(
            entered_kp_ids={1},
            covered_weeks=(WEEK_0,),
            _external_conn=conn,
        )
        conn.rollback()
    finally:
        conn.close()

    assert _promise_status(db_path, promise_id) == "active"
    assert _alloc_statuses(db_path, promise_id) == [(WEEK_0, "active")]
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {WEEK_0: 5}


# --- Task 13: promise_tracks_per_day knob ------------------------------------

KNOB_ACTOR = {"id": 1, "username": "tester", "role": "admin"}


def test_get_tracks_per_day_defaults_to_three(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    row = _service(db_path).get_tracks_per_day(user=KNOB_ACTOR)

    assert row.tracks_per_day == 3
    assert row.min == 1
    assert row.max == 5
    assert row.updated_by == "system"


def test_set_tracks_per_day_writes_audit_and_changes_new_quotes(
    tmp_path: Path,
) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=1)
        conn.commit()

    service = _service(db_path, now=datetime(2026, 9, 3, 15, 30, 0))
    updated = service.set_tracks_per_day(2, user=KNOB_ACTOR)

    assert updated.tracks_per_day == 2
    assert updated.updated_by == "tester"
    assert updated.updated_at == datetime(2026, 9, 3, 15, 30, 0)

    quote = service.get_quote(1, user=ADMIN)
    assert quote.knob == 2
    assert quote.weeks[0].capacity == 4


def test_set_tracks_per_day_does_not_touch_active_promises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        promise_id = _insert_alloc(conn, kp_id=1, week_start=WEEK_0, tracks=8)
        conn.commit()

    _service(db_path).set_tracks_per_day(5, user=KNOB_ACTOR)

    repo = PromiseRepository(db_path=db_path)
    assert repo.get_promise_tracks_per_day() == 5
    assert repo.sum_promised_by_week() == {WEEK_0: 8}
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT tracks_total, promised_date, status FROM kp_promise WHERE id = ?",
            (promise_id,),
        ).fetchone()
    assert tuple(row) == (8, "2026-09-25", "active")


@pytest.mark.parametrize("bad", [0, 6, -1])
def test_set_tracks_per_day_rejects_out_of_range(tmp_path: Path, bad: int) -> None:
    db_path = _fresh_db(tmp_path)
    with pytest.raises(PromiseKnobInvalidError):
        _service(db_path).set_tracks_per_day(bad, user=KNOB_ACTOR)
    assert _service(db_path).get_tracks_per_day(user=KNOB_ACTOR).tracks_per_day == 3


# --- Task 14: release on delete + recalc on composition edit -----------------

OWNER_ID = 7


def _seed_owner(conn, kp_id: int, owner_user_id: int = OWNER_ID) -> None:
    conn.execute(
        """
        INSERT INTO kp_meta (kp_id, status, owner_user_id, product_type)
        VALUES (?, 'в работе', ?, 'plates')
        """,
        (kp_id, owner_user_id),
    )


def _set_plate_qty(conn, kp_id: int, qty: int) -> None:
    conn.execute("UPDATE kp_plates SET qty = ? WHERE kp_id = ?", (qty, kp_id))


def _insert_dated_promise(
    conn,
    *,
    kp_id: int,
    promised_date: date,
    allocations: tuple[tuple[date, int], ...],
    kind: str = "promise",
    created_by: str = "alice",
) -> int:
    tracks_total = sum(tracks for _week, tracks in allocations)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO kp_promise (
            kp_id, tracks_total, promised_date, kind, status,
            created_by, created_at, expires_at
        ) VALUES (?, ?, ?, ?, 'active', ?, ?, NULL)
        """,
        (kp_id, tracks_total, promised_date.isoformat(), kind, created_by, _TS),
    )
    promise_id = int(cur.lastrowid)
    for week_start, tracks in allocations:
        cur.execute(
            """
            INSERT INTO kp_promise_alloc (promise_id, week_start, tracks, status)
            VALUES (?, ?, ?, 'active')
            """,
            (promise_id, week_start.isoformat(), tracks),
        )
    return promise_id


def test_release_on_delete_frees_promise_and_hold_buckets(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_kp(conn, 2)
        promise_id = _insert_dated_promise(
            conn,
            kp_id=1,
            promised_date=date(2026, 9, 4),
            allocations=((WEEK_0, 4),),
        )
        hold_id = _insert_dated_promise(
            conn,
            kp_id=1,
            promised_date=date(2026, 9, 4),
            allocations=((WEEK_0, 2),),
            kind="hold",
        )
        _insert_dated_promise(
            conn,
            kp_id=2,
            promised_date=date(2026, 9, 4),
            allocations=((WEEK_0, 1),),
        )
        conn.commit()

    released = _service(db_path).release_on_delete(1)

    assert set(released) == {promise_id, hold_id}
    assert _promise_status(db_path, promise_id) == "released"
    assert _promise_status(db_path, hold_id) == "released"
    repo = PromiseRepository(db_path=db_path)
    assert repo.sum_promised_by_week() == {WEEK_0: 1}
    assert repo.sum_held_by_week(today=TODAY) == {}


def test_release_on_delete_is_noop_without_active(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        conn.commit()

    assert _service(db_path).release_on_delete(1) == ()


def test_recalc_increase_that_misses_window_notifies(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_owner(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=4)
        promise_id = _insert_dated_promise(
            conn,
            kp_id=1,
            promised_date=date(2026, 9, 4),
            allocations=((WEEK_0, 4),),
        )
        _set_plate_qty(conn, 1, 8)
        conn.commit()

    result = _service(db_path).recalc_on_composition_change(1)

    assert result.promise_id == promise_id
    assert result.old_promised_date == date(2026, 9, 4)
    assert result.new_promised_date == date(2026, 9, 11)
    assert result.tracks == 8
    assert result.notified is True
    assert result.notification_id is not None

    repo = PromiseRepository(db_path=db_path)
    assert repo.sum_promised_by_week() == {WEEK_1: 8}
    notes = repo.list_notifications(user_id=OWNER_ID)
    assert len(notes) == 1
    assert notes[0]["kind"] == "promised_date_shifted"
    payload = json.loads(notes[0]["payload_json"])
    assert payload["kp_id"] == 1
    assert payload["old_promised_date"] == "2026-09-04"
    assert payload["new_promised_date"] == "2026-09-11"


def test_recalc_decrease_updates_without_notification(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_owner(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=8)
        promise_id = _insert_dated_promise(
            conn,
            kp_id=1,
            promised_date=date(2026, 9, 11),
            allocations=((WEEK_0, 6), (WEEK_1, 2)),
        )
        _set_plate_qty(conn, 1, 1)
        conn.commit()

    result = _service(db_path).recalc_on_composition_change(1)

    assert result.promise_id == promise_id
    assert result.old_promised_date == date(2026, 9, 11)
    assert result.new_promised_date == date(2026, 9, 4)
    assert result.tracks == 1
    assert result.notified is False
    assert result.notification_id is None
    assert PromiseRepository(db_path=db_path).list_notifications(user_id=OWNER_ID) == []
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {WEEK_0: 1}


def test_recalc_same_window_does_not_notify(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_owner(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=2)
        _insert_dated_promise(
            conn,
            kp_id=1,
            promised_date=date(2026, 9, 4),
            allocations=((WEEK_0, 2),),
        )
        _set_plate_qty(conn, 1, 3)
        conn.commit()

    result = _service(db_path).recalc_on_composition_change(1)

    assert result.tracks == 3
    assert result.new_promised_date == date(2026, 9, 4)
    assert result.notified is False
    assert PromiseRepository(db_path=db_path).list_notifications(user_id=OWNER_ID) == []
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {WEEK_0: 3}


def test_recalc_without_active_promise_is_noop(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        _seed_kp(conn, 1)
        _seed_plate(conn, 1, length_m=MAX_TRACK_LENGTH_M, qty=2)
        conn.commit()

    result = _service(db_path).recalc_on_composition_change(1)

    assert result.promise_id is None
    assert result.notified is False
    assert PromiseRepository(db_path=db_path).sum_promised_by_week() == {}
