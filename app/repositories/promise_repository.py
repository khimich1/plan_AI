"""Promise journal: settings, reads, hold writes, lazy hold expiry."""

from __future__ import annotations

import json
import sqlite3
from collections.abc import Collection, Sequence
from contextlib import contextmanager
from datetime import date, datetime, time

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.production.promise_buckets import (
    DEFAULT_PROMISE_TRACKS_PER_DAY,
    DEFAULT_TRACK_BUFFER,
    clamp_promise_knob,
)

SETTING_TRACKS_PER_DAY = "promise_tracks_per_day"
SETTING_BUFFER = "promise_buffer"

_HOLD_COLS = (
    "id, kp_id, tracks_total, promised_date, kind, status, "
    "created_by, created_at, expires_at"
)


def _iso_dt(value: datetime) -> str:
    return value.isoformat(timespec="seconds")


def _moment_for_day(day: date) -> datetime:
    return datetime.combine(day, time(12, 0, 0))


class PromiseRepository:
    """Reads promise/hold allocations; writes holds; expires stale holds on read."""

    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def _connect(self):
        ensure_schema(self.db_path)
        return _connect(self.db_path)

    def get_setting(self, key: str) -> str | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT value FROM kp_setting WHERE key = ?",
                (key,),
            ).fetchone()
        return None if row is None else str(row[0])

    def get_promise_tracks_per_day(self) -> int:
        return int(self.get_promise_tracks_per_day_row()["tracks_per_day"])

    def get_promise_tracks_per_day_row(self) -> dict:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT value, updated_by, updated_at FROM kp_setting WHERE key = ?",
                (SETTING_TRACKS_PER_DAY,),
            ).fetchone()
        if row is None:
            return {
                "tracks_per_day": DEFAULT_PROMISE_TRACKS_PER_DAY,
                "updated_by": None,
                "updated_at": None,
            }
        try:
            value = clamp_promise_knob(int(row[0]))
        except (TypeError, ValueError):
            value = DEFAULT_PROMISE_TRACKS_PER_DAY
        return {
            "tracks_per_day": value,
            "updated_by": None if row[1] is None else str(row[1]),
            "updated_at": None if row[2] is None else str(row[2]),
        }

    def set_promise_tracks_per_day(
        self,
        value: int,
        *,
        updated_by: str,
        updated_at: datetime,
    ) -> dict:
        """Write knob + audit. Does not rewrite active promises."""
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO kp_setting (key, value, updated_by, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    value = excluded.value,
                    updated_by = excluded.updated_by,
                    updated_at = excluded.updated_at
                """,
                (
                    SETTING_TRACKS_PER_DAY,
                    str(int(value)),
                    updated_by,
                    _iso_dt(updated_at),
                ),
            )
        return self.get_promise_tracks_per_day_row()

    def get_promise_buffer(self) -> float:
        raw = self.get_setting(SETTING_BUFFER)
        if raw is None:
            return DEFAULT_TRACK_BUFFER
        try:
            value = float(raw)
        except (TypeError, ValueError):
            return DEFAULT_TRACK_BUFFER
        return value if value > 0 else DEFAULT_TRACK_BUFFER

    def sum_promised_by_week(self, *, exclude_kp_id: int | None = None) -> dict[date, int]:
        return self._sum_allocs_by_week(kind="promise", exclude_kp_id=exclude_kp_id)

    def sum_held_by_week(
        self,
        *,
        today: date,
        exclude_kp_id: int | None = None,
        now: datetime | None = None,
    ) -> dict[date, int]:
        self.expire_stale_holds(now=now or _moment_for_day(today))
        return self._sum_allocs_by_week(kind="hold", exclude_kp_id=exclude_kp_id)

    def list_week_allocs(
        self,
        week_start: date,
        kinds: Sequence[str] = ("promise", "hold"),
        *,
        now: datetime | None = None,
    ) -> list[dict]:
        """Active hold/promise allocs for a week. Does not exclude any kp_id."""
        moment = now or datetime.now()
        wanted = tuple(kinds)
        if not wanted:
            return []
        placeholders = ",".join("?" * len(wanted))
        with self._connect() as conn:
            self._expire_stale_holds(conn, now=moment)
            rows = conn.execute(
                f"""
                SELECT p.kp_id, p.kind, a.tracks, p.promised_date
                FROM kp_promise_alloc a
                JOIN kp_promise p ON p.id = a.promise_id
                WHERE a.week_start = ?
                  AND p.status = 'active'
                  AND a.status = 'active'
                  AND p.kind IN ({placeholders})
                ORDER BY CASE p.kind WHEN 'promise' THEN 0 ELSE 1 END, p.kp_id, a.id
                """,
                (week_start.isoformat(), *wanted),
            ).fetchall()
        return [
            {
                "kp_id": int(row[0]),
                "kind": str(row[1]),
                "tracks": int(row[2]),
                "promised_date": date.fromisoformat(str(row[3])),
            }
            for row in rows
        ]

    def expire_stale_holds(self, *, now: datetime | None = None) -> int:
        moment = now or datetime.now()
        with self._connect() as conn:
            return self._expire_stale_holds(conn, now=moment)

    def get_active_hold(self, kp_id: int, *, now: datetime | None = None) -> dict | None:
        moment = now or datetime.now()
        with self._connect() as conn:
            self._expire_stale_holds(conn, now=moment)
            row = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise "
                "WHERE kp_id = ? AND kind = 'hold' AND status = 'active' "
                "ORDER BY id DESC LIMIT 1",
                (int(kp_id),),
            ).fetchone()
            return None if row is None else self._hold_from_row(conn, row)

    def get_latest_hold(self, kp_id: int, *, now: datetime | None = None) -> dict | None:
        moment = now or datetime.now()
        with self._connect() as conn:
            self._expire_stale_holds(conn, now=moment)
            row = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise "
                "WHERE kp_id = ? AND kind = 'hold' "
                "ORDER BY id DESC LIMIT 1",
                (int(kp_id),),
            ).fetchone()
            return None if row is None else self._hold_from_row(conn, row)

    def insert_hold(
        self,
        *,
        kp_id: int,
        tracks_total: int,
        promised_date: date,
        allocations: Sequence[tuple[date, int]],
        created_by: str,
        created_at: datetime,
        expires_at: datetime,
        now: datetime | None = None,
    ) -> dict:
        moment = now or created_at
        with self._connect() as conn:
            self._expire_stale_holds(conn, now=moment)
            conn.execute(
                "UPDATE kp_promise SET status = 'released' "
                "WHERE kp_id = ? AND kind = 'hold' AND status = 'active'",
                (int(kp_id),),
            )
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO kp_promise (
                    kp_id, tracks_total, promised_date, kind, status,
                    created_by, created_at, expires_at
                ) VALUES (?, ?, ?, 'hold', 'active', ?, ?, ?)
                """,
                (
                    int(kp_id),
                    int(tracks_total),
                    promised_date.isoformat(),
                    created_by,
                    _iso_dt(created_at),
                    _iso_dt(expires_at),
                ),
            )
            promise_id = int(cur.lastrowid)
            for week_start, tracks in allocations:
                cur.execute(
                    """
                    INSERT INTO kp_promise_alloc (
                        promise_id, week_start, tracks, status
                    ) VALUES (?, ?, ?, 'active')
                    """,
                    (promise_id, week_start.isoformat(), int(tracks)),
                )
            row = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise WHERE id = ?",
                (promise_id,),
            ).fetchone()
            return self._hold_from_row(conn, row)

    @contextmanager
    def _using(self, external):
        if external is not None:
            yield external
            return
        conn = self._connect()
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def list_wizard_promise_allocs(
        self,
        *,
        week_starts: Sequence[date] | None = None,
        include_all_overdue: bool = True,
    ) -> list[dict]:
        """Active/overdue promise allocations for the planner wizard.

        Holds and consumed/released/expired promises are excluded. When
        ``week_starts`` is set, return allocs of those weeks; overdue
        allocs are still included if ``include_all_overdue`` is true.
        """
        sql = """
            SELECT p.kp_id, p.promised_date, p.tracks_total,
                   a.week_start, a.tracks, a.status
            FROM kp_promise_alloc a
            JOIN kp_promise p ON p.id = a.promise_id
            WHERE p.kind = 'promise'
              AND p.status = 'active'
              AND a.status IN ('active', 'overdue')
        """
        params: list[object] = []
        if week_starts is not None:
            weeks = [
                week.isoformat() if isinstance(week, date) else str(week)[:10]
                for week in week_starts
            ]
            if weeks:
                placeholders = ",".join("?" * len(weeks))
                if include_all_overdue:
                    sql += (
                        f" AND (a.week_start IN ({placeholders})"
                        " OR a.status = 'overdue')"
                    )
                else:
                    sql += f" AND a.week_start IN ({placeholders})"
                params.extend(weeks)
            elif include_all_overdue:
                sql += " AND a.status = 'overdue'"
            else:
                return []
        sql += " ORDER BY a.week_start, p.kp_id, a.id"

        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [
            {
                "kp_id": int(row[0]),
                "promised_date": date.fromisoformat(str(row[1])),
                "tracks_total": int(row[2]),
                "week_start": date.fromisoformat(str(row[3])),
                "tracks": int(row[4]),
                "status": str(row[5]),
            }
            for row in rows
        ]

    def list_overdue_allocations(self, *, _external_conn=None) -> list[dict]:
        sql = """
            SELECT a.id, a.promise_id, p.kp_id, a.week_start, a.tracks
            FROM kp_promise_alloc a
            JOIN kp_promise p ON p.id = a.promise_id
            WHERE p.kind = 'promise'
              AND a.status = 'overdue'
            ORDER BY a.week_start, p.kp_id, a.id
        """
        with self._using(_external_conn) as conn:
            rows = conn.execute(sql).fetchall()
        return [
            {
                "alloc_id": int(row[0]),
                "promise_id": int(row[1]),
                "kp_id": int(row[2]),
                "week_start": date.fromisoformat(str(row[3])),
                "tracks": int(row[4]),
            }
            for row in rows
        ]

    def apply_plan_commit_settlement(
        self,
        *,
        entered_kp_ids: Collection[int],
        covered_weeks: Sequence[date],
        _external_conn=None,
    ) -> dict[str, tuple[int, ...]]:
        """Consume entered-week allocs; overdue missed ones. Caller owns tx if given."""
        weeks = [week.isoformat() if isinstance(week, date) else str(week)[:10]
                 for week in covered_weeks]
        entered = {int(kp_id) for kp_id in entered_kp_ids}
        empty = {
            "consumed_alloc_ids": (),
            "overdue_alloc_ids": (),
            "consumed_promise_ids": (),
        }
        if not weeks:
            return empty

        placeholders = ",".join("?" * len(weeks))
        with self._using(_external_conn) as conn:
            alloc_rows = conn.execute(
                f"""
                SELECT a.id, a.promise_id, p.kp_id
                FROM kp_promise_alloc a
                JOIN kp_promise p ON p.id = a.promise_id
                WHERE p.kind = 'promise'
                  AND p.status = 'active'
                  AND a.status = 'active'
                  AND a.week_start IN ({placeholders})
                """,
                weeks,
            ).fetchall()
            consume_ids: list[int] = []
            overdue_ids: list[int] = []
            affected: set[int] = set()
            for alloc_id, promise_id, kp_id in alloc_rows:
                affected.add(int(promise_id))
                if int(kp_id) in entered:
                    consume_ids.append(int(alloc_id))
                else:
                    overdue_ids.append(int(alloc_id))

            self._set_alloc_status(conn, consume_ids, "consumed")
            self._set_alloc_status(conn, overdue_ids, "overdue")

            if entered:
                leftover = conn.execute(
                    f"""
                    SELECT a.id, a.promise_id, p.kp_id
                    FROM kp_promise_alloc a
                    JOIN kp_promise p ON p.id = a.promise_id
                    WHERE p.kind = 'promise'
                      AND p.status = 'active'
                      AND a.status = 'active'
                      AND p.kp_id IN ({",".join("?" * len(entered))})
                      AND NOT EXISTS (
                          SELECT 1 FROM kp_plates pl
                          WHERE pl.kp_id = p.kp_id
                            AND pl.status = 'в производстве'
                            AND pl.qty > 0
                      )
                    """,
                    tuple(entered),
                ).fetchall()
                leftover_ids = [int(row[0]) for row in leftover]
                for _alloc_id, promise_id, _kp_id in leftover:
                    affected.add(int(promise_id))
                self._set_alloc_status(conn, leftover_ids, "consumed")
                consume_ids.extend(leftover_ids)

            consumed_promises = self._consume_settled_promises(conn, affected)
            return {
                "consumed_alloc_ids": tuple(consume_ids),
                "overdue_alloc_ids": tuple(overdue_ids),
                "consumed_promise_ids": consumed_promises,
            }

    def record_exclusions(
        self,
        *,
        plan_id: str,
        items: Sequence[tuple[int, date, str]],
        excluded_by: str,
        created_at: datetime,
        _external_conn=None,
    ) -> list[dict]:
        """Write exclusion rows + owner notifications in one tx."""
        moment = _iso_dt(created_at)
        written: list[dict] = []
        with self._using(_external_conn) as conn:
            for kp_id, week_start, reason in items:
                cur = conn.execute(
                    """
                    INSERT INTO kp_promise_exclusion (
                        kp_id, plan_id, week_start, reason, excluded_by, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(kp_id),
                        plan_id,
                        week_start.isoformat(),
                        reason,
                        excluded_by,
                        moment,
                    ),
                )
                exclusion_id = int(cur.lastrowid)
                owner_id = self._promise_owner_user_id(conn, int(kp_id))
                notification_id = None
                if owner_id is not None:
                    payload = json.dumps(
                        {
                            "kp_id": int(kp_id),
                            "week_start": week_start.isoformat(),
                            "reason": reason,
                        },
                        ensure_ascii=False,
                    )
                    note = conn.execute(
                        """
                        INSERT INTO notifications (
                            user_id, kind, payload_json, created_at
                        ) VALUES (?, 'promise_excluded', ?, ?)
                        """,
                        (owner_id, payload, moment),
                    )
                    notification_id = int(note.lastrowid)
                written.append(
                    {
                        "exclusion_id": exclusion_id,
                        "notification_id": notification_id,
                        "kp_id": int(kp_id),
                        "week_start": week_start,
                        "reason": reason,
                        "excluded_by": excluded_by,
                        "plan_id": plan_id,
                    }
                )
        return written

    def list_exclusions(
        self,
        *,
        plan_id: str | None = None,
        kp_id: int | None = None,
        _external_conn=None,
    ) -> list[dict]:
        sql = """
            SELECT id, kp_id, plan_id, week_start, reason, excluded_by, created_at
            FROM kp_promise_exclusion
            WHERE 1 = 1
        """
        params: list[object] = []
        if plan_id is not None:
            sql += " AND plan_id = ?"
            params.append(plan_id)
        if kp_id is not None:
            sql += " AND kp_id = ?"
            params.append(int(kp_id))
        sql += " ORDER BY id"
        with self._using(_external_conn) as conn:
            rows = conn.execute(sql, params).fetchall()
        return [
            {
                "id": int(row[0]),
                "kp_id": int(row[1]),
                "plan_id": str(row[2]),
                "week_start": date.fromisoformat(str(row[3])),
                "reason": str(row[4]),
                "excluded_by": None if row[5] is None else str(row[5]),
                "created_at": str(row[6]),
            }
            for row in rows
        ]

    def list_notifications(
        self,
        *,
        user_id: int | None = None,
        kind: str | None = None,
        unread_only: bool = False,
        _external_conn=None,
    ) -> list[dict]:
        sql = """
            SELECT id, user_id, kind, payload_json, read_at, created_at
            FROM notifications
            WHERE 1 = 1
        """
        params: list[object] = []
        if user_id is not None:
            sql += " AND user_id = ?"
            params.append(int(user_id))
        if kind is not None:
            sql += " AND kind = ?"
            params.append(kind)
        if unread_only:
            sql += " AND read_at IS NULL"
        sql += " ORDER BY id"
        with self._using(_external_conn) as conn:
            rows = conn.execute(sql, params).fetchall()
        return [
            {
                "id": int(row[0]),
                "user_id": int(row[1]),
                "kind": str(row[2]),
                "payload_json": None if row[3] is None else str(row[3]),
                "read_at": None if row[4] is None else str(row[4]),
                "created_at": str(row[5]),
            }
            for row in rows
        ]

    def count_unread_notifications(self, *, user_id: int, _external_conn=None) -> int:
        with self._using(_external_conn) as conn:
            row = conn.execute(
                """
                SELECT COUNT(*) FROM notifications
                WHERE user_id = ? AND read_at IS NULL
                """,
                (int(user_id),),
            ).fetchone()
        return 0 if row is None else int(row[0])

    def mark_notification_read(
        self,
        notification_id: int,
        *,
        user_id: int,
        read_at: datetime,
        _external_conn=None,
    ) -> dict | None:
        moment = _iso_dt(read_at)
        with self._using(_external_conn) as conn:
            row = conn.execute(
                """
                SELECT id, user_id, kind, payload_json, read_at, created_at
                FROM notifications
                WHERE id = ? AND user_id = ?
                """,
                (int(notification_id), int(user_id)),
            ).fetchone()
            if row is None:
                return None
            current_read = None if row[4] is None else str(row[4])
            if current_read is None:
                conn.execute(
                    """
                    UPDATE notifications SET read_at = ?
                    WHERE id = ? AND user_id = ?
                    """,
                    (moment, int(notification_id), int(user_id)),
                )
                current_read = moment
            return {
                "id": int(row[0]),
                "user_id": int(row[1]),
                "kind": str(row[2]),
                "payload_json": None if row[3] is None else str(row[3]),
                "read_at": current_read,
                "created_at": str(row[5]),
            }

    def _promise_owner_user_id(self, conn, kp_id: int) -> int | None:
        row = conn.execute(
            """
            SELECT created_by FROM kp_promise
            WHERE kp_id = ? AND kind = 'promise'
              AND status IN ('active', 'consumed')
            ORDER BY id DESC LIMIT 1
            """,
            (kp_id,),
        ).fetchone()
        resolved = self._user_id_from_actor(conn, None if row is None else row[0])
        if resolved is not None:
            return resolved
        meta = conn.execute(
            "SELECT owner_user_id FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
        if meta is None or meta[0] is None:
            return None
        return int(meta[0])

    def _user_id_from_actor(self, conn, created_by: object) -> int | None:
        if created_by is None:
            return None
        actor = str(created_by).strip()
        if not actor:
            return None
        if actor.isdigit():
            return int(actor)
        try:
            row = conn.execute(
                "SELECT id FROM app_users WHERE username = ?",
                (actor,),
            ).fetchone()
        except sqlite3.OperationalError:
            return None
        return None if row is None else int(row[0])

    def _set_alloc_status(self, conn, alloc_ids: Sequence[int], status: str) -> None:
        if not alloc_ids:
            return
        conn.execute(
            f"UPDATE kp_promise_alloc SET status = ? WHERE id IN ({','.join('?' * len(alloc_ids))})",
            (status, *alloc_ids),
        )

    def _consume_settled_promises(
        self, conn, promise_ids: Collection[int]
    ) -> tuple[int, ...]:
        if not promise_ids:
            return ()
        ids = tuple(int(pid) for pid in promise_ids)
        conn.execute(
            f"""
            UPDATE kp_promise
            SET status = 'consumed'
            WHERE id IN ({",".join("?" * len(ids))})
              AND kind = 'promise'
              AND status = 'active'
              AND NOT EXISTS (
                  SELECT 1 FROM kp_promise_alloc a
                  WHERE a.promise_id = kp_promise.id
                    AND a.status != 'consumed'
              )
              AND NOT EXISTS (
                  SELECT 1 FROM kp_plates pl
                  WHERE pl.kp_id = kp_promise.kp_id
                    AND pl.status = 'в производстве'
                    AND pl.qty > 0
              )
            """,
            ids,
        )
        rows = conn.execute(
            f"""
            SELECT id FROM kp_promise
            WHERE id IN ({",".join("?" * len(ids))}) AND status = 'consumed'
            """,
            ids,
        ).fetchall()
        return tuple(int(row[0]) for row in rows)

    def get_active_promise(self, kp_id: int, *, _external_conn=None) -> dict | None:
        """Latest active hard promise (not hold) for the KP."""
        with self._using(_external_conn) as conn:
            row = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise "
                "WHERE kp_id = ? AND kind = 'promise' AND status = 'active' "
                "ORDER BY id DESC LIMIT 1",
                (int(kp_id),),
            ).fetchone()
            return None if row is None else self._hold_from_row(conn, row)

    def release_active_for_kp(
        self,
        kp_id: int,
        *,
        kinds: Sequence[str] | None = None,
        _external_conn=None,
    ) -> tuple[int, ...]:
        """Set active promise/hold rows to released. Alloc sums then ignore them."""
        wanted = tuple(kinds) if kinds else ("promise", "hold")
        placeholders = ",".join("?" * len(wanted))
        with self._using(_external_conn) as conn:
            rows = conn.execute(
                f"""
                SELECT id FROM kp_promise
                WHERE kp_id = ? AND status = 'active' AND kind IN ({placeholders})
                ORDER BY id
                """,
                (int(kp_id), *wanted),
            ).fetchall()
            ids = tuple(int(row[0]) for row in rows)
            if ids:
                conn.execute(
                    f"UPDATE kp_promise SET status = 'released' "
                    f"WHERE id IN ({','.join('?' * len(ids))})",
                    ids,
                )
            return ids

    def apply_promise_recalc(
        self,
        *,
        promise_id: int,
        kp_id: int,
        tracks_total: int,
        promised_date: date,
        allocations: Sequence[tuple[date, int]],
        created_at: datetime,
        notify_kind: str | None = None,
        notify_payload: dict | None = None,
        _external_conn=None,
    ) -> int | None:
        """Rewrite active promise window; optionally notify owner in the same tx."""
        with self._using(_external_conn) as conn:
            conn.execute(
                """
                UPDATE kp_promise
                SET tracks_total = ?, promised_date = ?
                WHERE id = ? AND kind = 'promise' AND status = 'active'
                """,
                (
                    int(tracks_total),
                    promised_date.isoformat(),
                    int(promise_id),
                ),
            )
            conn.execute(
                "DELETE FROM kp_promise_alloc WHERE promise_id = ?",
                (int(promise_id),),
            )
            for week_start, tracks in allocations:
                conn.execute(
                    """
                    INSERT INTO kp_promise_alloc (
                        promise_id, week_start, tracks, status
                    ) VALUES (?, ?, ?, 'active')
                    """,
                    (
                        int(promise_id),
                        week_start.isoformat(),
                        int(tracks),
                    ),
                )
            if not notify_kind or notify_payload is None:
                return None
            owner_id = self._promise_owner_user_id(conn, int(kp_id))
            if owner_id is None:
                return None
            return self._insert_notification(
                conn,
                user_id=owner_id,
                kind=notify_kind,
                payload=notify_payload,
                created_at=created_at,
            )

    def _insert_notification(
        self,
        conn,
        *,
        user_id: int,
        kind: str,
        payload: dict,
        created_at: datetime,
    ) -> int:
        cur = conn.execute(
            """
            INSERT INTO notifications (user_id, kind, payload_json, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                int(user_id),
                kind,
                json.dumps(payload, ensure_ascii=False),
                _iso_dt(created_at),
            ),
        )
        return int(cur.lastrowid)

    def release_hold(self, kp_id: int, *, now: datetime | None = None) -> dict | None:
        moment = now or datetime.now()
        with self._connect() as conn:
            self._expire_stale_holds(conn, now=moment)
            row = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise "
                "WHERE kp_id = ? AND kind = 'hold' AND status = 'active' "
                "ORDER BY id DESC LIMIT 1",
                (int(kp_id),),
            ).fetchone()
            if row is None:
                return None
            conn.execute(
                "UPDATE kp_promise SET status = 'released' WHERE id = ?",
                (int(row[0]),),
            )
            updated = conn.execute(
                f"SELECT {_HOLD_COLS} FROM kp_promise WHERE id = ?",
                (int(row[0]),),
            ).fetchone()
            return self._hold_from_row(conn, updated)

    def _expire_stale_holds(self, conn, *, now: datetime) -> int:
        cur = conn.execute(
            """
            UPDATE kp_promise
            SET status = 'expired'
            WHERE kind = 'hold'
              AND status = 'active'
              AND expires_at IS NOT NULL
              AND expires_at < ?
            """,
            (_iso_dt(now),),
        )
        return int(cur.rowcount)

    def _sum_allocs_by_week(
        self,
        *,
        kind: str,
        exclude_kp_id: int | None = None,
    ) -> dict[date, int]:
        sql = """
            SELECT a.week_start, SUM(a.tracks)
            FROM kp_promise_alloc a
            JOIN kp_promise p ON p.id = a.promise_id
            WHERE p.kind = ?
              AND p.status = 'active'
              AND a.status = 'active'
        """
        params: list[object] = [kind]
        if exclude_kp_id is not None:
            sql += " AND p.kp_id != ?"
            params.append(int(exclude_kp_id))
        sql += " GROUP BY a.week_start"

        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
        result: dict[date, int] = {}
        for week_raw, tracks in rows:
            result[date.fromisoformat(str(week_raw))] = int(tracks)
        return result

    def _hold_from_row(self, conn, row) -> dict:
        promise_id = int(row[0])
        alloc_rows = conn.execute(
            """
            SELECT week_start, tracks FROM kp_promise_alloc
            WHERE promise_id = ? ORDER BY week_start
            """,
            (promise_id,),
        ).fetchall()
        return {
            "id": promise_id,
            "kp_id": int(row[1]),
            "tracks_total": int(row[2]),
            "promised_date": str(row[3]),
            "kind": str(row[4]),
            "status": str(row[5]),
            "created_by": None if row[6] is None else str(row[6]),
            "created_at": str(row[7]),
            "expires_at": None if row[8] is None else str(row[8]),
            "allocations": [
                (date.fromisoformat(str(week_raw)), int(tracks))
                for week_raw, tracks in alloc_rows
            ],
        }
