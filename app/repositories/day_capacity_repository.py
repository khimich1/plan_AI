"""SQL persistence for per-day production capacity overrides."""

from __future__ import annotations

from datetime import date, datetime
from typing import Union

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.production.capacity import TRACKS_PER_DAY_HARD_CAP, clamp_day_max
from core.production_capacity import TRACKS_PER_DAY_DEFAULT

DateLike = Union[date, str]


class DayCapacityRepository:
    """CRUD for ``day_capacity_override``; default tracks from ``TRACKS_PER_DAY_DEFAULT``."""

    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def _connect(self):
        ensure_schema(self.db_path)
        return _connect(self.db_path)

    @staticmethod
    def _to_iso_date(value: DateLike) -> str:
        if isinstance(value, date):
            return value.isoformat()
        return date.fromisoformat(str(value)).isoformat()

    def get_max_tracks(self, day: DateLike) -> int:
        day_iso = self._to_iso_date(day)
        with self._connect() as conn:
            row = conn.execute(
                "SELECT max_tracks FROM day_capacity_override WHERE date = ?",
                (day_iso,),
            ).fetchone()
        if row is None:
            return clamp_day_max(int(TRACKS_PER_DAY_DEFAULT))
        return clamp_day_max(int(row[0]))

    def list_overrides(self) -> dict[date, int]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT date, max_tracks FROM day_capacity_override ORDER BY date"
            ).fetchall()
        return {
            date.fromisoformat(str(row[0])): clamp_day_max(int(row[1])) for row in rows
        }

    def set_override(
        self,
        day: DateLike,
        max_tracks: int,
        updated_by: str | None = None,
    ) -> None:
        if max_tracks < 0:
            raise ValueError("max_tracks must be >= 0")
        if max_tracks > TRACKS_PER_DAY_HARD_CAP:
            raise ValueError(
                f"max_tracks must be <= {TRACKS_PER_DAY_HARD_CAP} (hard cap)"
            )
        day_iso = self._to_iso_date(day)
        updated_at = datetime.now().isoformat(timespec="seconds")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO day_capacity_override (date, max_tracks, updated_at, updated_by)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(date) DO UPDATE SET
                    max_tracks = excluded.max_tracks,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
                """,
                (day_iso, int(max_tracks), updated_at, updated_by),
            )
            conn.commit()
