"""Protocols for production planning persistence (no app imports)."""

from __future__ import annotations

from typing import Any, Protocol


class PlanPersistPort(Protocol):
    """Minimal repository surface used by the planning persist phase."""

    def get_global_occupancy(
        self, *, exclude_plan_id: str | None = None
    ) -> dict[str, int]: ...

    def build_plan_from_tracks(
        self,
        *,
        plan_id: str | None,
        new_tracks_list: list,
        start_date: str,
        tracks_per_day: int,
        plate_lookup_exact: dict | None = None,
        plate_lookup_by_length: dict | None = None,
        orders_2d: list | None = None,
        optimization_result: dict | None = None,
        plan_name: str | None = None,
        global_occupancy: dict[str, int] | None = None,
        precomputed_tracks_by_day: dict[str, list] | None = None,
        auto_save: bool = False,
    ) -> tuple[dict, dict]: ...

    def get(self, plan_id: str) -> dict[str, Any] | None: ...

    def create(self, payload: dict[str, Any]) -> dict[str, Any]: ...

    def save(
        self, payload: dict[str, Any], expected_version: int
    ) -> dict[str, Any]: ...

    def set_active(self, plan_id: str) -> None: ...
