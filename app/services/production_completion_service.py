from __future__ import annotations

from collections import defaultdict
from typing import Any

from app.repositories.plan_repository import PlanRepository
from app.services.day_view_service import build_day_view_detail
from core import kp_db


class ProductionCompletionError(ValueError):
    """Ошибка валидации данных при завершении производственного дня."""


class ProductionCompletionService:
    """Списывает плиты завершённого дня из плана в SQLite-учёт выполнения."""

    def __init__(
        self,
        *,
        db_path: str,
        plan_repository: PlanRepository | None = None,
    ) -> None:
        self.db_path = db_path
        self.plan_repository = plan_repository or PlanRepository()

    def complete_day(
        self,
        *,
        plan_id: str,
        target_date: str,
        rejected_plates: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        plan = self.plan_repository.load_plan(plan_id)
        if not plan:
            raise ProductionCompletionError("Plan not found")

        day_number = self._get_day_number(plan, target_date)
        day_view = build_day_view_detail(target_date)
        plates_by_kp, rejection_stats = self._collect_plates_by_kp(
            day_view,
            plan_id,
            rejected_plates or [],
        )

        total_moved = 0
        completed_kps: list[int] = []

        for kp_id, plates in plates_by_kp.items():
            moved = kp_db.move_plates_to_completed(
                kp_id,
                plates,
                day_number,
                self.db_path,
                plan_ids=[plan_id],
            )
            total_moved += moved

            if kp_db.check_and_update_kp_completion(kp_id, self.db_path):
                completed_kps.append(kp_id)

        return {
            "moved_plates": total_moved,
            "completed_kps": sorted(set(completed_kps)),
            "affected_kps": sorted(plates_by_kp.keys()),
            "day_number": day_number,
            **rejection_stats,
        }

    @staticmethod
    def _get_day_number(plan: dict | None, target_date: str) -> int:
        if not plan:
            return 1
        day = (plan.get("days") or {}).get(target_date) or {}
        return int(day.get("day_number") or 1)

    @staticmethod
    def _collect_plates_by_kp(
        day_view: dict[str, Any] | None,
        plan_id: str,
        rejected_plates: list[dict[str, Any]],
    ) -> tuple[dict[int, list[dict[str, Any]]], dict[str, int]]:
        grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
        if not day_view:
            raise ProductionCompletionError("Day not found")

        rejected_by_position = ProductionCompletionService._build_rejection_map(
            rejected_plates
        )
        seen_positions: set[tuple[int, int]] = set()
        rejected_positions = 0
        rejected_qty_total = 0
        plan_found = False

        for block in day_view.get("plans") or []:
            if block.get("plan_id") != plan_id:
                continue
            plan_found = True
            for track in block.get("tracks") or []:
                track_number = int(track.get("track_number") or 0)
                for plate_index, plate in enumerate(track.get("plates_info") or []):
                    position = (track_number, plate_index)
                    seen_positions.add(position)

                    reject_qty = rejected_by_position.get(position, 0)
                    total_qty = int(plate.get("qty") or 0)
                    if reject_qty > total_qty:
                        raise ProductionCompletionError(
                            "Rejected quantity cannot exceed plate quantity"
                        )

                    completed_qty = total_qty - reject_qty
                    if reject_qty:
                        rejected_positions += 1
                        rejected_qty_total += reject_qty
                    if completed_qty <= 0:
                        continue

                    kp_id = plate.get("kp_id")
                    if not kp_id:
                        continue
                    plate_to_move = {**plate, "qty": completed_qty}
                    grouped[int(kp_id)].append(
                        ProductionCompletionService._to_completed_plate_payload(
                            plate_to_move
                        )
                    )

        if not plan_found:
            raise ProductionCompletionError("Plan not found for selected day")

        unknown_positions = set(rejected_by_position) - seen_positions
        if unknown_positions:
            raise ProductionCompletionError("Rejected plate position not found")

        return grouped, {
            "rejected_plates": rejected_qty_total,
            "rejected_positions": rejected_positions,
        }

    @staticmethod
    def _build_rejection_map(
        rejected_plates: list[dict[str, Any]],
    ) -> dict[tuple[int, int], int]:
        rejected_by_position: dict[tuple[int, int], int] = defaultdict(int)
        for item in rejected_plates:
            track_number = int(item.get("track_number") or 0)
            plate_index = int(item.get("plate_index") or 0)
            qty = int(item.get("qty") or 0)
            if track_number < 1 or plate_index < 0 or qty < 0:
                raise ProductionCompletionError("Invalid rejected plate payload")
            if qty == 0:
                continue
            rejected_by_position[(track_number, plate_index)] += qty
        return dict(rejected_by_position)

    @staticmethod
    def _to_completed_plate_payload(plate: dict[str, Any]) -> dict[str, Any]:
        load_code = int(plate.get("load_code") or 8)
        return {
            "plate_name": plate.get("plate_name") or "",
            "length_m": float(plate.get("length_m") or 0),
            "width_m": float(plate.get("width_mm") or 0) / 1000.0,
            "load_class": load_code * 100,
            "qty": int(plate.get("qty") or 0),
            "kp_id": int(plate["kp_id"]),
        }
