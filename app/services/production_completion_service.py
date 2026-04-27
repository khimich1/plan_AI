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
        actor: str | None = None,
    ) -> dict[str, Any]:
        plan = self.plan_repository.load_plan(plan_id)
        if not plan:
            raise ProductionCompletionError("Plan not found")

        day_number = self._get_day_number(plan, target_date)
        day_view = build_day_view_detail(target_date)
        (
            plates_by_kp,
            rejected_by_kp,
            completion_stats,
        ) = self._collect_plates_by_kp(
            day_view,
            plan_id,
            rejected_plates or [],
        )

        planned_qty_total = completion_stats["planned_qty_total"]
        completed_requested_qty = completion_stats["completed_requested_qty"]
        rejected_requested_qty = completion_stats["rejected_requested_qty"]
        skipped_without_kp = completion_stats["skipped_without_kp"]
        if planned_qty_total <= 0:
            raise ProductionCompletionError(
                "В выбранном плане на дату нет плит для завершения."
            )
        if skipped_without_kp:
            raise ProductionCompletionError(
                "Не удалось завершить день: у части плит нет привязки к КП "
                f"(позиций: {len(skipped_without_kp)})."
            )

        total_moved = 0
        unmoved_plates: list[dict[str, Any]] = []
        for kp_id, plates in plates_by_kp.items():
            move_result = kp_db.move_plates_to_completed(
                kp_id,
                plates,
                day_number,
                self.db_path,
                plan_ids=[plan_id],
                actor=actor,
                return_unmoved=True,
            )
            if isinstance(move_result, tuple):
                moved, unmoved = move_result
            else:
                moved = int(move_result or 0)
                # Backward compatibility for tests/mocks that return plain int.
                # In this branch we don't know exact DB misses, so use requested
                # payload as best-effort details for user-facing error.
                unmoved = [
                    {
                        "kp_id": int(kp_id),
                        "plate_name": str(plate.get("plate_name") or ""),
                        "qty": int(plate.get("qty") or 0),
                        "length_m": float(plate.get("length_m") or 0),
                        "width_m": float(plate.get("width_m") or 0),
                        "load_class": int(plate.get("load_class") or 0),
                    }
                    for plate in plates
                    if int(plate.get("qty") or 0) > 0
                ]
            total_moved += moved
            unmoved_plates.extend(unmoved)

        # Бракованные плиты возвращаем в 'в производстве', чтобы они попали
        # в следующее планирование. Иначе строки залипают со status='в плане'
        # и мастер планирования рапортует «Не найдено плит для планирования».
        rejected_flat: list[dict[str, Any]] = [
            {"kp_id": kp_id, "plate_name": item["plate_name"], "qty": item["qty"]}
            for kp_id, items in rejected_by_kp.items()
            for item in items
        ]
        rejected_returned = self._return_rejected(
            rejected_flat, self.db_path, actor=actor
        )

        if total_moved < completed_requested_qty:
            missing_qty = completed_requested_qty - total_moved
            missing_details = self._format_unmoved_plates(unmoved_plates)
            raise ProductionCompletionError(
                "Не удалось завершить день: не списано "
                f"{missing_qty} плит из "
                f"{completed_requested_qty}. "
                f"Не хватает: {missing_details}."
            )
        if rejected_returned < rejected_requested_qty:
            raise ProductionCompletionError(
                "Не удалось завершить день: не возвращено в производство "
                f"{rejected_requested_qty - rejected_returned} бракованных плит "
                f"из {rejected_requested_qty}."
            )

        # Проверка автозавершения КП — ТОЛЬКО после возврата брака,
        # иначе КП с полностью забракованным днём ошибочно станет 'выполнено'.
        completed_kps: list[int] = []
        affected_kp_ids = set(plates_by_kp.keys()) | set(rejected_by_kp.keys())
        for kp_id in affected_kp_ids:
            if kp_db.check_and_update_kp_completion(kp_id, self.db_path):
                completed_kps.append(kp_id)

        return {
            "moved_plates": total_moved,
            "rejected_returned": rejected_returned,
            "completed_kps": sorted(set(completed_kps)),
            "affected_kps": sorted(affected_kp_ids),
            "day_number": day_number,
            **completion_stats,
        }

    @staticmethod
    def _format_unmoved_plates(
        unmoved_plates: list[dict[str, Any]],
        *,
        max_items: int = 8,
    ) -> str:
        grouped: dict[tuple[int, str], int] = defaultdict(int)
        for item in unmoved_plates:
            kp_id = int(item.get("kp_id") or 0)
            plate_name = str(item.get("plate_name") or "").strip() or "Без названия"
            qty = int(item.get("qty") or 0)
            if qty <= 0:
                continue
            grouped[(kp_id, plate_name)] += qty

        if not grouped:
            return "не удалось определить позиции"

        parts = [
            f"КП {kp_id}: {plate_name} — {qty} шт"
            for (kp_id, plate_name), qty in sorted(
                grouped.items(),
                key=lambda item: (-item[1], item[0][0], item[0][1]),
            )
        ]
        if len(parts) > max_items:
            hidden_count = len(parts) - max_items
            parts = parts[:max_items] + [f"... и ещё {hidden_count} поз."]
        return "; ".join(parts)

    @staticmethod
    def _return_rejected(
        rejected: list[dict[str, Any]],
        db_path: str,
        *,
        actor: str | None = None,
    ) -> int:
        """Возвращает в 'в производстве' все позиции, помеченные как брак.

        Контракт ``rejected``: список словарей ``{kp_id, plate_name, qty}``.
        Невалидные элементы (без ``kp_id``/``plate_name`` или с qty<=0) пропускаются.
        Возвращает суммарное qty, которое БД успешно перевела в 'в производстве'.

        Используется и web-сервисом, и Telegram-ботом — единая точка возврата брака,
        чтобы поведение «брак → следующее планирование» было идентичным.
        ``actor`` прокидывается в audit-лог ``plate_status_log``.
        """
        total = 0
        for item in rejected:
            kp_id = item.get("kp_id")
            plate_name = item.get("plate_name")
            qty = int(item.get("qty") or 0)
            if not kp_id or not plate_name or qty <= 0:
                continue
            ok = kp_db.return_plates_to_production(
                kp_id=int(kp_id),
                plate_name=plate_name,
                qty=qty,
                db_path=db_path,
                actor=actor,
                reason="rejected",
            )
            if ok:
                total += qty
        return total

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
    ) -> tuple[
        dict[int, list[dict[str, Any]]],
        dict[int, list[dict[str, Any]]],
        dict[str, Any],
    ]:
        grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
        rejected_grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
        if not day_view:
            raise ProductionCompletionError("Day not found")

        rejected_by_position = ProductionCompletionService._build_rejection_map(
            rejected_plates
        )
        seen_positions: set[tuple[int, int]] = set()
        rejected_positions = 0
        rejected_qty_total = 0
        planned_qty_total = 0
        completed_requested_qty = 0
        skipped_without_kp: list[dict[str, Any]] = []
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
                    if total_qty <= 0:
                        continue
                    planned_qty_total += total_qty

                    completed_qty = total_qty - reject_qty
                    if reject_qty:
                        rejected_positions += 1
                        rejected_qty_total += reject_qty

                    kp_id = plate.get("kp_id")
                    if not kp_id:
                        skipped_without_kp.append(
                            {
                                "track_number": track_number,
                                "plate_index": plate_index,
                                "plate_name": plate.get("plate_name") or "",
                                "qty": total_qty,
                            }
                        )
                        continue
                    if reject_qty and kp_id:
                        rejected_grouped[int(kp_id)].append(
                            {
                                "plate_name": plate.get("plate_name") or "",
                                "qty": int(reject_qty),
                            }
                        )

                    if completed_qty <= 0:
                        continue
                    completed_requested_qty += completed_qty
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

        return (
            grouped,
            rejected_grouped,
            {
                "planned_qty_total": planned_qty_total,
                "completed_requested_qty": completed_requested_qty,
                "rejected_requested_qty": rejected_qty_total,
                "rejected_plates": rejected_qty_total,
                "rejected_positions": rejected_positions,
                "skipped_without_kp": skipped_without_kp,
                "skipped_without_kp_count": len(skipped_without_kp),
            },
        )

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
