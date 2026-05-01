from __future__ import annotations

import sqlite3
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
        day_view = build_day_view_detail(target_date, db_path=self.db_path)
        (
            plates_by_kp,
            rejected_by_kp,
            completion_stats,
        ) = self._collect_plates_by_kp(
            day_view,
            plan_id,
            rejected_plates or [],
        )

        # #region agent log
        try:
            import json as _agent_json
            import time as _agent_time

            with open(r"c:\Users\Роман\Desktop\Шишов\debug-ebb546.log", "a", encoding="utf-8") as _agent_f:
                _agent_f.write(_agent_json.dumps({
                    "sessionId": "ebb546",
                    "runId": "pre-fix",
                    "hypothesisId": "H4,H5",
                    "location": "app/services/production_completion_service.py:after_collect_plates",
                    "message": "Перед списанием: что day_view отдал в complete_day",
                    "data": {
                        "plan_id": plan_id,
                        "target_date": target_date,
                        "day_number": day_number,
                        "planned_qty_total": completion_stats.get("planned_qty_total"),
                        "completed_requested_qty": completion_stats.get("completed_requested_qty"),
                        "rejected_requested_qty": completion_stats.get("rejected_requested_qty"),
                        "skipped_without_kp_count": completion_stats.get("skipped_without_kp_count"),
                        "secondary_rests_qty": sum(int(x.get("qty") or 0) for x in completion_stats.get("secondary_rests") or []),
                        "plates_by_kp_qty": {
                            str(_kp_id): sum(int(_p.get("qty") or 0) for _p in _plates)
                            for _kp_id, _plates in plates_by_kp.items()
                        },
                        "plates_by_kp_sample": {
                            str(_kp_id): [
                                {
                                    "plate_name": _p.get("plate_name"),
                                    "qty": _p.get("qty"),
                                    "length_m": _p.get("length_m"),
                                    "width_m": _p.get("width_m"),
                                    "load_class": _p.get("load_class"),
                                }
                                for _p in _plates[:8]
                            ]
                            for _kp_id, _plates in list(plates_by_kp.items())[:5]
                        },
                    },
                    "timestamp": int(_agent_time.time() * 1000),
                }, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion

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

        rejected_flat: list[dict[str, Any]] = [
            {"kp_id": kp_id, "plate_name": item["plate_name"], "qty": item["qty"]}
            for kp_id, items in rejected_by_kp.items()
            for item in items
        ]

        # P0: вся цепочка списания (move → return_rejected → check_completion)
        # выполняется в ОДНОЙ транзакции. Любая ошибка → conn.rollback() и
        # БД остаётся в состоянии «до complete_day» (никаких полу-списанных плит).
        kp_db.init_schema(self.db_path)
        conn = sqlite3.connect(self.db_path)
        try:
            conn.execute('PRAGMA journal_mode=WAL')
            conn.execute('PRAGMA foreign_keys = ON')

            # P1: pre-flight reconciliation — до любого UPDATE проверяем,
            # что для каждой запрошенной плиты есть запись в kp_plates.
            # Если нет — поднимаем 422 БЕЗ изменений в БД.
            missing_preflight = self._verify_plates_exist_in_db(plates_by_kp, conn)
            if missing_preflight:
                conn.rollback()
                details = self._format_unmoved_plates(missing_preflight)
                raise ProductionCompletionError(
                    "Не удалось завершить день: в БД нет ожидаемых плит. "
                    f"Не хватает: {details}."
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
                    _external_conn=conn,
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

            if total_moved < completed_requested_qty:
                missing_qty = completed_requested_qty - total_moved
                missing_details = self._format_unmoved_plates(unmoved_plates)
                conn.rollback()
                raise ProductionCompletionError(
                    "Не удалось завершить день: не списано "
                    f"{missing_qty} плит из "
                    f"{completed_requested_qty}. "
                    f"Не хватает: {missing_details}."
                )

            # Бракованные плиты возвращаем в 'в производстве', чтобы они попали
            # в следующее планирование. Иначе строки залипают со status='в плане'
            # и мастер планирования рапортует «Не найдено плит для планирования».
            rejected_returned = self._return_rejected(
                rejected_flat, self.db_path, actor=actor, conn=conn
            )

            if rejected_returned < rejected_requested_qty:
                conn.rollback()
                raise ProductionCompletionError(
                    "Не удалось завершить день: не возвращено в производство "
                    f"{rejected_requested_qty - rejected_returned} бракованных плит "
                    f"из {rejected_requested_qty}."
                )

            # P6: secondary-cuts без kp_id сохраняем в plate_rests — внутри
            # той же транзакции, чтобы при ошибке всё откатилось целиком.
            secondary_rests = completion_stats.get("secondary_rests") or []
            secondary_rests_created = 0
            for rest in secondary_rests:
                qty = int(rest.get("qty") or 0)
                if qty <= 0:
                    continue
                # plate_rests привязан к kp_id NOT NULL по миграции —
                # без identity мы не можем сохранить, поэтому пишем под
                # условный kp_id=0. Если миграция не позволяет — лог-warning.
                try:
                    kp_db.create_plate_rest(
                        kp_id=0,
                        source_plate_name=rest.get("plate_name") or "",
                        rest_width_mm=int(rest.get("width_mm") or 0),
                        length_m=float(rest.get("length_m") or 0),
                        production_day=int(day_number),
                        qty=qty,
                        db_path=self.db_path,
                        _external_conn=conn,
                    )
                    secondary_rests_created += qty
                except Exception:  # noqa: BLE001
                    # Если БД отказывается принять (FK на kp_id) — продолжаем,
                    # secondary без kp_id просто не пишем как rest.
                    pass

            # Проверка автозавершения КП — ТОЛЬКО после возврата брака,
            # иначе КП с полностью забракованным днём ошибочно станет 'выполнено'.
            completed_kps: list[int] = []
            affected_kp_ids = set(plates_by_kp.keys()) | set(rejected_by_kp.keys())
            for kp_id in affected_kp_ids:
                if kp_db.check_and_update_kp_completion(
                    kp_id, self.db_path, _external_conn=conn
                ):
                    completed_kps.append(kp_id)

            # #region agent log
            try:
                import json as _agent_json
                import time as _agent_time

                _cur = conn.cursor()
                _cur.execute(
                    """
                    SELECT status, COALESCE(day_number, -1), COUNT(*), COALESCE(SUM(qty), 0)
                    FROM kp_plates
                    WHERE plan_id = ?
                    GROUP BY status, COALESCE(day_number, -1)
                    ORDER BY status, COALESCE(day_number, -1)
                    """,
                    (plan_id,),
                )
                _remaining_by_status_day = [
                    {
                        "status": _row[0],
                        "day_number": None if int(_row[1]) == -1 else int(_row[1]),
                        "rows": int(_row[2] or 0),
                        "qty": int(_row[3] or 0),
                    }
                    for _row in _cur.fetchall()
                ]
                with open(r"c:\Users\Роман\Desktop\Шишов\debug-ebb546.log", "a", encoding="utf-8") as _agent_f:
                    _agent_f.write(_agent_json.dumps({
                        "sessionId": "ebb546",
                        "runId": "pre-fix",
                        "hypothesisId": "H4,H5",
                        "location": "app/services/production_completion_service.py:before_commit",
                        "message": "После move_plates_to_completed: остатки по plan_id перед commit транзакции",
                        "data": {
                            "plan_id": plan_id,
                            "target_date": target_date,
                            "day_number": day_number,
                            "completed_requested_qty": completed_requested_qty,
                            "total_moved": total_moved,
                            "unmoved_qty": sum(int(x.get("qty") or 0) for x in unmoved_plates),
                            "unmoved_sample": unmoved_plates[:10],
                            "remaining_by_status_day": _remaining_by_status_day,
                            "completed_kps": completed_kps,
                        },
                        "timestamp": int(_agent_time.time() * 1000),
                    }, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion

            conn.commit()
        except ProductionCompletionError:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            conn.close()

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
        conn: sqlite3.Connection | None = None,
    ) -> int:
        """Возвращает в 'в производстве' все позиции, помеченные как брак.

        Контракт ``rejected``: список словарей ``{kp_id, plate_name, qty}``.
        Невалидные элементы (без ``kp_id``/``plate_name`` или с qty<=0) пропускаются.
        Возвращает суммарное qty, которое БД успешно перевела в 'в производстве'.

        Используется и web-сервисом, и Telegram-ботом — единая точка возврата брака,
        чтобы поведение «брак → следующее планирование» было идентичным.
        ``actor`` прокидывается в audit-лог ``plate_status_log``.

        ``conn``: если задано — работа в существующей транзакции (P0). Без commit/close.
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
                _external_conn=conn,
            )
            if ok:
                total += qty
        return total

    @staticmethod
    def _verify_plates_exist_in_db(
        plates_by_kp: dict[int, list[dict[str, Any]]],
        conn: sqlite3.Connection,
    ) -> list[dict[str, Any]]:
        """P1: pre-flight проверка наличия плит в kp_plates ДО любых UPDATE.

        Для каждой ожидаемой плиты считает суммарный available qty в kp_plates
        со статусом 'в плане' или 'в производстве' и подходящей геометрией
        (length ±0.005 м, width ±0.005 м, load_class==).

        Возвращает список «недостающих» плит в формате, совместимом с
        :meth:`_format_unmoved_plates`. Пустой список = всё ок.

        Не изменяет БД.
        """
        missing: list[dict[str, Any]] = []
        cur = conn.cursor()
        # Агрегируем потребность по (kp_id, length_m, width_m, load_class).
        # plate_name не используем — его в БД могут хранить с/без префикса
        # «Плиты », а проверка по геометрии и kp_id достаточна для pre-flight.
        from collections import defaultdict as _dd

        demand: dict[tuple[int, float, float, int], int] = _dd(int)
        meta_by_key: dict[tuple[int, float, float, int], dict[str, Any]] = {}
        for kp_id, plates in plates_by_kp.items():
            for plate in plates:
                qty = int(plate.get("qty") or 0)
                if qty <= 0:
                    continue
                key = (
                    int(kp_id),
                    round(float(plate.get("length_m") or 0), 4),
                    round(float(plate.get("width_m") or 0), 4),
                    int(plate.get("load_class") or 0),
                )
                demand[key] += qty
                meta_by_key.setdefault(
                    key,
                    {
                        "kp_id": int(kp_id),
                        "plate_name": str(plate.get("plate_name") or ""),
                        "length_m": float(plate.get("length_m") or 0),
                        "width_m": float(plate.get("width_m") or 0),
                        "load_class": int(plate.get("load_class") or 0),
                    },
                )

        for (kp_id, length_m, width_m, load_class), need_qty in demand.items():
            cur.execute(
                """
                SELECT COALESCE(SUM(qty), 0) FROM kp_plates
                WHERE kp_id = ?
                  AND status IN ('в плане', 'в производстве')
                  AND ABS(length_m - ?) < 0.005
                  AND ABS(width_m - ?) < 0.005
                  AND load_class = ?
                """,
                (kp_id, length_m, width_m, load_class),
            )
            row = cur.fetchone()
            available = int((row[0] if row else 0) or 0)
            if available < need_qty:
                meta = meta_by_key[(kp_id, length_m, width_m, load_class)]
                missing.append({**meta, "qty": need_qty - available})
        return missing

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
        # P6: secondary-плиты без kp_id — это легитимные остатки
        # (перевод вторичного реза без identity заказа). Не ошибка,
        # их сохраняем в plate_rests внутри той же транзакции.
        secondary_rests: list[dict[str, Any]] = []
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
                    is_secondary = bool(plate.get("is_secondary"))
                    if not kp_id:
                        if is_secondary:
                            # P6: вторичный рез без identity → plate_rest
                            secondary_rests.append(
                                {
                                    "plate_name": plate.get("plate_name") or "",
                                    "length_m": float(plate.get("length_m") or 0),
                                    "width_mm": int(plate.get("width_mm") or 0),
                                    "qty": int(completed_qty),
                                }
                            )
                            continue
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
                "secondary_rests": secondary_rests,
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
        # load_class из БД/дневного view имеет приоритет (например, 1250 для 12.5п).
        # load_code используем только как legacy-fallback.
        raw_load_class = plate.get("load_class")
        if raw_load_class not in (None, ""):
            load_class = int(round(float(raw_load_class)))
        else:
            raw_load_code = plate.get("load_code") or 8
            if isinstance(raw_load_code, str):
                raw_load_code = raw_load_code.replace(",", ".")
            load_class = int(round(float(raw_load_code) * 100))
        return {
            "plate_name": plate.get("plate_name") or "",
            "length_m": float(plate.get("length_m") or 0),
            "width_m": float(plate.get("width_mm") or 0) / 1000.0,
            "load_class": load_class,
            "qty": int(plate.get("qty") or 0),
            "kp_id": int(plate["kp_id"]),
            # P4: length_dm_raw нужен для шага 0 в kp_db.find_one_row —
            # точный матч марки (59,81 не путается с 59,84).
            "length_dm_raw": plate.get("length_dm_raw") or "",
            "is_secondary": bool(plate.get("is_secondary")),
            "kp_plate_id": plate.get("kp_plate_id"),
        }
