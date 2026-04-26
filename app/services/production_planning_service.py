"""Web-ориентированный сервис планирования производства.

Повторяет серверную часть бота `bot.handlers.production_execution.load_and_plan_production`,
но без Telegram-зависимостей: принимает параметры, возвращает сохранённый план.
Используется веб-эндпоинтом ``POST /api/v1/production/plans/build``.
"""
from __future__ import annotations

import copy
import logging
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import core.config_and_data as cfg
from app.core.settings import get_settings
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from core.plan_commit import PlanCommitError, commit_plan_plates
from core.serialization import strip_plate_audit_from_plan
from bot.handlers import plan_manager
from core import kp_db
from core.reinforcement_db import get_reinforcement

logger = logging.getLogger(__name__)

FilterMethod = Literal["all", "kp"]


class ProductionPlanBuildError(RuntimeError):
    """Доменная ошибка сборки плана (валидное но нестроящееся состояние)."""


class ProductionPlanningService:
    """Чистое ядро планирования: SQL → оптимизация → дорожки → план."""

    def __init__(
        self,
        *,
        plita_db_path: str | None = None,
        pb_db_path: str | None = None,
    ) -> None:
        settings = get_settings()
        self.plita_db_path = str(plita_db_path or settings.plita_db_path)
        self.pb_db_path = str(pb_db_path or settings.pb_db_path)
        self.optimization_service = OptimizationService()

    def build_plan(
        self,
        *,
        start_date: str,
        tracks_count: int,
        filter_method: FilterMethod,
        selected_kp_ids: list[int] | None = None,
        selected_plate_ids: dict[int, list[int]] | None = None,
        active_plan_id: str | None = None,
        plan_name: str | None = None,
    ) -> dict[str, Any]:
        """Собирает план по заданным фильтрам.

        Returns:
            dict: ``{"plan": plan_dict, "stats": stats_dict, "summary": {...}}``
        """
        self._validate_inputs(
            start_date=start_date,
            tracks_count=tracks_count,
            filter_method=filter_method,
            selected_kp_ids=selected_kp_ids,
        )

        kp_db.init_schema(self.plita_db_path)

        kp_list = self._load_kp_list(
            filter_method=filter_method,
            selected_kp_ids=selected_kp_ids,
        )
        if not kp_list:
            raise ProductionPlanBuildError("Нет подходящих КП для производства.")

        kp_list.sort(key=lambda x: x["date"])

        plates_by_date_reinforcement = self._load_plates_for_kps(
            kp_list=kp_list,
            selected_plate_ids=selected_plate_ids or {},
        )

        selected_plates = self._flatten_plates(plates_by_date_reinforcement)
        if not selected_plates:
            raise ProductionPlanBuildError("Не найдено плит для планирования.")

        orders_2d, plate_lookup_exact, plate_lookup_by_length = self._build_orders(
            selected_plates
        )

        all_tracks_list, optimization_result = self._run_optimization_and_split(
            orders_2d=orders_2d
        )
        if not all_tracks_list:
            raise ProductionPlanBuildError("Оптимизация не дала результата.")

        # Учитываем дорожки уже занятые другими планами: если active_plan_id
        # передан — исключаем его собственные дорожки, чтобы не считать дважды.
        global_occupancy = plan_manager.get_global_day_occupancy(
            exclude_plan_id=active_plan_id,
        )

        plan, stats = plan_manager.add_tracks_to_plan(
            plan_id=active_plan_id,
            new_tracks_list=all_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            orders_2d=orders_2d,
            optimization_result=optimization_result,
            plan_name=plan_name,
            auto_save=False,
            global_occupancy=global_occupancy,
        )
        plan_id = plan["id"]

        # Коммитим план в БД до записи на диск: если пометка провалится,
        # файл плана остаётся нетронутым и шаг рестартуем без ручного cleanup.
        try:
            commit_plan_plates(
                plan_id=plan_id,
                orders_2d=orders_2d,
                optimization_result=optimization_result,
                all_tracks_list=all_tracks_list,
                db_path=self.plita_db_path,
            )
        except PlanCommitError as exc:
            logger.error(
                "[WEB-PLAN] Не удалось закоммитить план %s: %s", plan_id, exc
            )
            raise ProductionPlanBuildError(str(exc)) from exc

        try:
            plan_manager.save_plan(plan)
            plan_manager.update_plan_metadata(plan)
            plan_manager.set_active_plan(plan_id)
        except Exception as exc:
            # Плиты уже помечены — откатываем, иначе получим «залипший» статус
            # 'в плане' без соответствующего файла плана.
            logger.exception(
                "[WEB-PLAN] Ошибка записи плана %s на диск: %s", plan_id, exc
            )
            try:
                kp_db.return_plan_plates_to_production(plan_id, self.plita_db_path)
            except Exception:
                logger.exception(
                    "[WEB-PLAN] Не удалось откатить плиты для плана %s", plan_id
                )
            raise ProductionPlanBuildError(
                "Не удалось сохранить план на диск."
            ) from exc

        safe_plan = strip_plate_audit_from_plan(plan)

        summary = {
            "total_tracks": len(all_tracks_list),
            "total_days": len(safe_plan.get("days", {})),
            "selected_plates_count": sum(int(p.get("qty", 0) or 0) for p in selected_plates),
            "kp_count": len(kp_list),
        }

        logger.info(
            "[WEB-PLAN] Построен план %s: треков=%s, дней=%s, КП=%s",
            safe_plan.get("id"),
            summary["total_tracks"],
            summary["total_days"],
            summary["kp_count"],
        )

        return {"plan": safe_plan, "stats": stats, "summary": summary}

    # ----- helpers -----

    def _validate_inputs(
        self,
        *,
        start_date: str,
        tracks_count: int,
        filter_method: FilterMethod,
        selected_kp_ids: list[int] | None,
    ) -> None:
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
        except ValueError as exc:
            raise ProductionPlanBuildError(
                f"Неверный формат start_date (ожидается YYYY-MM-DD): {start_date}"
            ) from exc
        if not (1 <= tracks_count <= 50):
            raise ProductionPlanBuildError(
                "tracks_count должен быть от 1 до 50."
            )
        if filter_method == "kp" and not selected_kp_ids:
            raise ProductionPlanBuildError(
                "Для filter_method='kp' нужно передать selected_kp_ids."
            )

    def _load_kp_list(
        self,
        *,
        filter_method: FilterMethod,
        selected_kp_ids: list[int] | None,
    ) -> list[dict[str, Any]]:
        with sqlite3.connect(self.plita_db_path) as conn:
            cur = conn.cursor()

            if filter_method == "kp":
                assert selected_kp_ids is not None
                placeholders = ",".join("?" * len(selected_kp_ids))
                cur.execute(
                    f"""
                    SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                    FROM KP_offers kp
                    JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                    WHERE kp.kp_id IN ({placeholders})
                      AND meta.status = 'в работе'
                    """,
                    tuple(selected_kp_ids),
                )
            else:  # "all"
                cur.execute(
                    """
                    SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                    FROM KP_offers kp
                    JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                    WHERE meta.status = 'в работе'
                    """
                )

            rows = cur.fetchall()

        result: list[dict[str, Any]] = []
        for kp_id, exec_terms, customer in rows:
            exec_date = self._parse_exec_date(exec_terms)
            result.append(
                {
                    "kp_id": kp_id,
                    "date": exec_date,
                    "customer": customer or "",
                }
            )
        return result

    @staticmethod
    def _parse_exec_date(raw: str | None) -> datetime:
        if not raw:
            return datetime.now()
        try:
            return datetime.strptime(raw, "%d.%m.%Y")
        except ValueError:
            return datetime.now()

    def _load_plates_for_kps(
        self,
        *,
        kp_list: list[dict[str, Any]],
        selected_plate_ids: dict[int, list[int]],
    ) -> dict[datetime, dict[float, list[dict[str, Any]]]]:
        plates: dict[datetime, dict[float, list[dict[str, Any]]]] = defaultdict(
            lambda: defaultdict(list)
        )

        with sqlite3.connect(self.plita_db_path) as conn:
            cur = conn.cursor()
            for kp in kp_list:
                kp_id = kp["kp_id"]
                plate_ids = selected_plate_ids.get(kp_id) or selected_plate_ids.get(str(kp_id))
                if plate_ids is not None and len(plate_ids) == 0:
                    continue
                if plate_ids:
                    placeholders = ",".join("?" * len(plate_ids))
                    cur.execute(
                        f"""
                        SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw
                        FROM kp_plates
                        WHERE kp_id = ? AND status = 'в производстве'
                          AND id IN ({placeholders})
                        ORDER BY position_number, id
                        """,
                        (kp_id,) + tuple(plate_ids),
                    )
                else:
                    cur.execute(
                        """
                        SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw
                        FROM kp_plates
                        WHERE kp_id = ? AND status = 'в производстве'
                        """,
                        (kp_id,),
                    )

                for row in cur.fetchall():
                    plate_name = row[0]
                    length_m = row[1]
                    width_m = row[2]
                    load_class = row[3] or 800
                    qty = int(row[4] or 0)
                    length_dm_raw = (row[5] or "") if len(row) > 5 else ""
                    if qty <= 0:
                        continue

                    load_code = cfg.normalize_load_code(load_class // 100)
                    width_mm = int(round((width_m or 0) * 1000))
                    if (
                        plate_name
                        and ("-12-8п" in plate_name or "-12-" in plate_name)
                        and (width_m is None or width_m < 0.9 or width_m > 1.5)
                    ):
                        width_mm = 1200

                    reinforcement = get_reinforcement(
                        length_m=length_m,
                        load_code=load_code,
                        source="series",
                        db_path=self.pb_db_path,
                        allow_fallback=True,
                    )
                    if reinforcement is None:
                        reinforcement = 999.0

                    plates[kp["date"]][reinforcement].append(
                        {
                            "plate_name": plate_name,
                            "length": length_m,
                            "width": width_mm,
                            "load_code": load_code,
                            "qty": qty,
                            "reinforcement": reinforcement,
                            "kp_id": kp_id,
                            "kp_date": kp["date"].strftime("%d.%m.%Y"),
                            "customer": kp["customer"],
                            "length_dm_raw": length_dm_raw,
                        }
                    )

        return plates

    @staticmethod
    def _flatten_plates(
        plates_by_date_reinforcement: dict[datetime, dict[float, list[dict[str, Any]]]]
    ) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for kp_date in sorted(plates_by_date_reinforcement.keys()):
            reinforcements = plates_by_date_reinforcement[kp_date]
            for reinforcement in sorted(reinforcements.keys()):
                result.extend(reinforcements[reinforcement])
        return result

    @staticmethod
    def _build_orders(
        selected_plates: list[dict[str, Any]]
    ) -> tuple[list[dict[str, Any]], dict, dict]:
        orders_2d: list[dict[str, Any]] = []
        for plate in selected_plates:
            orders_2d.append(
                {
                    "length": plate["length"],
                    "width": plate["width"],
                    "qty": plate["qty"],
                    "load_code": cfg.normalize_load_code(plate["load_code"]),
                    "reinforcement": plate["reinforcement"],
                    "kp_date": plate.get("kp_date", "неизвестно"),
                    "customer": plate.get("customer", "неизвестно"),
                    "plate_name": plate.get("plate_name", ""),
                    "kp_id": plate.get("kp_id"),
                    "length_dm_raw": plate.get("length_dm_raw", "") or "",
                }
            )

        plate_lookup_exact: dict[tuple[float, int], list[dict[str, Any]]] = {}
        plate_lookup_by_length: dict[float, list[dict[str, Any]]] = {}

        def parse_date_for_sort(date_str: str) -> datetime:
            if not date_str or date_str == "неизвестно":
                return datetime.max
            try:
                return datetime.strptime(date_str, "%d.%m.%Y")
            except ValueError:
                return datetime.max

        for order in orders_2d:
            length_key = round(order["length"], 2)
            width_key = order["width"]
            entry_exact = {
                "kp_date": order.get("kp_date", "неизвестно"),
                "customer": order.get("customer", "неизвестно"),
                "plate_name": order.get("plate_name", ""),
                "reinforcement": order.get("reinforcement", 0),
                "load_code": cfg.normalize_load_code(order.get("load_code", 8)),
                "qty_remaining": order.get("qty", 1),
                "kp_id": order.get("kp_id"),
            }
            plate_lookup_exact.setdefault((length_key, width_key), []).append(entry_exact)

            entry_by_length = {
                "kp_date": order.get("kp_date", "неизвестно"),
                "customer": order.get("customer", "неизвестно"),
                "plate_name": order.get("plate_name", ""),
                "reinforcement": order.get("reinforcement", 0),
                "load_code": cfg.normalize_load_code(order.get("load_code", 8)),
                "qty_remaining": order.get("qty", 1),
                "kp_id": order.get("kp_id"),
            }
            plate_lookup_by_length.setdefault(length_key, []).append(entry_by_length)

        for key in plate_lookup_exact:
            plate_lookup_exact[key].sort(
                key=lambda x: parse_date_for_sort(x.get("kp_date", ""))
            )
        for key in plate_lookup_by_length:
            plate_lookup_by_length[key].sort(
                key=lambda x: parse_date_for_sort(x.get("kp_date", ""))
            )

        return orders_2d, plate_lookup_exact, plate_lookup_by_length

    def _run_optimization_and_split(
        self, *, orders_2d: list[dict[str, Any]]
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        if not orders_2d:
            return [], {}

        plate_order = AppPlateOrder.from_orders_2d(orders_2d)

        # build_layout_sequence читает cfg.PLATE_LOAD_DETAILS; делаем временный snapshot
        saved_plate_load_details = copy.deepcopy(cfg.PLATE_LOAD_DETAILS)
        saved_plate_length_raw = copy.deepcopy(cfg.PLATE_LENGTH_DM_RAW)
        try:
            cfg.PLATE_LOAD_DETAILS.clear()
            for key, qty in plate_order.plate_load_details.items():
                length, width_m, load_code, raw = key
                normalized_key = (length, width_m, int(float(load_code)), raw)
                cfg.PLATE_LOAD_DETAILS[normalized_key] = int(qty)
            cfg.PLATE_LENGTH_DM_RAW.clear()
            for key, raw in plate_order.plate_length_dm_raw.items():
                length, width_m, load_code, raw_val = key
                cfg.PLATE_LENGTH_DM_RAW[(length, width_m, int(float(load_code)), raw_val)] = raw

            context = self.optimization_service.optimize(
                plate_order,
                orders_2d=orders_2d,
            )
            optimization_result = context.optimization_result or {}
            self._log_unmapped_optimizer_assignments(optimization_result)
            if not optimization_result or optimization_result.get("total_plates", 0) == 0:
                return [], optimization_result

            from core.visualization import split_sequence_into_tracks
            from viz_modules.layout_sequence import build_layout_sequence

            with self.optimization_service.legacy_runtime(context):
                seq = build_layout_sequence()
                all_tracks_list = split_sequence_into_tracks(seq) or []
        finally:
            cfg.PLATE_LOAD_DETAILS.clear()
            cfg.PLATE_LOAD_DETAILS.update(saved_plate_load_details)
            cfg.PLATE_LENGTH_DM_RAW.clear()
            cfg.PLATE_LENGTH_DM_RAW.update(saved_plate_length_raw)

        return all_tracks_list, optimization_result

    @staticmethod
    def _log_unmapped_optimizer_assignments(optimization_result: dict[str, Any]) -> None:
        unmapped = []
        for assignment in optimization_result.get("plate_assignments", []) or []:
            source = str(assignment.get("source") or "unknown")
            if source not in {"primary", "secondary"}:
                continue
            if assignment.get("kp_id") and assignment.get("plate_name"):
                continue
            unmapped.append(
                {
                    "source": source,
                    "length": assignment.get("length"),
                    "width": assignment.get("width"),
                    "load_code": assignment.get("load_code"),
                    "identity_match_type": assignment.get("identity_match_type"),
                    "has_kp_id": bool(assignment.get("kp_id")),
                    "has_plate_name": bool(assignment.get("plate_name")),
                }
            )

        if unmapped:
            logger.error(
                "[WEB-PLAN] Optimizer assignments без exact identity: count=%s sample=%s",
                len(unmapped),
                unmapped[:10],
            )
