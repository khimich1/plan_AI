"""Web-ориентированный сервис планирования производства.

Повторяет серверную часть бота `bot.handlers.production_execution.load_and_plan_production`,
но без Telegram-зависимостей: принимает параметры, возвращает сохранённый план.
Используется веб-эндпоинтом ``POST /api/v1/production/plans/build``.
"""
from __future__ import annotations

import copy
import logging
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import core.config_and_data as cfg
from app.core.settings import get_settings
from app.domain.enums import PlateStatus
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from core.plan_commit import PlanCommitError, commit_plan_plates
from core.serialization import strip_plate_audit_from_plan
from core.debug_paths import get_debug_log_path
from bot.handlers import plan_manager
from core import kp_db
from core.reinforcement_db import get_reinforcement

logger = logging.getLogger(__name__)
_DEBUG_AGENT_LOG = get_debug_log_path("debug-ebb546.log")

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
        fill_targets: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Собирает план по заданным фильтрам.

        Если передан ``fill_targets`` — это режим «дозаполнения дней»: дорожки
        раскладываются строго по выбранным датам, лишние плиты обрезаются
        и НЕ помечаются как ``в плане`` (остаются ``в производстве``).

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

        # #region agent log
        try:
            import json as _agent_json
            import time as _agent_time
            from collections import Counter as _AgentCounter

            _selected_counter = _AgentCounter(
                f"{p.get('kp_id')}|{p.get('plate_name')}"
                for p in selected_plates
            )
            _orders_counter = _AgentCounter(
                f"{o.get('kp_id')}|{o.get('plate_name')}"
                for o in orders_2d
                for _ in range(int(o.get("qty") or 0))
            )
            with open(_DEBUG_AGENT_LOG, "a", encoding="utf-8") as _agent_f:
                _agent_f.write(_agent_json.dumps({
                    "sessionId": "ebb546",
                    "runId": "stage-localization",
                    "hypothesisId": "S3,S4",
                    "location": "app/services/production_planning_service.py:after_build_orders",
                    "message": "Стадии 3-4: загрузка плит и orders_2d перед оптимизатором",
                    "data": {
                        "kp_ids": [kp.get("kp_id") for kp in kp_list],
                        "selected_plates_rows": len(selected_plates),
                        "selected_plates_qty": sum(int(p.get("qty") or 0) for p in selected_plates),
                        "orders_rows": len(orders_2d),
                        "orders_qty": sum(int(o.get("qty") or 0) for o in orders_2d),
                        "selected_top": _selected_counter.most_common(12),
                        "orders_top": _orders_counter.most_common(12),
                    },
                    "timestamp": int(_agent_time.time() * 1000),
                }, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion

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

        precomputed_tracks_by_day: dict[str, list[dict[str, Any]]] | None = None
        if fill_targets:
            # Режим «дозаполнения»: жёсткое размещение по выбранным дням,
            # без переноса лишних дорожек на следующий рабочий день.
            self._validate_fill_targets(fill_targets, global_occupancy)

            cap = sum(int(t["tracks"]) for t in fill_targets)
            if cap < len(all_tracks_list):
                logger.info(
                    "[WEB-PLAN] fill_targets cap=%s, обрезаем дорожек: %s -> %s",
                    cap,
                    len(all_tracks_list),
                    cap,
                )
                all_tracks_list = all_tracks_list[:cap]
                optimization_result = self._trim_assignments_to_tracks(
                    optimization_result=optimization_result,
                    kept_tracks=all_tracks_list,
                )

            precomputed_tracks_by_day = self._build_tracks_by_day_from_targets(
                kept_tracks=all_tracks_list,
                fill_targets=fill_targets,
            )

            # Дозаполнение всегда новый план: дописывать в чужой план дорожки
            # в его уже планируемый день — некорректно по семантике планов.
            active_plan_id = None
            sorted_dates = sorted(t["date"] for t in fill_targets)
            start_date = sorted_dates[0]
            tracks_per_day_effective = max(int(t["tracks"]) for t in fill_targets)
        else:
            tracks_per_day_effective = tracks_count

        plan, stats = plan_manager.add_tracks_to_plan(
            plan_id=active_plan_id,
            new_tracks_list=all_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_per_day_effective,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            orders_2d=orders_2d,
            optimization_result=optimization_result,
            plan_name=plan_name,
            auto_save=False,
            global_occupancy=global_occupancy,
            precomputed_tracks_by_day=precomputed_tracks_by_day,
        )
        plan_id = plan["id"]

        # P5: собираем tracks_by_day из готового plan'а — там tracks уже
        # разложены по датам и имеют day_number в day-ноде. Прокидываем
        # это в commit_plan_plates, чтобы он записал day_number в БД и
        # kp_plate_id в каждый track.items[].
        tracks_by_day_for_commit: dict[str, list[dict[str, Any]]] = {}
        for date_key, day_data in (plan.get("days") or {}).items():
            day_number = int((day_data or {}).get("day_number") or 0)
            day_tracks = (day_data or {}).get("tracks") or []
            for track in day_tracks:
                if isinstance(track, dict):
                    track.setdefault("production_day", day_number)
            tracks_by_day_for_commit[date_key] = day_tracks

        # #region agent log
        try:
            import json as _agent_json
            import time as _agent_time
            from collections import Counter as _AgentCounter

            _physical_total = 0
            _without_identity = 0
            _secondary_without_identity = 0
            _identity_counts: _AgentCounter[str] = _AgentCounter()
            for _day_tracks in tracks_by_day_for_commit.values():
                for _track in _day_tracks or []:
                    for _item in (_track or {}).get("items") or []:
                        if not isinstance(_item, dict):
                            continue
                        _items_to_check = [(_item, False)] + [
                            (_sec, True)
                            for _sec in (_item.get("secondary_cuts") or [])
                            if isinstance(_sec, dict)
                        ]
                        for _physical, _is_secondary in _items_to_check:
                            _physical_total += 1
                            _kp_id = _physical.get("kp_id")
                            _plate_name = _physical.get("plate_name") or ""
                            if not (_kp_id and _plate_name):
                                _without_identity += 1
                                if _is_secondary:
                                    _secondary_without_identity += 1
                            else:
                                _identity_counts[f"{_kp_id}|{_plate_name}"] += 1
            with open(_DEBUG_AGENT_LOG, "a", encoding="utf-8") as _agent_f:
                _agent_f.write(_agent_json.dumps({
                    "sessionId": "ebb546",
                    "runId": "pre-fix",
                    "hypothesisId": "H1,H2",
                    "location": "app/services/production_planning_service.py:before_commit",
                    "message": "План перед commit_plan_plates: tracks_by_day и identity в физических items",
                    "data": {
                        "plan_id": plan_id,
                        "dates": sorted(tracks_by_day_for_commit.keys()),
                        "tracks_by_date": {k: len(v or []) for k, v in tracks_by_day_for_commit.items()},
                        "physical_items_total": _physical_total,
                        "physical_items_without_identity": _without_identity,
                        "secondary_without_identity": _secondary_without_identity,
                        "top_identity_counts": _identity_counts.most_common(12),
                        "orders_total_qty": sum(int(o.get("qty") or 0) for o in orders_2d),
                    },
                    "timestamp": int(_agent_time.time() * 1000),
                }, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion

        # Коммитим план в БД до записи на диск: если пометка провалится,
        # файл плана остаётся нетронутым и шаг рестартуем без ручного cleanup.
        try:
            commit_plan_plates(
                plan_id=plan_id,
                orders_2d=orders_2d,
                optimization_result=optimization_result,
                all_tracks_list=all_tracks_list,
                db_path=self.plita_db_path,
                tracks_by_day=tracks_by_day_for_commit,
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
                        WHERE kp_id = ? AND status = ?
                          AND id IN ({placeholders})
                        ORDER BY position_number, id
                        """,
                        (kp_id, PlateStatus.IN_PRODUCTION.value) + tuple(plate_ids),
                    )
                else:
                    cur.execute(
                        """
                        SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw
                        FROM kp_plates
                        WHERE kp_id = ? AND status = ?
                        """,
                        (kp_id, PlateStatus.IN_PRODUCTION.value),
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
            # P4: length_dm_raw нужен в lookup-таблицах, чтобы day_view → _to_completed_plate_payload
            # мог пробросить его в find_one_row (шаг 0 — самый точный матч 59,8 vs 59,9).
            entry_exact = {
                "kp_date": order.get("kp_date", "неизвестно"),
                "customer": order.get("customer", "неизвестно"),
                "plate_name": order.get("plate_name", ""),
                "reinforcement": order.get("reinforcement", 0),
                "load_code": cfg.normalize_load_code(order.get("load_code", 8)),
                "qty_remaining": order.get("qty", 1),
                "kp_id": order.get("kp_id"),
                "length_dm_raw": order.get("length_dm_raw", "") or "",
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
                "length_dm_raw": order.get("length_dm_raw", "") or "",
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

        from core.plate_attribution import (
            backfill_assignment_identity,
            backfill_track_items_identity,
        )

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

            # P8.1: backfill identity у plate_assignments, чтобы slot_exhausted
            # / secondary_unmapped перестали быть блокером в plan_commit.
            backfilled = backfill_assignment_identity(
                optimization_result.get("plate_assignments", []) or [],
                orders_2d,
            )
            if backfilled:
                logger.info(
                    "[WEB-PLAN] Восстановлена identity у %s plate_assignments-записей",
                    backfilled,
                )

            self._log_unmapped_optimizer_assignments(optimization_result)
            if not optimization_result or optimization_result.get("total_plates", 0) == 0:
                return [], optimization_result

            from core.visualization import LayoutIntegrityError, split_sequence_into_tracks
            from viz_modules.layout_sequence import build_layout_sequence

            with self.optimization_service.legacy_runtime(context):
                seq = build_layout_sequence()
                try:
                    all_tracks_list = split_sequence_into_tracks(
                        seq,
                        strict_layout_integrity=True,
                    ) or []
                except LayoutIntegrityError as exc:
                    raise ProductionPlanBuildError(
                        f"Нарушена целостность раскладки дорожек: {exc}"
                    ) from exc

            # #region agent log
            try:
                import json as _agent_json
                import time as _agent_time
                from collections import Counter as _AgentCounter

                def _count_sequence_items(_seq: list[dict[str, Any]]) -> tuple[int, int, _AgentCounter[str]]:
                    _total = 0
                    _without_identity = 0
                    _counts: _AgentCounter[str] = _AgentCounter()
                    if (
                        isinstance(_seq, list)
                        and _seq
                        and isinstance(_seq[0], dict)
                        and isinstance(_seq[0].get("sequence"), list)
                    ):
                        _iter_items: list[dict[str, Any]] = []
                        for _group in _seq:
                            if not isinstance(_group, dict):
                                continue
                            for _item in _group.get("sequence") or []:
                                if isinstance(_item, dict):
                                    _iter_items.append(_item)
                    else:
                        _iter_items = [
                            _item for _item in (_seq or [])
                            if isinstance(_item, dict)
                        ]
                    for _item in _iter_items:
                        _total += 1
                        _kp_id = _item.get("kp_id")
                        _plate_name = _item.get("plate_name") or _item.get("label")
                        if _kp_id and _plate_name:
                            _counts[f"{_kp_id}|{_plate_name}"] += 1
                        else:
                            _without_identity += 1
                    return _total, _without_identity, _counts

                def _count_track_items(_tracks: list[dict[str, Any]]) -> tuple[int, int, _AgentCounter[str]]:
                    _total = 0
                    _without_identity = 0
                    _counts: _AgentCounter[str] = _AgentCounter()
                    for _track in _tracks or []:
                        if not isinstance(_track, dict):
                            continue
                        for _item in _track.get("items") or []:
                            if not isinstance(_item, dict):
                                continue
                            _physical = [_item] + [
                                _sec for _sec in (_item.get("secondary_cuts") or [])
                                if isinstance(_sec, dict)
                            ]
                            for _p in _physical:
                                _total += 1
                                _kp_id = _p.get("kp_id")
                                _plate_name = _p.get("plate_name") or _p.get("label")
                                if _kp_id and _plate_name:
                                    _counts[f"{_kp_id}|{_plate_name}"] += 1
                                else:
                                    _without_identity += 1
                    return _total, _without_identity, _counts

                _seq_total, _seq_without_identity, _seq_counts = _count_sequence_items(seq)
                _tracks_total, _tracks_without_identity, _tracks_counts = _count_track_items(all_tracks_list)
                with open(_DEBUG_AGENT_LOG, "a", encoding="utf-8") as _agent_f:
                    _agent_f.write(_agent_json.dumps({
                        "sessionId": "ebb546",
                        "runId": "stage-localization",
                        "hypothesisId": "S5B,S5C",
                        "location": "app/services/production_planning_service.py:after_layout_split",
                        "message": "Стадия 5B-5C: build_layout_sequence и split_sequence_into_tracks",
                        "data": {
                            "orders_qty": sum(int(o.get("qty") or 0) for o in orders_2d),
                            "sequence_items_total": _seq_total,
                            "sequence_without_identity": _seq_without_identity,
                            "sequence_top": _seq_counts.most_common(15),
                            "tracks_count": len(all_tracks_list),
                            "track_physical_items_total": _tracks_total,
                            "track_physical_without_identity": _tracks_without_identity,
                            "track_top": _tracks_counts.most_common(15),
                        },
                        "timestamp": int(_agent_time.time() * 1000),
                    }, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
        finally:
            cfg.PLATE_LOAD_DETAILS.clear()
            cfg.PLATE_LOAD_DETAILS.update(saved_plate_load_details)
            cfg.PLATE_LENGTH_DM_RAW.clear()
            cfg.PLATE_LENGTH_DM_RAW.update(saved_plate_length_raw)

        # P7/P8: подключаем общую с ботом RESCUE-логику. Без неё web-side
        # планирование молча теряло позиции, которых не хватало.
        # P8.4: после Phase 4 источник правды — plate_assignments
        # (с backfill identity), а не all_tracks_list. Fuzzy-матч ушёл,
        # подсчёт делается по точной identity (kp_id, plate_name).
        try:
            from core.rescue_tracks import build_rescue_tracks

            plate_assignments = optimization_result.get("plate_assignments", []) or []
            rescue_tracks, missing_counts, rescue_assignments = build_rescue_tracks(
                orders_2d=orders_2d,
                plate_assignments=plate_assignments,
            )
            if rescue_tracks:
                logger.info(
                    "[WEB-PLAN] RESCUE: добавлено %s дорожек / %s assignments, "
                    "missing identities=%s",
                    len(rescue_tracks),
                    len(rescue_assignments),
                    len(missing_counts),
                )
                all_tracks_list.extend(rescue_tracks)
                if rescue_assignments:
                    optimization_result.setdefault("plate_assignments", []).extend(
                        rescue_assignments
                    )
        except Exception:
            logger.exception("[WEB-PLAN] RESCUE-логика упала — продолжаем без rescue")

        # P9: backfill identity у track-items / secondary_cuts. Зеркало
        # backfill_assignment_identity, но для дерева track["items"][*]
        # /items[*]["secondary_cuts"][*]. Без этого secondary без kp_id
        # не учитываются в _count_track_items_by_day -> плиты помечаются
        # без day_number, не отображаются в day_view, не списываются.
        # ВАЖНО: должен пройти ДО _trim_assignments_to_tracks (в режиме
        # fill_targets), потому что trim считает kept-плиты ПО identity
        # items: если identity нет — assignments выкидываются как лишние.
        backfilled_items = backfill_track_items_identity(
            all_tracks_list,
            orders_2d,
        )
        if backfilled_items:
            logger.info(
                "[WEB-PLAN] Восстановлена identity у %s track items "
                "(root + secondary_cuts)",
                backfilled_items,
            )

        try:
            fallback_tracks, fallback_missing = self._build_assignment_gap_fallback_tracks(
                plate_assignments=optimization_result.get("plate_assignments", []) or [],
                tracks_list=all_tracks_list,
            )
            if fallback_tracks:
                logger.warning(
                    "[WEB-PLAN] FALLBACK track-gap: добавлено %s дорожек для "
                    "%s плит из assignments, отсутствующих в tracks: %s",
                    len(fallback_tracks),
                    sum(fallback_missing.values()),
                    fallback_missing,
                )
                all_tracks_list.extend(fallback_tracks)
        except Exception:
            logger.exception(
                "[WEB-PLAN] FALLBACK track-gap логика упала — продолжаем без неё"
            )

        return all_tracks_list, optimization_result

    @staticmethod
    def _build_assignment_gap_fallback_tracks(
        *,
        plate_assignments: list[dict[str, Any]],
        tracks_list: list[dict[str, Any]],
    ) -> tuple[list[dict[str, Any]], Counter]:
        """Возвращает fallback-треки для плит, потерянных между assignments и tracks."""
        assignments_by_unit_id: dict[str, dict[str, Any]] = {}
        missing_counter: Counter = Counter()
        for assignment in plate_assignments or []:
            unit_id = assignment.get("unit_id")
            if not unit_id:
                continue
            assignments_by_unit_id[str(unit_id)] = assignment

        present_unit_ids: set[str] = set()
        for track in tracks_list or []:
            if not isinstance(track, dict):
                continue
            for item in track.get("items") or []:
                if not isinstance(item, dict):
                    continue
                root_id = item.get("unit_id")
                if root_id:
                    present_unit_ids.add(str(root_id))
                for sec in item.get("secondary_cuts") or []:
                    if isinstance(sec, dict) and sec.get("unit_id"):
                        present_unit_ids.add(str(sec.get("unit_id")))

        missing_assignments = [
            a for uid, a in assignments_by_unit_id.items()
            if uid not in present_unit_ids
        ]
        fallback_tracks: list[dict[str, Any]] = []
        for assignment in missing_assignments:
            source = str(assignment.get("source") or "")
            width_mm = int(round(float(assignment.get("width") or 1200)))
            length = float(assignment.get("length") or 6.0)
            mode = "split" if source == "secondary" else "solid"
            item: dict[str, Any] = {
                "mode": mode,
                "length": length,
                "width": width_mm / 1000.0,
                "unit_id": assignment.get("unit_id"),
                "layout_uid": str(assignment.get("unit_id")),
                "parent_unit_id": assignment.get("parent_unit_id"),
                "kp_id": assignment.get("kp_id"),
                "customer": assignment.get("customer"),
                "kp_date": assignment.get("kp_date"),
                "plate_name": assignment.get("plate_name"),
                "load_code": assignment.get("load_code", 8),
                "placement_status": "fallback",
                "origin_reason": "track_gap_missing_unit_id",
            }
            if mode == "solid":
                item["label"] = assignment.get("plate_name") or f"fallback:{width_mm}"
            else:
                main_w = width_mm / 1000.0
                item["main_w"] = main_w
                item["rest_w"] = 0.0
                item["label_main"] = assignment.get("plate_name") or f"fallback:{width_mm}"
                item["label_rest"] = None
                item["secondary_cuts"] = []
            track_label = "FALLBACK"
            fallback_tracks.append({
                "label": track_label,
                "placement_status": "fallback",
                "items": [item],
                "length": length,
            })
            missing_counter[(assignment.get("kp_id"), assignment.get("plate_name") or "")] += 1

        return fallback_tracks, missing_counter

    @staticmethod
    def _validate_fill_targets(
        fill_targets: list[dict[str, Any]],
        global_occupancy: dict[str, int],
    ) -> None:
        """Проверяет, что каждая дата существует и имеет достаточно свободных слотов."""
        if not fill_targets:
            raise ProductionPlanBuildError("fill_targets пуст.")
        max_per_day = plan_manager.MAX_TRACKS_PER_DAY
        for t in fill_targets:
            date = t.get("date") or ""
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError as exc:
                raise ProductionPlanBuildError(
                    f"Неверный формат даты в fill_targets: {date}"
                ) from exc
            occupied = int(global_occupancy.get(date, 0) or 0)
            free = max(0, max_per_day - occupied)
            requested = int(t.get("tracks", 0) or 0)
            if requested <= 0:
                raise ProductionPlanBuildError(
                    f"На {date} запрошено {requested} дорожек — должно быть >= 1."
                )
            if requested > free:
                raise ProductionPlanBuildError(
                    f"На {date} свободно {free} дорожек, запрошено {requested}."
                )

    @staticmethod
    def _build_tracks_by_day_from_targets(
        *,
        kept_tracks: list[dict[str, Any]],
        fill_targets: list[dict[str, Any]],
    ) -> dict[str, list[dict[str, Any]]]:
        """Раскладывает kept_tracks по датам строго в порядке fill_targets.

        Не использует ``distribute_tracks_by_days`` — нам нужно положить дорожки
        ровно в указанные даты, без переноса на следующий рабочий день.
        """
        result: dict[str, list[dict[str, Any]]] = {}
        cursor = 0
        for t in fill_targets:
            requested = int(t["tracks"])
            chunk = kept_tracks[cursor:cursor + requested]
            if chunk:
                result[t["date"]] = chunk
            cursor += requested
        return result

    @staticmethod
    def _trim_assignments_to_tracks(
        *,
        optimization_result: dict[str, Any],
        kept_tracks: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Оставляет в plate_assignments столько плит, сколько лежит в kept_tracks.

        ``commit_plan_plates`` помечает плиты «в плане» по
        ``optimization_result["plate_assignments"]`` (для primary/secondary)
        и по track items с ``label='РЕСКЬЮ'`` (для rescue). Если просто срезать
        ``all_tracks_list``, но оставить ``plate_assignments`` нетронутым — будут
        помечены и плиты выкинутых дорожек, либо ``commit_plan_plates`` упадёт
        на ``leftovers_by_source``. Поэтому считаем плиты по идентичности
        ``(kp_id, plate_name)`` в kept_tracks (без RESCUE — он считается из
        track items напрямую) и оставляем в assignments не больше этого числа.
        Порядок исходного списка сохраняется, поэтому primary-плиты той же пары
        идентичности приоритетнее secondary, как и было до обрезания.
        """
        kept_counts: Counter[tuple[Any, str]] = Counter()
        for track in kept_tracks:
            if track.get("label") == "РЕСКЬЮ":
                continue
            for item in track.get("items") or []:
                if not item:
                    continue
                kp_id = item.get("kp_id")
                plate_name = item.get("plate_name") or ""
                if kp_id and plate_name:
                    kept_counts[(kp_id, plate_name)] += 1

        remaining = Counter(kept_counts)
        filtered: list[dict[str, Any]] = []
        for assignment in optimization_result.get("plate_assignments", []) or []:
            source = assignment.get("source")
            if source not in ("primary", "secondary"):
                # rescue/прочие источники не управляются через assignments
                filtered.append(assignment)
                continue
            key = (assignment.get("kp_id"), assignment.get("plate_name") or "")
            if remaining.get(key, 0) > 0:
                filtered.append(assignment)
                remaining[key] -= 1

        return {**optimization_result, "plate_assignments": filtered}

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
