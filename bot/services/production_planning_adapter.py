"""Thin bot adapter over :mod:`core.production.planning` / :class:`ProductionPlanningService`."""

from __future__ import annotations

import logging
import sqlite3
from datetime import datetime
from typing import Any, Literal

import core.config_and_data as cfg
from app.services.production_planning_service import (
    ProductionPlanBuildError,
    ProductionPlanningService,
)
from bot.services import kp_persistence as kp_db
from core.plate_order_context import PlateOrderContext
from core.production.dto import LoadConfig, LoadResult, PlanBuildInput
from core.production.errors import PlanBuildError
from core.production.planning import _build_orders, load

logger = logging.getLogger(__name__)

BotFilterMethod = Literal["date", "kp", "all", "customer"]


def load_kp_list_for_bot_filter(
    *,
    db_path: str,
    filter_method: BotFilterMethod,
    target_date: datetime | None = None,
    kp_ids: list[int] | None = None,
    customer_name: str = "",
) -> list[dict[str, Any]]:
    """Загружает список КП по фильтру бота (date/kp/all/customer)."""
    kp_db.ensure_schema(db_path)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    kp_list: list[dict[str, Any]] = []

    try:
        if filter_method == "date":
            if target_date is None:
                return []
            cur.execute(
                """
                SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                FROM KP_offers kp
                JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                WHERE meta.status = 'в работе'
                """
            )
            for kp_id, exec_terms, customer in cur.fetchall():
                if not exec_terms:
                    continue
                try:
                    exec_date = datetime.strptime(exec_terms, "%d.%m.%Y")
                except ValueError:
                    continue
                if exec_date <= target_date:
                    kp_list.append(
                        {
                            "kp_id": kp_id,
                            "date": exec_date,
                            "customer": customer,
                        }
                    )

        elif filter_method == "kp":
            if not kp_ids:
                return []
            placeholders = ",".join("?" * len(kp_ids))
            cur.execute(
                f"""
                SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                FROM KP_offers kp
                JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                WHERE kp.kp_id IN ({placeholders})
                  AND meta.status = 'в работе'
                """,
                tuple(kp_ids),
            )
            for kp_id, exec_terms, customer in cur.fetchall():
                exec_date = (
                    datetime.strptime(exec_terms, "%d.%m.%Y")
                    if exec_terms
                    else datetime.now()
                )
                kp_list.append(
                    {
                        "kp_id": kp_id,
                        "date": exec_date,
                        "customer": customer,
                    }
                )

        elif filter_method == "all":
            cur.execute(
                """
                SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                FROM KP_offers kp
                JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                WHERE meta.status = 'в работе'
                """
            )
            for kp_id, exec_terms, customer in cur.fetchall():
                exec_date = (
                    datetime.strptime(exec_terms, "%d.%m.%Y")
                    if exec_terms
                    else datetime.now()
                )
                kp_list.append(
                    {
                        "kp_id": kp_id,
                        "date": exec_date,
                        "customer": customer,
                    }
                )

        elif filter_method == "customer":
            cur.execute(
                """
                SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                FROM KP_offers kp
                JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                WHERE meta.status = 'в работе'
                  AND kp.customer_name = ?
                """,
                (customer_name,),
            )
            for kp_id, exec_terms, customer in cur.fetchall():
                exec_date = (
                    datetime.strptime(exec_terms, "%d.%m.%Y")
                    if exec_terms
                    else datetime.now()
                )
                kp_list.append(
                    {
                        "kp_id": kp_id,
                        "date": exec_date,
                        "customer": customer,
                    }
                )
    finally:
        conn.close()

    kp_list.sort(key=lambda x: x["date"])
    return kp_list


def normalize_kp_plate_ids(
    raw: dict[Any, Any] | None,
) -> dict[str, list[int]] | None:
    if not raw:
        return None
    result: dict[str, list[int]] = {}
    for key, value in raw.items():
        sk = str(key)
        if value is not None and not isinstance(value, list):
            value = (
                list(value)
                if hasattr(value, "__iter__") and not isinstance(value, str)
                else []
            )
        result[sk] = value
    return result


def load_plates_for_production(
    service: ProductionPlanningService,
    *,
    kp_list: list[dict[str, Any]],
    kp_plate_ids: dict[str, list[int]] | None,
    start_date: str,
    tracks_count: int,
) -> LoadResult:
    """Загружает плиты через core.load (filter_method=kp)."""
    selected_plate_ids: dict[int, list[int]] | None = None
    if kp_plate_ids:
        selected_plate_ids = {}
        for kp in kp_list:
            kp_id = int(kp["kp_id"])
            ids = kp_plate_ids.get(str(kp_id))
            if ids is not None:
                selected_plate_ids[kp_id] = list(ids)

    plan_input = PlanBuildInput(
        start_date=start_date,
        tracks_count=tracks_count,
        filter_method="kp",
        selected_kp_ids=tuple(int(k["kp_id"]) for k in kp_list),
        selected_plate_ids=selected_plate_ids,
    )
    try:
        return load(
            plan_input,
            config=LoadConfig(
                plita_db_path=service.plita_db_path,
                pb_db_path=service.pb_db_path,
            ),
        )
    except PlanBuildError as exc:
        raise ProductionPlanBuildError(str(exc)) from exc


def apply_rest_matching(
    service: ProductionPlanningService,
    selected_plates: list[dict[str, Any]],
    *,
    db_path: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], bool]:
    """Подбирает остатки; возвращает (plates_from_rests, plates_for_optimizer, all_from_rests)."""
    plates_from_rests: list[dict[str, Any]] = []
    plates_for_optimizer: list[dict[str, Any]] = []

    for plate_data in selected_plates:
        length_m = plate_data["length"]
        width_mm = plate_data["width"]
        qty_needed = plate_data["qty"]

        matching_rests = service.find_matching_rests(
            length_m=length_m,
            width_mm=width_mm,
            qty_needed=qty_needed,
            db_path=db_path,
        )

        qty_from_rests = 0
        if matching_rests:
            for rest_info in matching_rests:
                qty_to_use = rest_info["qty_to_use"]
                qty_from_rests += qty_to_use
                plates_from_rests.append(
                    {
                        "plate_name": plate_data.get("plate_name", ""),
                        "length_m": length_m,
                        "width_mm": width_mm,
                        "qty": qty_to_use,
                        "kp_id": plate_data.get("kp_id"),
                        "kp_date": plate_data.get("kp_date", "неизвестно"),
                        "customer": plate_data.get("customer", "неизвестно"),
                        "load_code": cfg.normalize_load_code(
                            plate_data.get("load_code", 8)
                        ),
                        "reinforcement": plate_data.get("reinforcement", 0),
                        "rest_id": rest_info["rest_id"],
                        "rest_length": rest_info["rest_length"],
                        "rest_width_mm": rest_info["rest_width_mm"],
                        "match_type": rest_info["match_type"],
                        "cut_cost": rest_info["cut_cost"],
                        "source_plate_name": rest_info["source_plate_name"],
                        "source_kp_id": rest_info["source_kp_id"],
                        "from_rest": True,
                    }
                )

        qty_remaining = qty_needed - qty_from_rests
        if qty_remaining > 0:
            plate_for_opt = plate_data.copy()
            plate_for_opt["qty"] = qty_remaining
            plates_for_optimizer.append(plate_for_opt)

    all_from_rests = bool(selected_plates) and not plates_for_optimizer
    return plates_from_rests, plates_for_optimizer, all_from_rests


def rebuild_load_result_for_plates(
    load_result: LoadResult,
    plates_for_optimizer: list[dict[str, Any]],
) -> LoadResult:
    """Пересобирает orders/lookup после вычета остатков."""
    orders_2d, plate_lookup_exact, plate_lookup_by_length = _build_orders(
        plates_for_optimizer
    )
    load_result.orders_2d = orders_2d
    load_result.selected_plates = plates_for_optimizer
    load_result.plate_lookup_exact = plate_lookup_exact
    load_result.plate_lookup_by_length = plate_lookup_by_length
    return load_result


def enrich_lookup_for_secondary_cuts(
    plate_lookup_exact: dict[tuple[float, int], list[dict[str, Any]]],
    plate_lookup_by_length: dict[float, list[dict[str, Any]]],
    orders_2d: list[dict[str, Any]],
    optimization_result: dict[str, Any],
) -> dict[tuple[float, int], list[dict[str, Any]]]:
    """Дополняет lookup для вторичных резов (bot-specific attribution)."""
    if not optimization_result.get("secondary_cuts"):
        return plate_lookup_exact

    orders_dict: dict[tuple[float, int, int], list[dict[str, Any]]] = {}
    for order in orders_2d:
        key = (
            round(order["length"], 2),
            order["width"],
            cfg.normalize_load_code(order.get("load_code", 8)),
        )
        orders_dict.setdefault(key, []).append(order)

    for sec_cut in optimization_result["secondary_cuts"]:
        target_key = sec_cut.get("target_order_key")
        if not target_key:
            continue

        target_length, target_width, target_load_code = target_key
        original_orders = orders_dict.get(
            (round(target_length, 2), target_width, target_load_code), []
        )
        if not original_orders:
            continue

        result_lengths = sec_cut.get("lengths", [])
        result_width = sec_cut["cuts"][0]

        for result_length in result_lengths:
            key_result = (round(result_length, 2), result_width)
            original_order = None
            for order in original_orders:
                original_key = (round(order["length"], 2), order["width"])
                if original_key in plate_lookup_exact:
                    for entry in plate_lookup_exact[original_key]:
                        if entry.get("qty_remaining", 0) > 0:
                            original_order = order
                            break
                if original_order:
                    break

            if not original_order:
                continue

            plate_lookup_exact.setdefault(key_result, []).append(
                {
                    "kp_date": original_order.get("kp_date", "неизвестно"),
                    "customer": original_order.get("customer", "неизвестно"),
                    "plate_name": original_order.get("plate_name", ""),
                    "reinforcement": original_order.get("reinforcement", 0),
                    "load_code": cfg.normalize_load_code(
                        original_order.get("load_code", 8)
                    ),
                    "qty_remaining": 1,
                    "kp_id": original_order.get("kp_id"),
                    "is_from_secondary": True,
                }
            )

            length_key = round(result_length, 2)
            plate_lookup_by_length.setdefault(length_key, []).append(
                {
                    "kp_date": original_order.get("kp_date", "неизвестно"),
                    "customer": original_order.get("customer", "неизвестно"),
                    "plate_name": original_order.get("plate_name", ""),
                    "reinforcement": original_order.get("reinforcement", 0),
                    "qty_remaining": 1,
                    "kp_id": original_order.get("kp_id"),
                    "is_from_secondary": True,
                }
            )

    return plate_lookup_exact


def build_plan_preview(
    service: ProductionPlanningService,
    *,
    start_date: str,
    tracks_count: int,
    filter_method: Literal["all", "kp"] = "all",
    selected_kp_ids: list[int] | None = None,
    selected_plate_ids: dict[int, list[int]] | None = None,
    layout_reinforcement_order: str = "asc",
    plate_order_ctx: PlateOrderContext | None = None,
) -> dict[str, Any]:
    """Bot/API parity path: core pipeline → plan structure без persist."""
    plan_input = PlanBuildInput(
        start_date=start_date,
        tracks_count=tracks_count,
        filter_method=filter_method,
        selected_kp_ids=tuple(selected_kp_ids or ()),
        selected_plate_ids=selected_plate_ids,
        layout_reinforcement_order=layout_reinforcement_order,  # type: ignore[arg-type]
    )
    try:
        load_result, opt_result = service.run_planning_pipeline(
            plan_input=plan_input,
            layout_reinforcement_order=layout_reinforcement_order,
            plate_order_ctx=plate_order_ctx,
        )
        return service.build_plan_structure(
            load_result,
            opt_result,
            start_date=start_date,
            tracks_count=tracks_count,
            layout_reinforcement_order=layout_reinforcement_order,
        )
    except (PlanBuildError, ProductionPlanBuildError) as exc:
        raise ProductionPlanBuildError(str(exc)) from exc


def plan_structure_signature(plan: dict[str, Any]) -> dict[str, Any]:
    """Нормализованная структура плана для cross-surface сравнения."""
    days_out: dict[str, Any] = {}
    for date_key in sorted((plan.get("days") or {}).keys()):
        day = plan["days"][date_key] or {}
        track_sigs: list[dict[str, Any]] = []
        for track in day.get("tracks") or []:
            items_sig: list[dict[str, Any]] = []
            for item in track.get("items") or []:
                if not item:
                    continue
                width_raw = item.get("width") or item.get("main_w") or 1.2
                if isinstance(width_raw, (int, float)) and 0 < width_raw < 10:
                    width_mm = int(round(float(width_raw) * 1000))
                else:
                    width_mm = int(round(float(width_raw or 1200)))
                items_sig.append(
                    {
                        "length": round(float(item.get("length") or 0), 2),
                        "width_mm": width_mm,
                        "kp_id": item.get("kp_id"),
                        "plate_name": item.get("plate_name"),
                        "load_code": cfg.normalize_load_code(
                            item.get("load_code", 8)
                        ),
                    }
                )
            track_sigs.append({"items": items_sig})
        days_out[date_key] = {
            "day_number": day.get("day_number"),
            "tracks": track_sigs,
        }

    return {
        "start_date": plan.get("start_date"),
        "tracks_count": plan.get("tracks_count"),
        "total_tracks": sum(
            len(day.get("tracks") or []) for day in days_out.values()
        ),
        "days": days_out,
    }
