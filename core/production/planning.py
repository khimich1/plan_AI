"""Pure production planning pipeline: validate → load → optimize → persist."""

from __future__ import annotations

import logging
from collections import Counter, defaultdict
from collections.abc import Callable
from datetime import datetime
from typing import Any

from core import kp_db_plates
from core.concrete_grade_resolver import enrich_orders_2d_concrete_grade, resolve_concrete_grade_from_order
from core.domain.plate_order import PlateOrder, normalize_load_code
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.optimization.layout_runtime_snapshot import build_layout_runtime_snapshot_from_plate_order_context
from core.optimization.result_contract import is_optimization_success
from core.plan_commit import PlanCommitError, commit_plan_plates
from core.plate_attribution import (
    backfill_assignment_identity,
    backfill_track_items_identity,
)
from core.plate_order_context import PlateOrderContext
from core.production.dto import (
    FilterMethod,
    LoadConfig,
    LoadResult,
    OptimizeConfig,
    OptimizeResult,
    PersistConfig,
    PersistResult,
    PlanBuildInput,
)
from core.production.errors import PlanBuildError
from core.production.ports import PlanLoadPort, PlanPersistPort
from core.reinforcement_db import get_reinforcement
from core.rescue_tracks import build_rescue_tracks
from core.serialization import strip_plate_audit_from_plan
from core.ports.visualization import build_layout_sequence
from core.visualization import (
    LayoutIntegrityError,
    TrackLayoutInvariantError,
    split_sequence_into_tracks,
)

logger = logging.getLogger(__name__)


def validate(plan_input: PlanBuildInput) -> None:
    """Validate planning request parameters."""
    try:
        datetime.strptime(plan_input.start_date, "%Y-%m-%d")
    except ValueError as exc:
        raise PlanBuildError(
            f"Неверный формат start_date (ожидается YYYY-MM-DD): {plan_input.start_date}"
        ) from exc
    if not (1 <= plan_input.tracks_count <= 50):
        raise PlanBuildError("tracks_count должен быть от 1 до 50.")
    if plan_input.filter_method == "kp" and not plan_input.selected_kp_ids:
        raise PlanBuildError(
            "Для filter_method='kp' нужно передать selected_kp_ids."
        )


def load(
    plan_input: PlanBuildInput,
    *,
    config: LoadConfig,
    plan_load: PlanLoadPort,
) -> LoadResult:
    """Load KPs and plates via repository port, build orders for optimization."""
    validate(plan_input)

    kp_list = _load_kp_list(
        plan_load=plan_load,
        filter_method=plan_input.filter_method,
        selected_kp_ids=list(plan_input.selected_kp_ids or ()),
    )
    if not kp_list:
        raise PlanBuildError("Нет подходящих КП для производства.")

    kp_list.sort(key=lambda x: x["date"])

    plates_by_date_reinforcement = _load_plates_for_kps(
        plan_load=plan_load,
        pb_db_path=config.pb_db_path,
        kp_list=kp_list,
        selected_plate_ids=plan_input.selected_plate_ids or {},
        selected_plate_qty=plan_input.selected_plate_qty or {},
    )

    selected_plates = _flatten_plates(plates_by_date_reinforcement)
    if not selected_plates:
        raise PlanBuildError("Не найдено плит для планирования.")

    orders_2d, plate_lookup_exact, plate_lookup_by_length = _build_orders(selected_plates)

    return LoadResult(
        kp_list=kp_list,
        selected_plates=selected_plates,
        orders_2d=orders_2d,
        plate_lookup_exact=plate_lookup_exact,
        plate_lookup_by_length=plate_lookup_by_length,
    )


def optimize(
    load_result: LoadResult,
    *,
    config: OptimizeConfig,
    plate_order_ctx: PlateOrderContext | None = None,
) -> OptimizeResult:
    """Run 2D optimization, layout split, rescue and identity backfill."""
    orders_2d = load_result.orders_2d
    layout_reinforcement_order = config.layout_reinforcement_order

    if not orders_2d:
        return OptimizeResult(all_tracks_list=[], optimization_result={})

    enrich_orders_2d_concrete_grade(orders_2d, db_path=str(config.pb_db_path))

    core_order = PlateOrder.from_orders_2d(orders_2d)
    plate_ctx = plate_order_ctx or PlateOrderContext.fresh_empty()
    plate_ctx.hydrate_from_order(core_order)

    with plate_ctx.bound():
        optimization_result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)

    backfilled = backfill_assignment_identity(
        optimization_result.get("plate_assignments", []) or [],
        orders_2d,
    )
    if backfilled:
        logger.info(
            "[CORE-PLAN] Восстановлена identity у %s plate_assignments-записей",
            backfilled,
        )

    _log_unmapped_optimizer_assignments(optimization_result)
    if (
        not is_optimization_success(optimization_result)
        or optimization_result.get("total_plates", 0) == 0
    ):
        return OptimizeResult(all_tracks_list=[], optimization_result=optimization_result)

    all_loads = sorted({int(float(item.get("load_code", 8))) for item in orders_2d})
    optimization_result["loads_in_group"] = all_loads
    plan_by_load = {"all": optimization_result}
    load_map = {load: ["all"] for load in all_loads}

    plate_ctx.load_optimization_snapshot(
        optimization_result=optimization_result,
        plan_by_load=plan_by_load,
        load_to_reinforcement_map=load_map,
    )

    with plate_ctx.bound():
        layout_rt = build_layout_runtime_snapshot_from_plate_order_context(
            plate_ctx,
            layout_reinforcement_order=layout_reinforcement_order,  # type: ignore[arg-type]
        )
        seq = build_layout_sequence(runtime=layout_rt)
        try:
            all_tracks_list = split_sequence_into_tracks(
                seq,
                strict_layout_integrity=True,
            ) or []
        except LayoutIntegrityError as exc:
            raise PlanBuildError(
                f"Нарушена целостность раскладки дорожек: {exc}"
            ) from exc
        except TrackLayoutInvariantError as exc:
            raise PlanBuildError(
                f"Не удалось разложить дорожки: нужна целая плита в начале — {exc}"
            ) from exc

        if config.track_top_up_from_following:
            from core.track_top_up import top_up_tracks_from_following

            top_up_tracks_from_following(all_tracks_list)

    try:
        plate_assignments = optimization_result.get("plate_assignments", []) or []
        rescue_tracks, missing_counts, rescue_assignments = build_rescue_tracks(
            orders_2d=orders_2d,
            plate_assignments=plate_assignments,
        )
        if rescue_tracks:
            logger.info(
                "[CORE-PLAN] RESCUE: добавлено %s дорожек / %s assignments, "
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
        logger.exception("[CORE-PLAN] RESCUE-логика упала — продолжаем без rescue")

    backfilled_items = backfill_track_items_identity(
        all_tracks_list,
        orders_2d,
    )
    if backfilled_items:
        logger.info(
            "[CORE-PLAN] Восстановлена identity у %s track items "
            "(root + secondary_cuts)",
            backfilled_items,
        )

    try:
        fallback_tracks, fallback_missing = _build_assignment_gap_fallback_tracks(
            plate_assignments=optimization_result.get("plate_assignments", []) or [],
            tracks_list=all_tracks_list,
        )
        if fallback_tracks:
            logger.warning(
                "[CORE-PLAN] FALLBACK track-gap: добавлено %s дорожек для "
                "%s плит из assignments, отсутствующих в tracks: %s",
                len(fallback_tracks),
                sum(fallback_missing.values()),
                fallback_missing,
            )
            all_tracks_list.extend(fallback_tracks)
    except Exception:
        logger.exception(
            "[CORE-PLAN] FALLBACK track-gap логика упала — продолжаем без неё"
        )

    return OptimizeResult(
        all_tracks_list=all_tracks_list,
        optimization_result=optimization_result,
    )


def persist(
    load_result: LoadResult,
    optimize_result: OptimizeResult,
    config: PersistConfig,
    repo: PlanPersistPort,
    *,
    ensure_unique_plan_id: Callable[[], str] | None = None,
) -> PersistResult:
    """Build plan structure, commit plate statuses, save via repository."""
    all_tracks_list = list(optimize_result.all_tracks_list)
    optimization_result = dict(optimize_result.optimization_result or {})

    if not all_tracks_list:
        raise PlanBuildError("Оптимизация не дала результата.")

    global_occupancy = repo.get_global_occupancy(
        exclude_plan_id=config.active_plan_id,
    )

    start_date = config.start_date
    tracks_per_day_effective = config.tracks_count
    active_plan_id = config.active_plan_id
    precomputed_tracks_by_day: dict[str, list[dict[str, Any]]] | None = None
    fill_targets = list(config.fill_targets or ())

    if fill_targets:
        validate_fill_targets(
            fill_targets,
            global_occupancy,
            max_tracks_per_day=config.max_tracks_per_day,
        )

        cap = sum(int(t["tracks"]) for t in fill_targets)
        if cap < len(all_tracks_list):
            logger.info(
                "[CORE-PLAN] fill_targets cap=%s, обрезаем дорожек: %s -> %s",
                cap,
                len(all_tracks_list),
                cap,
            )
            all_tracks_list = all_tracks_list[:cap]
            optimization_result = trim_assignments_to_tracks(
                optimization_result=optimization_result,
                kept_tracks=all_tracks_list,
            )

        precomputed_tracks_by_day = build_tracks_by_day_from_targets(
            kept_tracks=all_tracks_list,
            fill_targets=fill_targets,
        )
        active_plan_id = None
        sorted_dates = sorted(t["date"] for t in fill_targets)
        start_date = sorted_dates[0]
        tracks_per_day_effective = max(int(t["tracks"]) for t in fill_targets)

    plan, stats = repo.build_plan_from_tracks(
        plan_id=active_plan_id,
        new_tracks_list=all_tracks_list,
        start_date=start_date,
        tracks_per_day=tracks_per_day_effective,
        plate_lookup_exact=load_result.plate_lookup_exact,
        plate_lookup_by_length=load_result.plate_lookup_by_length,
        orders_2d=load_result.orders_2d,
        optimization_result=optimization_result,
        plan_name=config.plan_name,
        auto_save=False,
        global_occupancy=global_occupancy,
        precomputed_tracks_by_day=precomputed_tracks_by_day,
    )

    plan_id = plan["id"]
    if stats.get("is_new_plan") and ensure_unique_plan_id is not None:
        while repo.get(plan_id) is not None:
            plan_id = ensure_unique_plan_id()
            plan["id"] = plan_id

    plan["layout_reinforcement_order"] = config.layout_reinforcement_order

    tracks_by_day_for_commit = _tracks_by_day_from_plan(plan)

    try:
        commit_plan_plates(
            plan_id=plan_id,
            orders_2d=load_result.orders_2d,
            optimization_result=optimization_result,
            all_tracks_list=all_tracks_list,
            db_path=config.plita_db_path,
            tracks_by_day=tracks_by_day_for_commit,
        )
    except PlanCommitError as exc:
        logger.error(
            "[CORE-PLAN] Не удалось закоммитить план %s: %s", plan_id, exc
        )
        raise PlanBuildError(str(exc)) from exc

    try:
        if stats.get("is_new_plan"):
            repo.create(plan)
        else:
            record = repo.get(plan_id)
            if record is None:
                repo.create(plan)
            else:
                repo.save(plan, expected_version=record["version"])
        repo.set_active(plan_id)
    except Exception as exc:
        logger.exception(
            "[CORE-PLAN] Ошибка записи плана %s в SQLite: %s", plan_id, exc
        )
        try:
            kp_db_plates.return_plan_plates_to_production(
                plan_id, config.plita_db_path
            )
        except Exception:
            logger.exception(
                "[CORE-PLAN] Не удалось откатить плиты для плана %s", plan_id
            )
        raise PlanBuildError("Не удалось сохранить план.") from exc

    saved_record = repo.get(plan_id)
    safe_plan = strip_plate_audit_from_plan(plan)
    if saved_record is not None:
        safe_plan = {**safe_plan, "version": saved_record["version"]}

    summary = {
        "total_tracks": len(all_tracks_list),
        "total_days": len(safe_plan.get("days", {})),
        "selected_plates_count": sum(
            int(p.get("qty", 0) or 0) for p in load_result.selected_plates
        ),
        "kp_count": len(load_result.kp_list),
    }

    logger.info(
        "[CORE-PLAN] Построен план %s: треков=%s, дней=%s, КП=%s",
        safe_plan.get("id"),
        summary["total_tracks"],
        summary["total_days"],
        summary["kp_count"],
    )

    return PersistResult(plan=safe_plan, stats=stats, summary=summary)


def validate_fill_targets(
    fill_targets: list[dict[str, Any]],
    global_occupancy: dict[str, int],
    *,
    max_tracks_per_day: int,
) -> None:
    """Проверяет, что каждая дата существует и имеет достаточно свободных слотов."""
    if not fill_targets:
        raise PlanBuildError("fill_targets пуст.")
    for t in fill_targets:
        date = t.get("date") or ""
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError as exc:
            raise PlanBuildError(
                f"Неверный формат даты в fill_targets: {date}"
            ) from exc
        occupied = int(global_occupancy.get(date, 0) or 0)
        free = max(0, max_tracks_per_day - occupied)
        requested = int(t.get("tracks", 0) or 0)
        if requested <= 0:
            raise PlanBuildError(
                f"На {date} запрошено {requested} дорожек — должно быть >= 1."
            )
        if requested > free:
            raise PlanBuildError(
                f"На {date} свободно {free} дорожек, запрошено {requested}."
            )


def build_tracks_by_day_from_targets(
    *,
    kept_tracks: list[dict[str, Any]],
    fill_targets: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    """Раскладывает kept_tracks по датам строго в порядке fill_targets."""
    result: dict[str, list[dict[str, Any]]] = {}
    cursor = 0
    for t in fill_targets:
        requested = int(t["tracks"])
        chunk = kept_tracks[cursor:cursor + requested]
        if chunk:
            result[t["date"]] = chunk
        cursor += requested
    return result


def trim_assignments_to_tracks(
    *,
    optimization_result: dict[str, Any],
    kept_tracks: list[dict[str, Any]],
) -> dict[str, Any]:
    """Оставляет в plate_assignments столько плит, сколько лежит в kept_tracks."""
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
            filtered.append(assignment)
            continue
        key = (assignment.get("kp_id"), assignment.get("plate_name") or "")
        if remaining.get(key, 0) > 0:
            filtered.append(assignment)
            remaining[key] -= 1

    return {**optimization_result, "plate_assignments": filtered}


def _tracks_by_day_from_plan(plan: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    tracks_by_day: dict[str, list[dict[str, Any]]] = {}
    for date_key, day_data in (plan.get("days") or {}).items():
        day_number = int((day_data or {}).get("day_number") or 0)
        day_tracks = (day_data or {}).get("tracks") or []
        for track in day_tracks:
            if isinstance(track, dict):
                track.setdefault("production_day", day_number)
        tracks_by_day[date_key] = day_tracks
    return tracks_by_day


# ----- load helpers (from production_planning_service) -----


def _load_kp_list(
    *,
    plan_load: PlanLoadPort,
    filter_method: FilterMethod,
    selected_kp_ids: list[int],
) -> list[dict[str, Any]]:
    rows = plan_load.fetch_kps_in_production(
        filter_method=filter_method,
        selected_kp_ids=selected_kp_ids,
    )

    result: list[dict[str, Any]] = []
    for kp_id, exec_terms, customer in rows:
        exec_date = _parse_exec_date(exec_terms)
        result.append(
            {
                "kp_id": kp_id,
                "date": exec_date,
                "customer": customer or "",
            }
        )
    return result


def _parse_exec_date(raw: str | None) -> datetime:
    if not raw:
        return datetime.now()
    try:
        return datetime.strptime(raw, "%d.%m.%Y")
    except ValueError:
        return datetime.now()


def _qty_override_for_plate(
    selected_plate_qty: dict[int, dict[int, int]],
    kp_id: int,
    plate_id: int,
) -> int | None:
    by_kp = selected_plate_qty.get(kp_id) or selected_plate_qty.get(str(kp_id))  # type: ignore[arg-type]
    if not by_kp:
        return None
    raw = by_kp.get(plate_id) if plate_id in by_kp else by_kp.get(str(plate_id))  # type: ignore[arg-type]
    if raw is None:
        return None
    return int(raw)


def _load_plates_for_kps(
    *,
    plan_load: PlanLoadPort,
    pb_db_path: str,
    kp_list: list[dict[str, Any]],
    selected_plate_ids: dict[int, list[int]],
    selected_plate_qty: dict[int, dict[int, int]],
) -> dict[datetime, dict[float, list[dict[str, Any]]]]:
    plates: dict[datetime, dict[float, list[dict[str, Any]]]] = defaultdict(
        lambda: defaultdict(list)
    )

    for kp in kp_list:
        kp_id = kp["kp_id"]
        plate_ids = selected_plate_ids.get(kp_id) or selected_plate_ids.get(str(kp_id))  # type: ignore[arg-type]
        if plate_ids is not None and len(plate_ids) == 0:
            continue

        rows = plan_load.fetch_plates_in_production_for_kp(
            kp_id=kp_id,
            plate_ids=list(plate_ids) if plate_ids else None,
        )

        for row in rows:
            plate_id = int(row[0])
            plate_name = row[1]
            length_m = row[2]
            width_m = row[3]
            load_class = row[4] or 800
            db_qty = int(row[5] or 0)
            length_dm_raw = (row[6] or "") if len(row) > 6 else ""
            concrete_grade_raw = str(row[7] or "").strip() if len(row) > 7 else ""

            qty_override = _qty_override_for_plate(
                selected_plate_qty, kp_id, plate_id
            )
            if qty_override is not None:
                if qty_override > db_qty:
                    raise PlanBuildError(
                        f"КП #{kp_id}, плита #{plate_id}: запрошено {qty_override}, "
                        f"доступно {db_qty}."
                    )
                qty = qty_override
            else:
                qty = db_qty

            if qty <= 0:
                continue

            load_code = normalize_load_code(load_class // 100)
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
                db_path=pb_db_path,
                allow_fallback=True,
            )
            if reinforcement is None:
                reinforcement = 999.0

            concrete_grade = concrete_grade_raw
            if not concrete_grade:
                concrete_grade = resolve_concrete_grade_from_order(
                    {
                        "concrete_grade": None,
                        "plate_name": plate_name or "",
                        "length": length_m or 0.0,
                        "load_code": load_code,
                    },
                    db_path=pb_db_path,
                )

            plates[kp["date"]][reinforcement].append(
                {
                    "plate_name": plate_name,
                    "length": length_m,
                    "width": width_mm,
                    "load_code": load_code,
                    "qty": qty,
                    "reinforcement": reinforcement,
                    "kp_id": kp_id,
                    "kp_plate_id": plate_id,
                    "kp_date": kp["date"].strftime("%d.%m.%Y"),
                    "customer": kp["customer"],
                    "length_dm_raw": length_dm_raw,
                    "concrete_grade": concrete_grade,
                }
            )

    return plates


def _flatten_plates(
    plates_by_date_reinforcement: dict[datetime, dict[float, list[dict[str, Any]]]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for kp_date in sorted(plates_by_date_reinforcement.keys()):
        reinforcements = plates_by_date_reinforcement[kp_date]
        for reinforcement in sorted(reinforcements.keys()):
            result.extend(reinforcements[reinforcement])
    return result


def _build_orders(
    selected_plates: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict, dict]:
    orders_2d: list[dict[str, Any]] = []
    for plate in selected_plates:
        orders_2d.append(
            {
                "length": plate["length"],
                "width": plate["width"],
                "qty": plate["qty"],
                "load_code": normalize_load_code(plate["load_code"]),
                "reinforcement": plate["reinforcement"],
                "kp_date": plate.get("kp_date", "неизвестно"),
                "customer": plate.get("customer", "неизвестно"),
                "plate_name": plate.get("plate_name", ""),
                "kp_id": plate.get("kp_id"),
                "length_dm_raw": plate.get("length_dm_raw", "") or "",
                "concrete_grade": plate.get("concrete_grade"),
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
            "load_code": normalize_load_code(order.get("load_code", 8)),
            "qty_remaining": order.get("qty", 1),
            "kp_id": order.get("kp_id"),
            "length_dm_raw": order.get("length_dm_raw", "") or "",
            "concrete_grade": order.get("concrete_grade"),
        }
        plate_lookup_exact.setdefault((length_key, width_key), []).append(entry_exact)

        entry_by_length = {
            "kp_date": order.get("kp_date", "неизвестно"),
            "customer": order.get("customer", "неизвестно"),
            "plate_name": order.get("plate_name", ""),
            "reinforcement": order.get("reinforcement", 0),
            "load_code": normalize_load_code(order.get("load_code", 8)),
            "qty_remaining": order.get("qty", 1),
            "kp_id": order.get("kp_id"),
            "length_dm_raw": order.get("length_dm_raw", "") or "",
            "concrete_grade": order.get("concrete_grade"),
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


# ----- optimize helpers -----


def _build_assignment_gap_fallback_tracks(
    *,
    plate_assignments: list[dict[str, Any]],
    tracks_list: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], Counter]:
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
        fallback_tracks.append({
            "label": "FALLBACK",
            "placement_status": "fallback",
            "items": [item],
            "length": length,
        })
        missing_counter[(assignment.get("kp_id"), assignment.get("plate_name") or "")] += 1

    return fallback_tracks, missing_counter


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
            "[CORE-PLAN] Optimizer assignments без exact identity: count=%s sample=%s",
            len(unmapped),
            unmapped[:10],
        )
