"""
Агрегация данных из нескольких планов (Гант, дорожки по дате).
"""
import logging
from datetime import datetime
from typing import Literal, Optional

from .plan_storage import (
    convert_lookup_keys_to_tuples,
    count_day_tracks,
    load_plan,
    load_plans_metadata,
)

logger = logging.getLogger(__name__)

_DedupMode = Literal["entry", "kp_plate"]


def get_all_tracks_from_plan(plan: dict) -> list:
    """
    Собирает все дорожки из плана в один список (для совместимости со старым кодом).
    """
    all_tracks = []

    sorted_days = sorted(plan.get('days', {}).items(), key=lambda x: x[0])

    for date_key, day_data in sorted_days:
        day_number = day_data.get('day_number', 1)

        for track in day_data.get('tracks', []):
            track_copy = track.copy() if isinstance(track, dict) else track

            if isinstance(track_copy, dict):
                track_copy['production_day'] = day_number

            all_tracks.append(track_copy)

    return all_tracks


def _entry_exists_in_lookup_list(entries: list, entry: dict, dedup_mode: _DedupMode) -> bool:
    if dedup_mode == "entry":
        return entry in entries

    for existing in entries:
        if (
            existing.get('kp_id') == entry.get('kp_id')
            and existing.get('plate_name') == entry.get('plate_name')
        ):
            return True
    return False


def _merge_plate_lookups(
    combined_exact: dict,
    combined_by_length: dict,
    plan: dict,
    *,
    dedup_mode: _DedupMode = "entry",
) -> None:
    """Объединяет plate_lookup из плана в общие словари."""
    plan_lookup_exact = convert_lookup_keys_to_tuples(plan.get('plate_lookup_exact', {}))
    for key, entries in plan_lookup_exact.items():
        if key not in combined_exact:
            combined_exact[key] = []

        for entry in entries:
            if not _entry_exists_in_lookup_list(combined_exact[key], entry, dedup_mode):
                combined_exact[key].append(entry)

    plan_lookup_by_length = plan.get('plate_lookup_by_length', {})
    for length_key, entries in plan_lookup_by_length.items():
        if isinstance(length_key, str):
            try:
                length_key = float(length_key)
            except Exception:
                pass

        if length_key not in combined_by_length:
            combined_by_length[length_key] = []

        for entry in entries:
            if not _entry_exists_in_lookup_list(combined_by_length[length_key], entry, dedup_mode):
                combined_by_length[length_key].append(entry)


def _iter_plan_tracks_for_date(plan: dict, date_key: str) -> list:
    """
    Возвращает дорожки плана на указанную дату с аннотацией источника.
    """
    if date_key not in plan.get('days', {}):
        return []

    plan_id = plan.get('id', '')
    plan_name = plan.get('name', f'План {plan_id}')
    day_data = plan['days'][date_key]
    day_tracks = day_data.get('tracks', []) or []
    count_day_tracks(day_data)

    annotated_tracks = []
    for idx, track in enumerate(day_tracks):
        if isinstance(track, dict):
            track['source_plan_id'] = plan_id
            track['source_plan_name'] = plan_name
            track['source_track_index'] = idx
        annotated_tracks.append(track)

    return annotated_tracks


def get_all_plans_gantt_data() -> Optional[dict]:
    """Собирает данные из ВСЕХ сохранённых планов для создания суммарной диаграммы Ганта."""
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])

    if not plans:
        logger.warning("[GANTT] Нет сохранённых планов для диаграммы")
        return None

    all_tracks_combined = []
    combined_plate_lookup_exact = {}
    combined_plate_lookup_by_length = {}

    earliest_date = None
    latest_date = None
    unique_dates = set()

    plans_loaded = 0

    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)

        if not plan:
            logger.warning(f"[GANTT] Не удалось загрузить план {plan_id}")
            continue

        plans_loaded += 1

        plan_tracks = get_all_tracks_from_plan(plan)
        all_tracks_combined.extend(plan_tracks)

        _merge_plate_lookups(
            combined_plate_lookup_exact,
            combined_plate_lookup_by_length,
            plan,
            dedup_mode="entry",
        )

        plan_start = plan.get('start_date')
        if plan_start:
            try:
                start_dt = datetime.strptime(plan_start, '%Y-%m-%d')
                if earliest_date is None or start_dt < earliest_date:
                    earliest_date = start_dt
            except ValueError:
                logger.warning(f"[GANTT] Неверный формат даты начала в плане {plan_id}: {plan_start}")

        for date_key in plan.get('days', {}).keys():
            unique_dates.add(date_key)
            try:
                day_dt = datetime.strptime(date_key, '%Y-%m-%d')
                if latest_date is None or day_dt > latest_date:
                    latest_date = day_dt
            except ValueError:
                logger.warning(f"[GANTT] Неверный формат даты дня: {date_key}")

    if plans_loaded == 0:
        logger.warning("[GANTT] Не удалось загрузить ни одного плана")
        return None

    if not all_tracks_combined:
        logger.warning("[GANTT] Нет дорожек в загруженных планах")
        return None

    if earliest_date is None:
        earliest_date = datetime.now()
    if latest_date is None:
        latest_date = datetime.now()

    logger.info(
        f"[GANTT] Собрано данных для диаграммы: {plans_loaded} планов, "
        f"{len(all_tracks_combined)} дорожек, период {earliest_date.strftime('%d.%m.%Y')} - "
        f"{latest_date.strftime('%d.%m.%Y')}"
    )

    return {
        'all_tracks': all_tracks_combined,
        'plate_lookup_exact': combined_plate_lookup_exact,
        'plate_lookup_by_length': combined_plate_lookup_by_length,
        'earliest_start_date': earliest_date,
        'latest_end_date': latest_date,
        'plans_count': plans_loaded,
        'total_days': len(unique_dates)
    }


def get_tracks_for_date_from_all_plans(date_key: str) -> Optional[dict]:
    """Собирает дорожки на конкретную дату из ВСЕХ сохранённых планов."""
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])

    if not plans:
        logger.warning(f"[MULTI_PLAN] Нет сохранённых планов для даты {date_key}")
        return None

    all_tracks_for_date = []
    combined_plate_lookup_exact = {}
    combined_plate_lookup_by_length = {}
    combined_orders_2d = []
    last_optimization_result = {}

    source_plan_ids = []
    plans_with_date = 0

    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)

        if not plan:
            logger.warning(f"[MULTI_PLAN] Не удалось загрузить план {plan_id}")
            continue

        if date_key not in plan.get('days', {}):
            continue

        plans_with_date += 1
        source_plan_ids.append(plan_id)

        day_tracks = _iter_plan_tracks_for_date(plan, date_key)
        all_tracks_for_date.extend(day_tracks)

        logger.info(f"[MULTI_PLAN] План {plan_id}: найдено {len(day_tracks)} дорожек на {date_key}")

        _merge_plate_lookups(
            combined_plate_lookup_exact,
            combined_plate_lookup_by_length,
            plan,
            dedup_mode="kp_plate",
        )

        plan_orders = plan.get('orders_2d', [])
        combined_orders_2d.extend(plan_orders)

        plan_opt_result = plan.get('optimization_result', {})
        if plan_opt_result:
            last_optimization_result = plan_opt_result

    if plans_with_date == 0:
        logger.warning(f"[MULTI_PLAN] Дата {date_key} не найдена ни в одном плане")
        return None

    logger.info(
        f"[MULTI_PLAN] Собрано данных для {date_key}: "
        f"{plans_with_date} планов, {len(all_tracks_for_date)} дорожек"
    )

    return {
        'tracks': all_tracks_for_date,
        'plate_lookup_exact': combined_plate_lookup_exact,
        'plate_lookup_by_length': combined_plate_lookup_by_length,
        'orders_2d': combined_orders_2d,
        'optimization_result': last_optimization_result,
        'plans_count': plans_with_date,
        'source_plans': source_plan_ids
    }
