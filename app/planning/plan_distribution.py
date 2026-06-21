"""
Распределение дорожек по дням и добавление дорожек к планам.
"""
import logging
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

from core.work_calendar import is_working_day, load_extra_workdays, load_holidays

from .plan_storage import (
    MAX_TRACKS_PER_DAY,
    count_day_tracks,
    create_plan_id,
    load_plan,
    save_plan,
    set_active_plan,
    update_plan_metadata,
)

logger = logging.getLogger(__name__)


@dataclass
class PlateLookupTables:
    """Lookup-таблицы плит для добавления к плану."""

    plate_lookup_exact: dict
    plate_lookup_by_length: dict


@dataclass
class PlanTracksPayload:
    """Параметры дорожек и связанных данных для добавления к плану."""

    new_tracks_list: list
    start_date: str
    tracks_per_day: int
    plate_lookups: PlateLookupTables
    orders_2d: list
    optimization_result: dict


def distribute_tracks_by_days(
    tracks_list: list,
    start_date: str,
    tracks_per_day: int,
    global_occupancy: Optional[Dict[str, int]] = None,
    max_per_day: int = MAX_TRACKS_PER_DAY,
) -> Dict[str, List]:
    """
    Разбивает список дорожек по дням с учётом глобальной занятости.
    """
    result: Dict[str, list] = {}

    try:
        current_date = datetime.strptime(start_date, '%Y-%m-%d')
    except ValueError:
        current_date = datetime.now()

    holidays = load_holidays()
    extra_workdays = load_extra_workdays()
    occupancy = global_occupancy or {}

    while not is_working_day(current_date.date(), holidays, extra_workdays):
        current_date += timedelta(days=1)

    track_index = 0
    while track_index < len(tracks_list):
        date_key = current_date.strftime('%Y-%m-%d')

        available = max_per_day - int(occupancy.get(date_key, 0) or 0)
        remaining = len(tracks_list) - track_index

        if available <= 0:
            logger.info(
                "[DISTRIBUTE] День %s заполнен (%s/%s) — пропускаем и ищем следующий рабочий",
                date_key,
                occupancy.get(date_key, 0),
                max_per_day,
            )
            current_date += timedelta(days=1)
            while not is_working_day(current_date.date(), holidays, extra_workdays):
                current_date += timedelta(days=1)
            continue

        chunk_size = min(tracks_per_day, available, remaining)
        if chunk_size > 0:
            result[date_key] = tracks_list[track_index:track_index + chunk_size]
            track_index += chunk_size

        current_date += timedelta(days=1)
        while not is_working_day(current_date.date(), holidays, extra_workdays):
            current_date += timedelta(days=1)

    return result


def _create_empty_plan(
    start_date: str,
    tracks_per_day: int,
    plan_name: Optional[str],
) -> Tuple[dict, str]:
    """Создаёт структуру нового плана и возвращает (plan, plan_id)."""
    plan_id = create_plan_id()

    if not plan_name:
        try:
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            plan_name = f"План с {start_dt.strftime('%d.%m.%Y')}"
        except Exception:
            plan_name = f"План {datetime.now().strftime('%d.%m.%Y %H:%M')}"

    plan = {
        'id': plan_id,
        'name': plan_name,
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'start_date': start_date,
        'tracks_count': tracks_per_day,
        'days': {},
        'plate_lookup_exact': {},
        'plate_lookup_by_length': {},
        'orders_2d': [],
        'optimization_result': {},
        'completed_days': []
    }
    logger.info(f"Создан новый план: {plan_id}")
    return plan, plan_id


def _append_tracks_to_days(
    plan: dict,
    tracks_by_day: Dict[str, list],
    tracks_per_day: int,
) -> Tuple[List[dict], List[dict]]:
    """Добавляет дорожки к дням плана. Возвращает (days_updated, days_created)."""
    days_updated: List[dict] = []
    days_created: List[dict] = []
    day_number = len(plan['days']) + 1

    for date_key, day_tracks in tracks_by_day.items():
        if date_key in plan['days']:
            existing_day = plan['days'][date_key]
            old_count = count_day_tracks(existing_day)

            existing_day['tracks'].extend(day_tracks)
            existing_day['saved_tracks_count'] = len(existing_day['tracks'])

            days_updated.append({
                'date': date_key,
                'old_count': old_count,
                'new_count': existing_day['saved_tracks_count'],
                'total': existing_day['total_tracks_count']
            })
            logger.info(
                f"День {date_key}: добавлено {len(day_tracks)} дорожек "
                f"({old_count} -> {existing_day['saved_tracks_count']})"
            )
        else:
            plan['days'][date_key] = {
                'date': date_key,
                'day_number': day_number,
                'tracks': day_tracks,
                'saved_tracks_count': len(day_tracks),
                'total_tracks_count': tracks_per_day,
                'completed': False
            }

            days_created.append({
                'date': date_key,
                'count': len(day_tracks),
                'total': tracks_per_day
            })
            logger.info(f"Создан новый день {date_key}: {len(day_tracks)} дорожек")
            day_number += 1

    return days_updated, days_created


def _merge_lookup_into_plan(
    plan: dict,
    plate_lookup_exact: dict,
    plate_lookup_by_length: dict,
) -> None:
    """Объединяет lookup-таблицы и связанные данные в план."""
    for key, value in plate_lookup_exact.items():
        str_key = str(key)
        if str_key not in plan['plate_lookup_exact']:
            plan['plate_lookup_exact'][str_key] = value
        elif isinstance(value, list):
            if isinstance(plan['plate_lookup_exact'][str_key], list):
                plan['plate_lookup_exact'][str_key].extend(value)
            else:
                plan['plate_lookup_exact'][str_key] = value

    for key, value in plate_lookup_by_length.items():
        str_key = str(key)
        if str_key not in plan['plate_lookup_by_length']:
            plan['plate_lookup_by_length'][str_key] = value
        elif isinstance(value, list):
            if isinstance(plan['plate_lookup_by_length'][str_key], list):
                plan['plate_lookup_by_length'][str_key].extend(value)
            else:
                plan['plate_lookup_by_length'][str_key] = value


def add_tracks_to_plan(
    plan_id: Optional[str],
    new_tracks_list: list,
    start_date: str,
    tracks_per_day: int,
    plate_lookup_exact: dict,
    plate_lookup_by_length: dict,
    orders_2d: list,
    optimization_result: dict,
    plan_name: Optional[str] = None,
    auto_save: bool = True,
    global_occupancy: Optional[Dict[str, int]] = None,
    max_per_day: int = MAX_TRACKS_PER_DAY,
    precomputed_tracks_by_day: Optional[Dict[str, list]] = None,
    existing_plan: Optional[dict] = None,
) -> Tuple[dict, dict]:
    """
    Добавляет дорожки к существующему плану или создаёт новый.
    """
    payload = PlanTracksPayload(
        new_tracks_list=new_tracks_list,
        start_date=start_date,
        tracks_per_day=tracks_per_day,
        plate_lookups=PlateLookupTables(
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
        ),
        orders_2d=orders_2d,
        optimization_result=optimization_result,
    )

    stats = {
        'days_updated': [],
        'days_created': [],
        'is_new_plan': False
    }

    plan = existing_plan
    if plan is None and plan_id:
        plan = load_plan(plan_id)

    if not plan:
        plan, plan_id = _create_empty_plan(
            payload.start_date,
            payload.tracks_per_day,
            plan_name,
        )
        stats['is_new_plan'] = True

    if precomputed_tracks_by_day is not None:
        tracks_by_day = precomputed_tracks_by_day
    else:
        tracks_by_day = distribute_tracks_by_days(
            payload.new_tracks_list,
            payload.start_date,
            payload.tracks_per_day,
            global_occupancy=global_occupancy,
            max_per_day=max_per_day,
        )

    days_updated, days_created = _append_tracks_to_days(
        plan,
        tracks_by_day,
        payload.tracks_per_day,
    )
    stats['days_updated'] = days_updated
    stats['days_created'] = days_created

    _merge_lookup_into_plan(
        plan,
        payload.plate_lookups.plate_lookup_exact,
        payload.plate_lookups.plate_lookup_by_length,
    )

    plan['orders_2d'].extend(payload.orders_2d)
    plan['optimization_result'] = payload.optimization_result

    if auto_save:
        save_plan(plan)
        update_plan_metadata(plan)
        set_active_plan(plan_id)

    return plan, stats
