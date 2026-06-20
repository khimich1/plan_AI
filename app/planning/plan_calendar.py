"""
Глобальный календарь и занятость дорожек по всем планам.
"""
import logging
from datetime import datetime, timedelta
from typing import Dict, Optional

from .plan_storage import (
    MAX_TRACKS_PER_DAY,
    count_day_tracks,
    load_plan,
    load_plans_metadata,
)

logger = logging.getLogger(__name__)


def get_global_day_occupancy(exclude_plan_id: Optional[str] = None) -> Dict[str, int]:
    """Подсчитывает занятость дорожек по ВСЕМ планам для каждой даты."""
    occupancy = {}

    metadata = load_plans_metadata()

    for plan_meta in metadata.get('plans', []):
        plan_id = plan_meta.get('id')

        if plan_id == exclude_plan_id:
            continue

        plan = load_plan(plan_id)
        if not plan:
            continue

        for date_key, day_data in plan.get('days', {}).items():
            tracks_count = count_day_tracks(day_data)

            if date_key in occupancy:
                occupancy[date_key] += tracks_count
            else:
                occupancy[date_key] = tracks_count

    return occupancy


def get_free_tracks_for_date(date: str, exclude_plan_id: Optional[str] = None) -> int:
    """Возвращает количество свободных дорожек на указанную дату."""
    occupancy = get_global_day_occupancy(exclude_plan_id)
    occupied = occupancy.get(date, 0)
    free = MAX_TRACKS_PER_DAY - occupied
    return max(0, free)


def get_global_days_info(plan: dict) -> Dict[str, dict]:
    """Получает информацию о днях с учётом ГЛОБАЛЬНОЙ загруженности."""
    result = {}

    global_occupancy = get_global_day_occupancy()

    for date_key, day_data in plan.get('days', {}).items():
        result[date_key] = {
            'occupied': global_occupancy.get(date_key, 0),
            'max': MAX_TRACKS_PER_DAY,
            'completed': day_data.get('completed', False),
            'day_number': day_data.get('day_number', 1)
        }

    return result


def get_global_calendar_info() -> Optional[dict]:
    """Собирает информацию о ВСЕХ днях из ВСЕХ планов для отображения единого календаря."""
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])

    if not plans:
        logger.warning("[GLOBAL_CALENDAR] Нет сохранённых планов")
        return None

    all_dates_data = {}
    earliest_date = None
    latest_date = None
    total_tracks_count = 0

    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)

        if not plan:
            logger.warning(f"[GLOBAL_CALENDAR] Не удалось загрузить план {plan_id}")
            continue

        for date_key, day_data in plan.get('days', {}).items():
            try:
                day_dt = datetime.strptime(date_key, '%Y-%m-%d')
                if earliest_date is None or day_dt < earliest_date:
                    earliest_date = day_dt
                if latest_date is None or day_dt > latest_date:
                    latest_date = day_dt
            except ValueError:
                logger.warning(f"[GLOBAL_CALENDAR] Неверный формат даты: {date_key}")
                continue

            tracks_count = count_day_tracks(day_data)
            is_completed = day_data.get('completed', False)

            if date_key not in all_dates_data:
                all_dates_data[date_key] = {
                    'occupied': 0,
                    'completed': False
                }

            all_dates_data[date_key]['occupied'] += tracks_count
            if is_completed:
                all_dates_data[date_key]['completed'] = True

            total_tracks_count += tracks_count

    if earliest_date is None or latest_date is None:
        logger.warning("[GLOBAL_CALENDAR] Не удалось определить диапазон дат")
        return None

    total_days = (latest_date - earliest_date).days + 1

    days_info = {}
    completed_days = []

    for day_offset in range(total_days):
        current_date = earliest_date + timedelta(days=day_offset)
        date_key = current_date.strftime('%Y-%m-%d')
        day_number = day_offset + 1

        date_data = all_dates_data.get(date_key, {'occupied': 0, 'completed': False})

        days_info[date_key] = {
            'occupied': date_data['occupied'],
            'max': MAX_TRACKS_PER_DAY,
            'completed': date_data['completed'],
            'day_number': day_number
        }

        if date_data['completed']:
            completed_days.append(day_number)

    logger.info(
        f"[GLOBAL_CALENDAR] Создан глобальный календарь: {len(plans)} планов, "
        f"{total_days} дней ({earliest_date.strftime('%d.%m.%Y')} - {latest_date.strftime('%d.%m.%Y')}), "
        f"{total_tracks_count} дорожек"
    )

    return {
        'start_date': earliest_date.strftime('%Y-%m-%d'),
        'total_days': total_days,
        'days_info': days_info,
        'completed_days': completed_days,
        'plans_count': len(plans),
        'tracks_count': total_tracks_count
    }
