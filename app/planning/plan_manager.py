"""
Модуль управления планами производства.

Тонкий фасад: реэкспортирует публичный API из подмодулей,
сохраняя обратную совместимость существующих импортов.
"""
import logging
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple, TypedDict

logger = logging.getLogger(__name__)

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from core import kp_db
from core.plan_track_removal import TrackRemovalError, collect_plate_returns_from_track
from core.work_calendar import is_working_day, load_extra_workdays, load_holidays

from .plan_aggregation import (
    get_all_plans_gantt_data,
    get_all_tracks_from_plan,
    get_tracks_for_date_from_all_plans,
)
from .plan_calendar import (
    get_free_tracks_for_date,
    get_global_calendar_info,
    get_global_day_occupancy,
    get_global_days_info,
)
from .plan_distribution import (
    PlateLookupTables,
    PlanTracksPayload,
    add_tracks_to_plan,
    distribute_tracks_by_days,
)
from .plan_storage import (
    BOT_DIR,
    MAX_TRACKS_PER_DAY,
    PLANS_DIR,
    PLANS_METADATA_PATH,
    InvalidPlanIdError,
    convert_lookup_keys_to_tuples,
    count_day_tracks,
    create_plan_id,
    delete_plan,
    ensure_plans_dir,
    get_active_plan,
    get_active_plan_id,
    get_all_active_plan_ids,
    get_plan_path,
    load_plan,
    load_plans_metadata,
    save_plan,
    save_plans_metadata,
    set_active_plan,
    update_plan_metadata,
)

__all__ = [
    "BOT_DIR",
    "MAX_TRACKS_PER_DAY",
    "PLANS_DIR",
    "PLANS_METADATA_PATH",
    "InvalidPlanIdError",
    "PlateLookupTables",
    "PlanTracksPayload",
    "RemoveTrackResult",
    "add_tracks_to_plan",
    "convert_lookup_keys_to_tuples",
    "count_day_tracks",
    "create_plan_id",
    "delete_plan",
    "distribute_tracks_by_days",
    "ensure_plans_dir",
    "format_plan_stats_message",
    "get_active_plan",
    "get_active_plan_id",
    "get_all_active_plan_ids",
    "get_all_plans_gantt_data",
    "get_all_tracks_from_plan",
    "get_day_tracks",
    "get_free_tracks_for_date",
    "get_global_calendar_info",
    "get_global_day_occupancy",
    "get_global_days_info",
    "get_plan_day_to_date_mapping",
    "get_plan_days_for_plate",
    "get_plan_days_info",
    "get_plan_path",
    "get_tracks_for_date_from_all_plans",
    "get_underfilled_tracks",
    "is_working_day",
    "load_extra_workdays",
    "load_holidays",
    "load_plan",
    "load_plans_metadata",
    "mark_day_completed",
    "remove_track_from_plan",
    "save_plan",
    "save_plans_metadata",
    "set_active_plan",
    "update_plan_metadata",
]


class RemoveTrackResult(TypedDict, total=False):
    """Результат удаления дорожки из плана."""

    plan_id: str
    date: str
    track_index: int
    plates_returned: int
    saved_tracks_count: int
    warnings: List[str]


def get_plan_days_info(plan: dict) -> dict:
    """Получает информацию о днях плана для отображения в календаре."""
    result = {}

    for date_key, day_data in plan.get('days', {}).items():
        result[date_key] = {
            'saved': count_day_tracks(day_data),
            'total': day_data.get('total_tracks_count', plan.get('tracks_count', 5)),
            'completed': day_data.get('completed', False),
            'day_number': day_data.get('day_number', 1)
        }

    return result


def get_plan_days_for_plate(plan_id: str, plate_name_substring: str) -> List[int]:
    """Возвращает номера дней плана, в треках которых встречается плита."""
    plan = load_plan(plan_id)
    if not plan or not plan.get('days'):
        return []
    days_with_plate = []
    for _date, day_data in sorted(plan.get('days', {}).items(), key=lambda x: x[0]):
        day_number = day_data.get('day_number')
        if day_number is None:
            continue
        for track in day_data.get('tracks', []):
            for item in track.get('items', []) or []:
                if not item:
                    continue
                name = (item.get('plate_name') or item.get('label') or '')
                if plate_name_substring in name:
                    days_with_plate.append(day_number)
                    break
                length = item.get('length') or item.get('target_length')
                width_m = item.get('width') if item.get('width') is not None else (item.get('main_w') if item.get('main_w') is not None else 1.2)
                width_m = float(width_m) if width_m is not None else 1.2
                if length is not None:
                    L = float(length)
                    if abs(L - 5.98) < 0.02 or abs(L - 5.99) < 0.02:
                        if abs(width_m - 1.2) < 0.05:
                            days_with_plate.append(day_number)
                            break
                for sec in (item.get('secondary_cuts') or []):
                    tl = sec.get('target_length')
                    if tl is not None:
                        L = float(tl)
                        if abs(L - 5.98) < 0.02 or abs(L - 5.99) < 0.02:
                            days_with_plate.append(day_number)
                            break
    return sorted(set(days_with_plate))


def get_plan_day_to_date_mapping(plan_id: str) -> dict:
    """Возвращает маппинг day_number → date для плана."""
    plan = load_plan(plan_id)
    if not plan or not plan.get('days'):
        return {}
    return {
        day_data.get('day_number'): _date
        for _date, day_data in plan.get('days', {}).items()
        if day_data.get('day_number') is not None
    }


def get_underfilled_tracks(
    plan: dict,
    *,
    max_length: float = 101.0,
    eps: float = 0.01,
) -> list[dict]:
    """Возвращает дорожки плана, которые можно дозаполнить до ``max_length`` м."""
    underfilled: list[dict] = []
    sorted_days = sorted(plan.get("days", {}).items(), key=lambda x: x[0])

    for date_key, day_data in sorted_days:
        day_number = int(day_data.get("day_number") or 0)
        for track_idx, track in enumerate(day_data.get("tracks") or []):
            if not isinstance(track, dict):
                continue
            track_length = float(track.get("length") or 0)
            free_space = round(max_length - track_length, 2)
            if free_space <= eps:
                continue
            underfilled.append(
                {
                    "date_key": date_key,
                    "day_number": day_number,
                    "track_idx": track_idx,
                    "track": track,
                    "track_length": round(track_length, 2),
                    "free_space": free_space,
                    "max_reinforcement": float(track.get("max_reinforcement") or 0),
                    "load_code": track.get("load_code", 8),
                }
            )
    return underfilled


def mark_day_completed(plan_id: str, date_key: str) -> bool:
    """Отмечает день как выполненный."""
    plan = load_plan(plan_id)
    if not plan:
        return False

    if date_key in plan.get('days', {}):
        plan['days'][date_key]['completed'] = True

        if 'completed_days' not in plan:
            plan['completed_days'] = []

        day_number = plan['days'][date_key].get('day_number', 1)
        if day_number not in plan['completed_days']:
            plan['completed_days'].append(day_number)
            plan['completed_days'].sort()

        save_plan(plan)
        return True

    return False


def get_day_tracks(plan: dict, day_number: int) -> Tuple[list, str]:
    """Получает дорожки для конкретного дня по номеру."""
    for date_key, day_data in plan.get('days', {}).items():
        if day_data.get('day_number') == day_number:
            return day_data.get('tracks', []), date_key

    return [], ''


def format_plan_stats_message(stats: dict) -> str:
    """Форматирует статистику изменений плана для отображения пользователю."""
    lines = []

    if stats['is_new_plan']:
        lines.append("✅ Создан новый план!\n")
    else:
        lines.append("✅ Дорожки добавлены к плану!\n")

    if stats['days_updated']:
        lines.append("📊 Обновлённые дни:")
        for day in stats['days_updated']:
            date_str = datetime.strptime(day['date'], '%Y-%m-%d').strftime('%d.%m')
            lines.append(f"  • {date_str}: {day['old_count']}/{day['total']} → {day['new_count']}/{day['total']}")

    if stats['days_created']:
        lines.append("\n📆 Новые дни:")
        for day in stats['days_created']:
            date_str = datetime.strptime(day['date'], '%Y-%m-%d').strftime('%d.%m')
            lines.append(f"  • {date_str}: {day['count']}/{day['total']}")

    total_days = len(stats['days_updated']) + len(stats['days_created'])
    lines.append(f"\nВсего затронуто дней: {total_days}")

    return '\n'.join(lines)


def remove_track_from_plan(
    plan_id: str,
    date_key: str,
    track_index: int,
    *,
    db_path: str,
    actor: Optional[str] = None,
) -> RemoveTrackResult:
    """
    Удаляет одну дорожку из плана на указанную дату.

    Порядок: сначала возврат плит в БД (транзакция), затем правка JSON-плана.
    """
    plan = load_plan(plan_id)
    if not plan:
        raise TrackRemovalError(
            f"План {plan_id!r} не найден",
            code="plan_not_found",
        )

    day = plan.get("days", {}).get(date_key)
    if day is None:
        raise TrackRemovalError(
            f"День {date_key!r} не найден в плане {plan_id!r}",
            code="day_not_found",
        )

    if day.get("completed"):
        raise TrackRemovalError(
            f"День {date_key!r} уже завершён — удаление дорожки невозможно",
            code="day_already_completed",
        )

    tracks = day.get("tracks") or []
    if track_index < 0 or track_index >= len(tracks):
        raise TrackRemovalError(
            f"Недопустимый track_index={track_index} (дорожек в дне: {len(tracks)})",
            code="invalid_track_index",
        )

    track = tracks[track_index]
    id_qty, legacy_identity_qty = collect_plate_returns_from_track(track)

    if not id_qty and not legacy_identity_qty:
        raise TrackRemovalError(
            "В дорожке не найдено kp_plate_id и legacy-идентичностей — "
            "удаление дорожки невозможно",
            code="no_plate_identity",
        )

    expected_count = sum(id_qty.values()) + sum(legacy_identity_qty.values())

    conn = sqlite3.connect(db_path)
    plates_returned = 0
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        db_result = kp_db.return_plate_rows_for_plan(
            plan_id,
            id_qty,
            db_path,
            actor=actor,
            legacy_identity_qty=legacy_identity_qty or None,
            _external_conn=conn,
        )
        plates_returned = int(db_result.get("plates_returned") or 0)
        db_warnings = db_result.get("warnings") or []
        if db_warnings or plates_returned < expected_count:
            conn.rollback()
            detail = (
                f"ожидалось вернуть {expected_count} плит(ы), "
                f"фактически {plates_returned}"
            )
            if db_warnings:
                detail = f"{detail}; предупреждения: {'; '.join(db_warnings)}"
            raise TrackRemovalError(
                f"Неполный возврат плит в производство: {detail}",
                code="incomplete_return",
            )
        conn.commit()
    except TrackRemovalError:
        raise
    except Exception as exc:
        conn.rollback()
        logger.exception(
            "[REMOVE_TRACK] Ошибка возврата плит plan_id=%s date=%s track_index=%s",
            plan_id,
            date_key,
            track_index,
        )
        raise TrackRemovalError(
            f"Не удалось вернуть плиты в производство: {exc}",
            code="db_return_failed",
        ) from exc
    finally:
        conn.close()

    tracks.pop(track_index)
    day["saved_tracks_count"] = len(tracks)
    saved_tracks_count = day["saved_tracks_count"]

    if not tracks:
        del plan["days"][date_key]
        saved_tracks_count = 0

    if not save_plan(plan):
        raise TrackRemovalError(
            f"Плиты возвращены в БД, но не удалось сохранить план {plan_id!r} на диск",
            code="plan_save_failed",
        )
    update_plan_metadata(plan)

    return {
        "plan_id": plan_id,
        "date": date_key,
        "track_index": track_index,
        "plates_returned": plates_returned,
        "saved_tracks_count": saved_tracks_count,
    }
