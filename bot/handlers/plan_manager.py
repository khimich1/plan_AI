"""
Модуль управления планами производства.

Содержит вспомогательные функции для:
- Загрузки и сохранения планов
- Работы с метаданными планов
- Распределения дорожек по дням
- Добавления дорожек к существующим планам
"""
import json
import os
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Добавляем путь к корню проекта для импорта core модулей
BOT_DIR_FOR_IMPORT = Path(__file__).parent.parent
PROJECT_ROOT_FOR_IMPORT = BOT_DIR_FOR_IMPORT.parent
sys.path.insert(0, str(PROJECT_ROOT_FOR_IMPORT))

from core import kp_db

# Пути к файлам
BOT_DIR = Path(__file__).parent.parent
PLANS_DIR = BOT_DIR / "data" / "plans"
PLANS_METADATA_PATH = BOT_DIR / "data" / "plans_metadata.json"

# Глобальный максимум дорожек в день (константа для всех планов)
MAX_TRACKS_PER_DAY = 5


def convert_lookup_keys_to_tuples(lookup_dict: dict) -> dict:
    """
    Конвертирует строковые ключи lookup-словарей обратно в кортежи или числа.
    
    JSON не поддерживает кортежи как ключи, поэтому при сохранении 
    они конвертируются в строки "(длина, ширина)" или "длина".
    Эта функция восстанавливает исходный формат.
    
    Args:
        lookup_dict: словарь со строковыми ключами
    
    Returns:
        dict: словарь с кортежами/числами в качестве ключей
    """
    import ast
    result = {}
    
    for key, value in lookup_dict.items():
        if isinstance(key, str):
            try:
                # Парсим строку
                parsed = ast.literal_eval(key)
                # Если это список или кортеж - конвертируем в кортеж
                if isinstance(parsed, (list, tuple)):
                    result[tuple(parsed)] = value
                else:
                    # Если это число - оставляем как число
                    result[parsed] = value
            except (ValueError, SyntaxError, TypeError):
                # Если не получилось распарсить - оставляем как есть
                result[key] = value
        else:
            # Если ключ уже кортеж/число - просто копируем
            result[key] = value
    
    return result


def ensure_plans_dir():
    """Создаёт папку для планов, если её нет."""
    PLANS_DIR.mkdir(parents=True, exist_ok=True)


def create_plan_id() -> str:
    """
    Генерирует уникальный ID плана на основе текущего времени.
    
    Returns:
        str: ID в формате 'plan_YYYYMMDD_HHMMSS'
    """
    return f"plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def load_plans_metadata() -> dict:
    """
    Загружает метаданные всех планов.
    
    Returns:
        dict: Словарь с ключами 'plans' (список планов) и 'active_plan_id'
    """
    if not PLANS_METADATA_PATH.exists():
        return {"plans": [], "active_plan_id": None}
    
    try:
        with open(PLANS_METADATA_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.exception(f"Ошибка загрузки метаданных планов: {e}")
        return {"plans": [], "active_plan_id": None}


def save_plans_metadata(metadata: dict):
    """
    Сохраняет метаданные планов.
    
    Args:
        metadata: Словарь с метаданными планов
    """
    ensure_plans_dir()
    try:
        with open(PLANS_METADATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.exception(f"Ошибка сохранения метаданных планов: {e}")


def get_plan_path(plan_id: str) -> Path:
    """Возвращает путь к файлу плана."""
    return PLANS_DIR / f"{plan_id}.json"


def load_plan(plan_id: str) -> Optional[dict]:
    """
    Загружает конкретный план по ID.
    
    Args:
        plan_id: ID плана
        
    Returns:
        dict или None: Данные плана или None если не найден
    """
    plan_path = get_plan_path(plan_id)
    if not plan_path.exists():
        logger.warning(f"План {plan_id} не найден: {plan_path}")
        return None
    
    try:
        with open(plan_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.exception(f"Ошибка загрузки плана {plan_id}: {e}")
        return None


def save_plan(plan_data: dict):
    """
    Сохраняет план в файл.
    
    Args:
        plan_data: Данные плана (должен содержать 'id')
    """
    ensure_plans_dir()
    plan_id = plan_data.get('id')
    if not plan_id:
        logger.error("План не содержит ID!")
        return
    
    plan_path = get_plan_path(plan_id)
    try:
        with open(plan_path, 'w', encoding='utf-8') as f:
            json.dump(plan_data, f, ensure_ascii=False, indent=2)
        logger.info(f"План {plan_id} сохранён в {plan_path}")
    except Exception as e:
        logger.exception(f"Ошибка сохранения плана {plan_id}: {e}")


def delete_plan(plan_id: str) -> bool:
    """
    Удаляет план и обновляет метаданные.
    
    ВАЖНО: При удалении плана все его плиты возвращаются в статус 'в производстве',
    чтобы они снова стали доступны для планирования.
    
    Args:
        plan_id: ID плана для удаления
        
    Returns:
        bool: True если удаление успешно
    """
    plan_path = get_plan_path(plan_id)
    
    # === ВОЗВРАЩАЕМ ПЛИТЫ ПЛАНА В ПРОИЗВОДСТВО ===
    # Это нужно сделать ДО удаления файла плана,
    # чтобы плиты не "зависли" в статусе 'в плане'
    db_path = str(PROJECT_ROOT_FOR_IMPORT / "plita.db")
    returned_count = kp_db.return_plan_plates_to_production(plan_id, db_path)
    if returned_count > 0:
        logger.info(f"При удалении плана {plan_id}: возвращено {returned_count} записей плит в производство")
    
    # Удаляем файл плана
    if plan_path.exists():
        try:
            os.remove(plan_path)
        except Exception as e:
            logger.exception(f"Ошибка удаления файла плана {plan_id}: {e}")
            return False
    
    # Обновляем метаданные
    metadata = load_plans_metadata()
    metadata['plans'] = [p for p in metadata['plans'] if p['id'] != plan_id]
    
    # Если удаляли активный план — сбрасываем active_plan_id
    if metadata.get('active_plan_id') == plan_id:
        metadata['active_plan_id'] = None
        # Если есть другие планы — делаем первый активным
        if metadata['plans']:
            metadata['active_plan_id'] = metadata['plans'][0]['id']
    
    save_plans_metadata(metadata)
    logger.info(f"План {plan_id} удалён")
    return True


def get_active_plan_id() -> Optional[str]:
    """
    Возвращает ID активного плана.
    
    Returns:
        str или None: ID активного плана
    """
    metadata = load_plans_metadata()
    return metadata.get('active_plan_id')


def set_active_plan(plan_id: str):
    """
    Устанавливает план как активный.
    
    Args:
        plan_id: ID плана
    """
    metadata = load_plans_metadata()
    metadata['active_plan_id'] = plan_id
    save_plans_metadata(metadata)
    logger.info(f"План {plan_id} установлен как активный")


def get_active_plan() -> Optional[dict]:
    """
    Загружает активный план.
    
    Returns:
        dict или None: Данные активного плана
    """
    active_id = get_active_plan_id()
    if not active_id:
        return None
    return load_plan(active_id)


def get_all_active_plan_ids() -> List[str]:
    """
    Возвращает список ID всех планов, которые сохранены в metadata.
    
    Используется как fallback для заполнения plan_ids при завершении дня,
    если source_plans пуст.
    
    Returns:
        list: Список plan_id из metadata (может быть пуст)
    """
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    if isinstance(plans, dict):
        return list(plans.keys())
    return [p.get('id') for p in plans if isinstance(p, dict) and p.get('id')]


def distribute_tracks_by_days(
    tracks_list: list,
    start_date: str,
    tracks_per_day: int
) -> Dict[str, List]:
    """
    Разбивает список дорожек по дням.
    
    Args:
        tracks_list: Список дорожек для распределения
        start_date: Дата начала в формате 'YYYY-MM-DD'
        tracks_per_day: Сколько дорожек в день
        
    Returns:
        dict: {
            "2026-01-22": [track1, track2, ...],
            "2026-01-23": [track6, track7, ...],
            ...
        }
    """
    result = {}
    
    # Парсим дату начала
    try:
        current_date = datetime.strptime(start_date, '%Y-%m-%d')
    except ValueError:
        current_date = datetime.now()
    
    # Распределяем дорожки по дням
    track_index = 0
    while track_index < len(tracks_list):
        date_key = current_date.strftime('%Y-%m-%d')
        
        # Берём tracks_per_day дорожек для этого дня
        day_tracks = tracks_list[track_index:track_index + tracks_per_day]
        result[date_key] = day_tracks
        
        track_index += tracks_per_day
        current_date += timedelta(days=1)
    
    return result


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
    auto_save: bool = True
) -> Tuple[dict, dict]:
    """
    Добавляет дорожки к существующему плану или создаёт новый.
    
    Логика:
    - Если plan_id указан и план существует — добавляем к нему
    - Иначе создаём новый план
    - Разбивает new_tracks_list по дням
    - Добавляет к существующим дням или создаёт новые
    
    Args:
        plan_id: ID плана (None для создания нового)
        new_tracks_list: Список новых дорожек
        start_date: Дата начала планирования
        tracks_per_day: Дорожек в день
        plate_lookup_exact: Lookup таблица плит
        plate_lookup_by_length: Lookup по длине
        orders_2d: Заказы 2D
        optimization_result: Результат оптимизации
        plan_name: Название плана (для нового)
        auto_save: Если True - сохраняет план автоматически (по умолчанию),
                   если False - только подготавливает план без сохранения
        
    Returns:
        Tuple[dict, dict]: (обновлённый план, статистика изменений)
    """
    # Статистика изменений
    stats = {
        'days_updated': [],  # Список обновлённых дней
        'days_created': [],  # Список созданных дней
        'is_new_plan': False
    }
    
    # Загружаем или создаём план
    plan = None
    if plan_id:
        plan = load_plan(plan_id)
    
    if not plan:
        # Создаём новый план
        plan_id = create_plan_id()
        
        # Формируем название плана
        if not plan_name:
            try:
                start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                plan_name = f"План с {start_dt.strftime('%d.%m.%Y')}"
            except:
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
        stats['is_new_plan'] = True
        logger.info(f"Создан новый план: {plan_id}")
    
    # Распределяем новые дорожки по дням
    tracks_by_day = distribute_tracks_by_days(new_tracks_list, start_date, tracks_per_day)
    
    # Добавляем дорожки к дням
    day_number = len(plan['days']) + 1  # Начинаем с последнего дня + 1
    
    for date_key, day_tracks in tracks_by_day.items():
        if date_key in plan['days']:
            # День уже существует — добавляем дорожки
            existing_day = plan['days'][date_key]
            old_count = existing_day.get('saved_tracks_count', len(existing_day.get('tracks', [])))
            
            # Добавляем новые дорожки
            existing_day['tracks'].extend(day_tracks)
            existing_day['saved_tracks_count'] = len(existing_day['tracks'])
            
            stats['days_updated'].append({
                'date': date_key,
                'old_count': old_count,
                'new_count': existing_day['saved_tracks_count'],
                'total': existing_day['total_tracks_count']
            })
            logger.info(f"День {date_key}: добавлено {len(day_tracks)} дорожек ({old_count} -> {existing_day['saved_tracks_count']})")
        else:
            # Создаём новый день
            plan['days'][date_key] = {
                'date': date_key,
                'day_number': day_number,
                'tracks': day_tracks,
                'saved_tracks_count': len(day_tracks),
                'total_tracks_count': tracks_per_day,
                'completed': False
            }
            
            stats['days_created'].append({
                'date': date_key,
                'count': len(day_tracks),
                'total': tracks_per_day
            })
            logger.info(f"Создан новый день {date_key}: {len(day_tracks)} дорожек")
            day_number += 1
    
    # Обновляем lookup таблицы и другие данные
    # Конвертируем ключи в строки для JSON
    for key, value in plate_lookup_exact.items():
        str_key = str(key)
        if str_key not in plan['plate_lookup_exact']:
            plan['plate_lookup_exact'][str_key] = value
        elif isinstance(value, list):
            # Объединяем списки
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
    
    # Добавляем orders_2d (расширяем список)
    plan['orders_2d'].extend(orders_2d)
    
    # Обновляем optimization_result (заменяем на последний)
    plan['optimization_result'] = optimization_result
    
    # Сохраняем план только если auto_save=True
    if auto_save:
        save_plan(plan)
        
        # Обновляем метаданные
        update_plan_metadata(plan)
        
        # Устанавливаем как активный
        set_active_plan(plan_id)
    
    return plan, stats


def update_plan_metadata(plan: dict):
    """
    Обновляет запись плана в метаданных.
    
    Args:
        plan: Данные плана
    """
    metadata = load_plans_metadata()
    plan_id = plan['id']
    
    # Подсчитываем статистику
    total_days = len(plan.get('days', {}))
    total_tracks = sum(
        day.get('saved_tracks_count', len(day.get('tracks', [])))
        for day in plan.get('days', {}).values()
    )
    
    # Формируем запись метаданных
    plan_meta = {
        'id': plan_id,
        'name': plan.get('name', f'План {plan_id}'),
        'created_at': plan.get('created_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        'start_date': plan.get('start_date', ''),
        'total_days': total_days,
        'tracks_count': plan.get('tracks_count', 5),
        'total_tracks': total_tracks
    }
    
    # Ищем существующую запись
    found = False
    for i, existing in enumerate(metadata['plans']):
        if existing['id'] == plan_id:
            metadata['plans'][i] = plan_meta
            found = True
            break
    
    if not found:
        metadata['plans'].append(plan_meta)
    
    save_plans_metadata(metadata)


def get_plan_days_info(plan: dict) -> Dict[str, dict]:
    """
    Получает информацию о днях плана для отображения в календаре.
    
    Args:
        plan: Данные плана
        
    Returns:
        dict: {
            "2026-01-22": {
                "saved": 3,
                "total": 5,
                "completed": False,
                "day_number": 1
            },
            ...
        }
    """
    result = {}
    
    for date_key, day_data in plan.get('days', {}).items():
        result[date_key] = {
            'saved': day_data.get('saved_tracks_count', len(day_data.get('tracks', []))),
            'total': day_data.get('total_tracks_count', plan.get('tracks_count', 5)),
            'completed': day_data.get('completed', False),
            'day_number': day_data.get('day_number', 1)
        }
    
    return result


def get_plan_days_for_plate(plan_id: str, plate_name_substring: str) -> List[int]:
    """
    Возвращает номера дней плана, в треках которых встречается плита с именем,
    содержащим plate_name_substring (или длина 5.98 и ширина 1.2 для «59,8-12»).
    Нужно для диагностики: «плита X не списалась — она в плане в днях Y, Z».
    """
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
                # Проверка по длине/ширине: 5.98 или 5.99 м и 1.2 м → 59,8-12 / 59,9-12 (списание по допуску)
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
    """
    Возвращает маппинг day_number → date для плана.
    Нужно для диагностики: при подтверждении дня N какую дату мы загружаем и совпадает ли с планом.
    """
    plan = load_plan(plan_id)
    if not plan or not plan.get('days'):
        return {}
    return {
        day_data.get('day_number'): _date
        for _date, day_data in plan.get('days', {}).items()
        if day_data.get('day_number') is not None
    }


def get_all_tracks_from_plan(plan: dict) -> list:
    """
    Собирает все дорожки из плана в один список (для совместимости со старым кодом).
    
    ВАЖНО: Каждая дорожка получает информацию о дне производства для корректной работы диаграммы Ганта!
    
    Args:
        plan: Данные плана
        
    Returns:
        list: Список всех дорожек с добавленным полем 'production_day'
    """
    all_tracks = []
    
    # Сортируем дни по дате
    sorted_days = sorted(plan.get('days', {}).items(), key=lambda x: x[0])
    
    for date_key, day_data in sorted_days:
        day_number = day_data.get('day_number', 1)
        
        # Добавляем каждую дорожку с информацией о дне
        for track in day_data.get('tracks', []):
            # Создаём копию дорожки и добавляем день производства
            track_copy = track.copy() if isinstance(track, dict) else track
            
            # Добавляем поле production_day
            if isinstance(track_copy, dict):
                track_copy['production_day'] = day_number
            
            all_tracks.append(track_copy)
    
    return all_tracks


def mark_day_completed(plan_id: str, date_key: str) -> bool:
    """
    Отмечает день как выполненный.
    
    Args:
        plan_id: ID плана
        date_key: Дата дня в формате 'YYYY-MM-DD'
        
    Returns:
        bool: True если успешно
    """
    plan = load_plan(plan_id)
    if not plan:
        return False
    
    if date_key in plan.get('days', {}):
        plan['days'][date_key]['completed'] = True
        
        # Добавляем в список completed_days (для совместимости)
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
    """
    Получает дорожки для конкретного дня по номеру.
    
    Args:
        plan: Данные плана
        day_number: Номер дня (1, 2, 3, ...)
        
    Returns:
        Tuple[list, str]: (список дорожек, дата дня)
    """
    for date_key, day_data in plan.get('days', {}).items():
        if day_data.get('day_number') == day_number:
            return day_data.get('tracks', []), date_key
    
    return [], ''


def format_plan_stats_message(stats: dict) -> str:
    """
    Форматирует статистику изменений плана для отображения пользователю.
    
    Args:
        stats: Статистика из add_tracks_to_plan
        
    Returns:
        str: Отформатированное сообщение
    """
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


def get_global_day_occupancy(exclude_plan_id: Optional[str] = None) -> Dict[str, int]:
    """
    Подсчитывает занятость дорожек по ВСЕМ планам для каждой даты.
    
    Простыми словами:
    - Загружает все планы из папки
    - Для каждой даты суммирует количество дорожек из всех планов
    - Возвращает словарь {"2026-01-22": 3, "2026-01-23": 5, ...}
    
    Args:
        exclude_plan_id: ID плана, который нужно исключить из подсчёта
                        (полезно при редактировании плана)
    
    Returns:
        dict: Словарь {дата: количество_занятых_дорожек}
    """
    occupancy = {}
    
    # Загружаем метаданные, чтобы получить список планов
    metadata = load_plans_metadata()
    
    for plan_meta in metadata.get('plans', []):
        plan_id = plan_meta.get('id')
        
        # Пропускаем исключённый план
        if plan_id == exclude_plan_id:
            continue
        
        # Загружаем данные плана
        plan = load_plan(plan_id)
        if not plan:
            continue
        
        # Суммируем дорожки по дням
        for date_key, day_data in plan.get('days', {}).items():
            tracks_count = day_data.get('saved_tracks_count', len(day_data.get('tracks', [])))
            
            if date_key in occupancy:
                occupancy[date_key] += tracks_count
            else:
                occupancy[date_key] = tracks_count
    
    return occupancy


def get_free_tracks_for_date(date: str, exclude_plan_id: Optional[str] = None) -> int:
    """
    Возвращает количество свободных дорожек на указанную дату.
    
    Простыми словами:
    - Смотрит сколько дорожек уже занято на эту дату во всех планах
    - Вычитает из максимума (5)
    - Возвращает сколько ещё можно запланировать
    
    Args:
        date: Дата в формате "YYYY-MM-DD"
        exclude_plan_id: ID плана, который не учитывать (при редактировании)
        
    Returns:
        int: Количество свободных дорожек (0 если день полностью занят)
    """
    occupancy = get_global_day_occupancy(exclude_plan_id)
    occupied = occupancy.get(date, 0)
    free = MAX_TRACKS_PER_DAY - occupied
    return max(0, free)


def get_global_days_info(plan: dict) -> Dict[str, dict]:
    """
    Получает информацию о днях с учётом ГЛОБАЛЬНОЙ загруженности.
    
    Простыми словами:
    - Для каждой даты в плане показывает сколько занято ВО ВСЕХ планах
    - Используется для отображения в календаре: "22.01 3/5"
    
    Args:
        plan: Текущий план (для получения дат и статусов completed)
        
    Returns:
        dict: {
            "2026-01-22": {
                "occupied": 3,      # Занято во всех планах
                "max": 5,           # Максимум в день
                "completed": False, # Статус из текущего плана
                "day_number": 1     # Номер дня в текущем плане
            },
            ...
        }
    """
    result = {}
    
    # Получаем глобальную загруженность
    global_occupancy = get_global_day_occupancy()
    
    # Для каждого дня в плане формируем информацию
    for date_key, day_data in plan.get('days', {}).items():
        result[date_key] = {
            'occupied': global_occupancy.get(date_key, 0),
            'max': MAX_TRACKS_PER_DAY,
            'completed': day_data.get('completed', False),
            'day_number': day_data.get('day_number', 1)
        }
    
    return result


def get_all_plans_gantt_data() -> Optional[dict]:
    """
    Собирает данные из ВСЕХ сохранённых планов для создания суммарной диаграммы Ганта.
    
    Простыми словами:
    - Загружает все планы из папки bot/data/plans/
    - Собирает все дорожки из всех планов в один большой список
    - Объединяет информацию о плитах (plate_lookup) из всех планов
    - Находит самую раннюю и позднюю дату среди всех планов
    
    Returns:
        dict или None: Словарь с данными для диаграммы Ганта:
        {
            'all_tracks': [...],  # Все дорожки из всех планов
            'plate_lookup_exact': {...},  # Объединённый словарь поиска плит
            'plate_lookup_by_length': {...},  # Объединённый словарь поиска по длине
            'earliest_start_date': datetime,  # Самая ранняя дата начала
            'latest_end_date': datetime,  # Самая поздняя дата окончания
            'plans_count': int,  # Количество планов
            'total_days': int  # Общее количество уникальных дней
        }
        Возвращает None если нет ни одного плана
    """
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    
    if not plans:
        logger.warning("[GANTT] Нет сохранённых планов для диаграммы")
        return None
    
    # Инициализируем контейнеры для объединения данных
    all_tracks_combined = []
    combined_plate_lookup_exact = {}
    combined_plate_lookup_by_length = {}
    
    earliest_date = None
    latest_date = None
    unique_dates = set()
    
    plans_loaded = 0
    
    # Проходим по всем планам
    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)
        
        if not plan:
            logger.warning(f"[GANTT] Не удалось загрузить план {plan_id}")
            continue
        
        plans_loaded += 1
        
        # === 1. Собираем дорожки из плана ===
        plan_tracks = get_all_tracks_from_plan(plan)
        all_tracks_combined.extend(plan_tracks)
        
        # === 2. Объединяем plate_lookup_exact ===
        plan_lookup_exact = convert_lookup_keys_to_tuples(plan.get('plate_lookup_exact', {}))
        for key, entries in plan_lookup_exact.items():
            if key not in combined_plate_lookup_exact:
                combined_plate_lookup_exact[key] = []
            
            # Добавляем записи, избегая дублирования
            for entry in entries:
                # Проверяем, нет ли уже такой записи
                if entry not in combined_plate_lookup_exact[key]:
                    combined_plate_lookup_exact[key].append(entry)
        
        # === 3. Объединяем plate_lookup_by_length ===
        plan_lookup_by_length = plan.get('plate_lookup_by_length', {})
        for length_key, entries in plan_lookup_by_length.items():
            # Аналогично преобразуем ключ
            if isinstance(length_key, str):
                try:
                    length_key = float(length_key)
                except:
                    pass
            
            if length_key not in combined_plate_lookup_by_length:
                combined_plate_lookup_by_length[length_key] = []
            
            for entry in entries:
                if entry not in combined_plate_lookup_by_length[length_key]:
                    combined_plate_lookup_by_length[length_key].append(entry)
        
        # === 4. Находим диапазон дат ===
        plan_start = plan.get('start_date')
        if plan_start:
            try:
                start_dt = datetime.strptime(plan_start, '%Y-%m-%d')
                if earliest_date is None or start_dt < earliest_date:
                    earliest_date = start_dt
            except ValueError:
                logger.warning(f"[GANTT] Неверный формат даты начала в плане {plan_id}: {plan_start}")
        
        # Проходим по всем дням плана
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
    
    # Если даты не найдены, используем текущую дату
    if earliest_date is None:
        earliest_date = datetime.now()
    if latest_date is None:
        latest_date = datetime.now()
    
    logger.info(f"[GANTT] Собрано данных для диаграммы: {plans_loaded} планов, "
                f"{len(all_tracks_combined)} дорожек, период {earliest_date.strftime('%d.%m.%Y')} - "
                f"{latest_date.strftime('%d.%m.%Y')}")
    
    return {
        'all_tracks': all_tracks_combined,
        'plate_lookup_exact': combined_plate_lookup_exact,
        'plate_lookup_by_length': combined_plate_lookup_by_length,
        'earliest_start_date': earliest_date,
        'latest_end_date': latest_date,
        'plans_count': plans_loaded,
        'total_days': len(unique_dates)
    }


def get_global_calendar_info() -> Optional[dict]:
    """
    Собирает информацию о ВСЕХ днях из ВСЕХ планов для отображения единого календаря.
    
    Простыми словами:
    - Загружает все планы из папки bot/data/plans/
    - Находит самую раннюю и самую позднюю дату среди всех планов
    - Для каждой даты считает общую загрузку дорожек
    - Определяет статус выполнения (день считается выполненным, если хотя бы в одном плане он отмечен как completed)
    
    Returns:
        dict или None: Словарь с данными для календаря:
        {
            'start_date': str,  # Самая ранняя дата в формате 'YYYY-MM-DD'
            'total_days': int,  # Количество дней от начала до конца
            'days_info': {  # Информация по каждой дате
                "2026-01-22": {
                    "occupied": 3,      # Занято дорожек во всех планах
                    "max": 5,           # Максимум дорожек в день
                    "completed": False, # Выполнен ли день
                    "day_number": 1     # Порядковый номер дня от start_date
                },
                ...
            },
            'completed_days': [1, 2, ...],  # Список номеров выполненных дней
            'plans_count': int,  # Количество планов
            'tracks_count': int  # Суммарное количество дорожек
        }
        Возвращает None если нет ни одного плана
    """
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    
    if not plans:
        logger.warning("[GLOBAL_CALENDAR] Нет сохранённых планов")
        return None
    
    # Собираем информацию о датах и загрузке
    all_dates_data = {}  # {date_key: {'occupied': int, 'completed': bool}}
    earliest_date = None
    latest_date = None
    total_tracks_count = 0
    
    # Проходим по всем планам
    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)
        
        if not plan:
            logger.warning(f"[GLOBAL_CALENDAR] Не удалось загрузить план {plan_id}")
            continue
        
        # Обрабатываем каждый день плана
        for date_key, day_data in plan.get('days', {}).items():
            # Парсим дату для определения диапазона
            try:
                day_dt = datetime.strptime(date_key, '%Y-%m-%d')
                if earliest_date is None or day_dt < earliest_date:
                    earliest_date = day_dt
                if latest_date is None or day_dt > latest_date:
                    latest_date = day_dt
            except ValueError:
                logger.warning(f"[GLOBAL_CALENDAR] Неверный формат даты: {date_key}")
                continue
            
            # Считаем загрузку
            tracks_count = day_data.get('saved_tracks_count', len(day_data.get('tracks', [])))
            is_completed = day_data.get('completed', False)
            
            # Добавляем или обновляем информацию о дате
            if date_key not in all_dates_data:
                all_dates_data[date_key] = {
                    'occupied': 0,
                    'completed': False
                }
            
            all_dates_data[date_key]['occupied'] += tracks_count
            # День считается выполненным, если он выполнен хотя бы в одном плане
            if is_completed:
                all_dates_data[date_key]['completed'] = True
            
            total_tracks_count += tracks_count
    
    # Если не удалось найти даты
    if earliest_date is None or latest_date is None:
        logger.warning("[GLOBAL_CALENDAR] Не удалось определить диапазон дат")
        return None
    
    # Вычисляем общее количество дней (от earliest до latest включительно)
    total_days = (latest_date - earliest_date).days + 1
    
    # Формируем days_info с информацией для каждого дня
    days_info = {}
    completed_days = []
    
    for day_offset in range(total_days):
        current_date = earliest_date + timedelta(days=day_offset)
        date_key = current_date.strftime('%Y-%m-%d')
        day_number = day_offset + 1
        
        # Берем данные из all_dates_data или ставим 0 если дня нет ни в одном плане
        date_data = all_dates_data.get(date_key, {'occupied': 0, 'completed': False})
        
        days_info[date_key] = {
            'occupied': date_data['occupied'],
            'max': MAX_TRACKS_PER_DAY,
            'completed': date_data['completed'],
            'day_number': day_number
        }
        
        # Добавляем в список выполненных дней
        if date_data['completed']:
            completed_days.append(day_number)
    
    logger.info(f"[GLOBAL_CALENDAR] Создан глобальный календарь: {len(plans)} планов, "
                f"{total_days} дней ({earliest_date.strftime('%d.%m.%Y')} - {latest_date.strftime('%d.%m.%Y')}), "
                f"{total_tracks_count} дорожек")
    
    return {
        'start_date': earliest_date.strftime('%Y-%m-%d'),
        'total_days': total_days,
        'days_info': days_info,
        'completed_days': completed_days,
        'plans_count': len(plans),
        'tracks_count': total_tracks_count
    }


def get_tracks_for_date_from_all_plans(date_key: str) -> Optional[dict]:
    """
    Собирает дорожки на конкретную дату из ВСЕХ сохранённых планов.
    
    Простыми словами:
    - Загружает все планы из папки bot/data/plans/
    - Для каждого плана ищет указанную дату
    - Собирает все дорожки с этой даты из всех планов
    - Объединяет lookup-таблицы для корректной работы генераторов документов
    
    Args:
        date_key: Дата в формате 'YYYY-MM-DD' (например, '2026-01-29')
    
    Returns:
        dict или None: Словарь с данными для дня:
        {
            'tracks': [...],  # Все дорожки на эту дату из всех планов
            'plate_lookup_exact': {...},  # Объединённые lookup
            'plate_lookup_by_length': {...},
            'orders_2d': [...],  # Все заказы из всех планов
            'optimization_result': {...},  # Результат оптимизации (последний)
            'plans_count': int,  # Количество планов, содержащих эту дату
            'source_plans': [...]  # Список ID планов-источников
        }
        Возвращает None если дата не найдена ни в одном плане
    """
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    
    if not plans:
        logger.warning(f"[MULTI_PLAN] Нет сохранённых планов для даты {date_key}")
        return None
    
    # Инициализируем контейнеры для объединения данных
    all_tracks_for_date = []
    combined_plate_lookup_exact = {}
    combined_plate_lookup_by_length = {}
    combined_orders_2d = []
    last_optimization_result = {}
    
    source_plan_ids = []
    plans_with_date = 0
    
    # Проходим по всем планам
    for plan_meta in plans:
        plan_id = plan_meta.get('id')
        plan = load_plan(plan_id)
        
        if not plan:
            logger.warning(f"[MULTI_PLAN] Не удалось загрузить план {plan_id}")
            continue
        
        # Проверяем, есть ли эта дата в плане
        if date_key not in plan.get('days', {}):
            continue
        
        plans_with_date += 1
        source_plan_ids.append(plan_id)
        
        # === 1. Собираем дорожки на эту дату ===
        day_data = plan['days'][date_key]
        day_tracks = day_data.get('tracks', [])
        
        # НОВОЕ: Добавляем информацию о плане-источнике к каждой дорожке
        plan_name = plan.get('name', f'План {plan_id}')
        for track in day_tracks:
            if isinstance(track, dict):
                track['source_plan_id'] = plan_id
                track['source_plan_name'] = plan_name
        
        all_tracks_for_date.extend(day_tracks)
        
        logger.info(f"[MULTI_PLAN] План {plan_id}: найдено {len(day_tracks)} дорожек на {date_key}")
        
        # === 2. Объединяем plate_lookup_exact ===
        plan_lookup_exact = convert_lookup_keys_to_tuples(plan.get('plate_lookup_exact', {}))
        for key, entries in plan_lookup_exact.items():
            if key not in combined_plate_lookup_exact:
                combined_plate_lookup_exact[key] = []
            
            # Добавляем записи, избегая дублирования
            for entry in entries:
                # Проверяем, нет ли уже такой записи (по kp_id и plate_name)
                entry_exists = False
                for existing in combined_plate_lookup_exact[key]:
                    if (existing.get('kp_id') == entry.get('kp_id') and
                        existing.get('plate_name') == entry.get('plate_name')):
                        entry_exists = True
                        break
                
                if not entry_exists:
                    combined_plate_lookup_exact[key].append(entry)
        
        # === 3. Объединяем plate_lookup_by_length ===
        plan_lookup_by_length = plan.get('plate_lookup_by_length', {})
        for length_key, entries in plan_lookup_by_length.items():
            # Преобразуем ключ
            if isinstance(length_key, str):
                try:
                    length_key = float(length_key)
                except:
                    pass
            
            if length_key not in combined_plate_lookup_by_length:
                combined_plate_lookup_by_length[length_key] = []
            
            for entry in entries:
                # Проверяем дубликаты
                entry_exists = False
                for existing in combined_plate_lookup_by_length[length_key]:
                    if (existing.get('kp_id') == entry.get('kp_id') and
                        existing.get('plate_name') == entry.get('plate_name')):
                        entry_exists = True
                        break
                
                if not entry_exists:
                    combined_plate_lookup_by_length[length_key].append(entry)
        
        # === 4. Собираем orders_2d ===
        plan_orders = plan.get('orders_2d', [])
        combined_orders_2d.extend(plan_orders)
        
        # === 5. Сохраняем последний результат оптимизации ===
        plan_opt_result = plan.get('optimization_result', {})
        if plan_opt_result:
            last_optimization_result = plan_opt_result
    
    # Если дата не найдена ни в одном плане
    if plans_with_date == 0:
        logger.warning(f"[MULTI_PLAN] Дата {date_key} не найдена ни в одном плане")
        return None
    
    logger.info(f"[MULTI_PLAN] Собрано данных для {date_key}: "
                f"{plans_with_date} планов, {len(all_tracks_for_date)} дорожек")
    
    return {
        'tracks': all_tracks_for_date,
        'plate_lookup_exact': combined_plate_lookup_exact,
        'plate_lookup_by_length': combined_plate_lookup_by_length,
        'orders_2d': combined_orders_2d,
        'optimization_result': last_optimization_result,
        'plans_count': plans_with_date,
        'source_plans': source_plan_ids
    }
