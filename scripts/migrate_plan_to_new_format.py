"""
Скрипт миграции старого current_plan.json в новый формат системы планов.

Что делает:
1. Читает старый current_plan.json
2. Конвертирует в новый формат с разбивкой по дням
3. Создаёт plans_metadata.json
4. Сохраняет в plans/plan_migrated.json

Использование:
    python scripts/migrate_plan_to_new_format.py
"""
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Добавляем корень проекта в путь
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

BOT_DIR = PROJECT_ROOT / "bot"
OLD_PLAN_PATH = BOT_DIR / "data" / "current_plan.json"
PLANS_DIR = BOT_DIR / "data" / "plans"
PLANS_METADATA_PATH = BOT_DIR / "data" / "plans_metadata.json"


def migrate_old_plan():
    """
    Конвертирует старый current_plan.json в новый формат.
    """
    print("=" * 50)
    print("Migrating plan to new format")
    print("=" * 50)
    
    # Проверяем существование старого плана
    if not OLD_PLAN_PATH.exists():
        print(f"[ERROR] Old plan not found: {OLD_PLAN_PATH}")
        print("   Migration not required.")
        return False
    
    # Загружаем старый план
    print(f"[INFO] Loading old plan: {OLD_PLAN_PATH}")
    try:
        with open(OLD_PLAN_PATH, 'r', encoding='utf-8') as f:
            old_plan = json.load(f)
    except Exception as e:
        print(f"[ERROR] Failed to read old plan: {e}")
        return False
    
    # Извлекаем данные из старого формата
    all_tracks_list = old_plan.get('all_tracks_list', [])
    tracks_count = old_plan.get('tracks_count', 5)
    start_date = old_plan.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = old_plan.get('completed_days', [])
    total_days = old_plan.get('total_days', 0)
    
    print(f"[INFO] Old plan data:")
    print(f"   - Tracks: {len(all_tracks_list)}")
    print(f"   - Days: {total_days}")
    print(f"   - Tracks per day: {tracks_count}")
    print(f"   - Completed days: {len(completed_days)}")
    
    if not all_tracks_list:
        print("[ERROR] No tracks in old plan. Migration impossible.")
        return False
    
    # Создаём папку для планов
    PLANS_DIR.mkdir(parents=True, exist_ok=True)
    
    # Генерируем ID для нового плана
    plan_id = f"plan_migrated_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Парсим дату начала
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
    except ValueError:
        start_dt = datetime.now()
    
    # Формируем название плана
    plan_name = f"Plan from {start_dt.strftime('%d.%m.%Y')} (migrated)"
    
    # Создаём новую структуру плана
    new_plan = {
        'id': plan_id,
        'name': plan_name,
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'start_date': start_date,
        'tracks_count': tracks_count,
        'days': {},
        'plate_lookup_exact': old_plan.get('plate_lookup_exact', {}),
        'plate_lookup_by_length': old_plan.get('plate_lookup_by_length', {}),
        'orders_2d': old_plan.get('orders_2d', []),
        'optimization_result': old_plan.get('optimization_result', {}),
        'completed_days': completed_days
    }
    
    # Распределяем дорожки по дням
    print(f"\n[INFO] Distributing tracks by days...")
    
    current_date = start_dt
    track_index = 0
    day_number = 1
    
    while track_index < len(all_tracks_list):
        date_key = current_date.strftime('%Y-%m-%d')
        
        # Берём дорожки для этого дня
        day_tracks = all_tracks_list[track_index:track_index + tracks_count]
        
        # Определяем, выполнен ли день
        is_completed = day_number in completed_days
        
        # Создаём запись для дня
        new_plan['days'][date_key] = {
            'date': date_key,
            'day_number': day_number,
            'tracks': day_tracks,
            'saved_tracks_count': len(day_tracks),
            'total_tracks_count': tracks_count,
            'completed': is_completed
        }
        
        completed_mark = "[DONE]" if is_completed else ""
        print(f"   Day {day_number} ({date_key}): {len(day_tracks)} tracks {completed_mark}")
        
        track_index += tracks_count
        current_date += timedelta(days=1)
        day_number += 1
    
    # Сохраняем новый план
    new_plan_path = PLANS_DIR / f"{plan_id}.json"
    print(f"\n[INFO] Saving new plan: {new_plan_path}")
    
    with open(new_plan_path, 'w', encoding='utf-8') as f:
        json.dump(new_plan, f, ensure_ascii=False, indent=2)
    
    # Создаём или обновляем метаданные
    print(f"[INFO] Creating plans metadata...")
    
    metadata = {
        'plans': [
            {
                'id': plan_id,
                'name': plan_name,
                'created_at': new_plan['created_at'],
                'start_date': start_date,
                'total_days': len(new_plan['days']),
                'tracks_count': tracks_count,
                'total_tracks': len(all_tracks_list)
            }
        ],
        'active_plan_id': plan_id
    }
    
    with open(PLANS_METADATA_PATH, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    
    print(f"\n[SUCCESS] Migration completed!")
    print(f"   - New plan: {new_plan_path}")
    print(f"   - Metadata: {PLANS_METADATA_PATH}")
    print(f"   - Plan ID: {plan_id}")
    print(f"   - Days in plan: {len(new_plan['days'])}")
    
    # Предлагаем удалить старый план
    print(f"\n[WARNING] Old plan still exists: {OLD_PLAN_PATH}")
    print(f"   You can delete it manually after verification.")
    
    return True


def main():
    """Точка входа."""
    success = migrate_old_plan()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
