"""Завершение производственного дня с учётом брака"""
import copy
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
import core.config_and_data as cfg

from ..keyboards import main_menu_kb, calendar_days_kb, plates_completion_kb
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    get_active_plan_id, mark_day_completed, get_tracks_for_date_from_all_plans,
    get_all_active_plan_ids
)

router = Router()


@router.callback_query(F.data.startswith("complete_day_"))
async def start_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Начало процесса завершения дня.
    Показываем плиты для отметки брака.
    """
    day_number = int(callback.data.split("_")[-1])
    data = await state.get_data()
    
    # НОВОЕ: Проверяем данные из сохранённого плана (current_day_tracks)
    current_day_tracks = data.get('current_day_tracks')
    from_saved_plan = data.get('from_saved_plan', False)
    
    if current_day_tracks:
        # Работаем с сохранённым планом - используем предзагруженные дорожки
        tracks_for_day = current_day_tracks
        plate_lookup_exact = data.get('current_day_plate_lookup_exact', {})
        plate_lookup_by_length = data.get('current_day_plate_lookup_by_length', {})
        start_index = data.get('current_day_start_index', 0)
        source_plans = data.get('current_day_source_plans', [])
        
        logger.info(f"[COMPLETION] День {day_number}: используем предзагруженные дорожки ({len(tracks_for_day)} шт)")
    elif from_saved_plan:
        # МУЛЬТИПЛАН: Работаем с сохранённым планом, но данные не загружены
        # Загружаем дорожки из ВСЕХ планов на эту дату
        plan_start_date = data.get('plan_start_date')
        
        if not plan_start_date:
            await callback.message.answer(
                "❌ Не удалось определить дату начала плана.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        # Вычисляем дату выбранного дня
        try:
            start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
            selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
        except ValueError as e:
            logger.error(f"[COMPLETION] Ошибка парсинга даты: {e}")
            await callback.message.answer(
                "❌ Ошибка обработки даты.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        # Загружаем дорожки из всех планов на эту дату
        logger.info(f"[COMPLETION] День {day_number} ({selected_date}): загружаем из всех планов...")
        multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
        
        if not multi_plan_data or not multi_plan_data.get('tracks'):
            await callback.message.answer(
                f"❌ Дата {datetime.strptime(selected_date, '%Y-%m-%d').strftime('%d.%m.%Y')} "
                f"не найдена ни в одном сохранённом плане.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        tracks_for_day = multi_plan_data['tracks']
        plate_lookup_exact = multi_plan_data['plate_lookup_exact']
        plate_lookup_by_length = multi_plan_data['plate_lookup_by_length']
        start_index = 0
        source_plans = multi_plan_data.get('source_plans', [])
        
        logger.info(f"[COMPLETION] День {day_number}: загружено {len(tracks_for_day)} дорожек из {multi_plan_data['plans_count']} планов")
        
        # Сохраняем source_plans для последующей отметки завершения во всех планах
        await state.update_data(current_day_source_plans=source_plans)
    else:
        # СТАРАЯ ЛОГИКА: Работаем с новым планом - извлекаем из all_tracks_list
        tracks_count = data.get('tracks_count', 1)
        all_tracks_list = data.get('all_tracks_list', [])
        plate_lookup_exact = data.get('plate_lookup_exact', {})
        plate_lookup_by_length = data.get('plate_lookup_by_length', {})
        source_plans = []
        
        if not all_tracks_list:
            await callback.message.answer(
                "❌ Данные о дорожках не найдены. Попробуйте заново.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        # Собираем все плиты дня
        start_index = (day_number - 1) * tracks_count
        end_index = min(day_number * tracks_count, len(all_tracks_list))
        tracks_for_day = all_tracks_list[start_index:end_index]
        
        logger.info(f"[COMPLETION] День {day_number}: используем данные из нового плана ({len(tracks_for_day)} дорожек)")
    
    # Создаем КОПИЮ lookup для завершения дня (чтобы не влиять на оригинал в state)
    completion_lookup_exact = copy.deepcopy(plate_lookup_exact)
    completion_lookup_by_length = copy.deepcopy(plate_lookup_by_length)
    
    # Функция для получения информации о плите (с списанием из lookup)
    def get_plate_info_smart(length, width, expected_kp_id=None, load_code=None):
        """
        Умный поиск информации о плите С УЧЕТОМ КОЛИЧЕСТВА.
        
        Логика:
        1. Ищем в списке записей по (length, width)
        2. При одном ключе могут быть 8п/10п/12,5п — предпочитаем запись с совпадающим load_code
        3. Уменьшаем qty_remaining на 1 (списываем плиту)
        4. Возвращаем информацию о КП
        
        ВАЖНО: Работаем с КОПИЕЙ lookup, чтобы не влиять на оригинал.
        
        FUZZY-ПОИСК: Если точный ключ не найден, ищем с tolerance 0.03м (30мм) по длине.
        """
        TOLERANCE = 0.03  # 30мм tolerance для fuzzy-поиска
        rounded_length = round(length, 2)
        if load_code is None:
            load_code = 8
        load_code = cfg.normalize_load_code(load_code)

        def load_code_matches(entry):
            entry_lc = entry.get('load_code')
            if entry_lc is None:
                entry_lc = 8
            return abs(cfg.normalize_load_code(entry_lc) - load_code) < 0.01

        def pick_best(entries):
            """Выбрать запись: приоритет (kp_id + load_code) -> load_code -> kp_id -> любая."""
            for entry in entries:
                if entry.get('qty_remaining', 0) <= 0:
                    continue
                if expected_kp_id and entry.get('kp_id') == expected_kp_id and load_code_matches(entry):
                    return entry
            for entry in entries:
                if entry.get('qty_remaining', 0) <= 0:
                    continue
                if load_code_matches(entry):
                    return entry
            for entry in entries:
                if entry.get('qty_remaining', 0) <= 0:
                    continue
                if expected_kp_id and entry.get('kp_id') == expected_kp_id:
                    return entry
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    return entry
            return None

        logger.debug(f"[TRACE] Ищем плиту: длина={length:.2f}м ({rounded_length:.2f}м), ширина={width}мм, load_code={load_code}")
        
        # 1. Точное совпадение по (length, width)
        key = (rounded_length, width)
        entries = completion_lookup_exact.get(key, [])
        if entries:
            logger.debug(f"[TRACE]   Найдено {len(entries)} записей для ключа {key}")
        best = pick_best(entries)
        if best:
            best['qty_remaining'] -= 1
            logger.debug(f"[TRACE]   ✓ Найдено ТОЧНОЕ совпадение, КП #{best.get('kp_id', '?')}")
            return best.copy()
        
        # 2. Fuzzy по длине и ширине в exact lookup
        logger.debug(f"[TRACE]   Пробуем fuzzy-поиск (tolerance={TOLERANCE}м)")
        best_match = None
        for lookup_key, entries in completion_lookup_exact.items():
            if not isinstance(lookup_key, (tuple, list)) or len(lookup_key) < 2:
                continue
            key_length = float(lookup_key[0]) if lookup_key[0] is not None else 0
            key_width = int(lookup_key[1]) if isinstance(lookup_key[1], (int, float)) else width
            if abs(key_width - width) > 20:
                continue
            if abs(key_length - rounded_length) <= TOLERANCE:
                best_match = pick_best(entries)
                if best_match:
                    break
        if best_match:
            best_match['qty_remaining'] -= 1
            logger.debug(f"[TRACE]   ✓ Найдено FUZZY совпадение, КП #{best_match.get('kp_id', '?')}")
            return best_match.copy()
        
        # 3. По длине (точный)
        entries = completion_lookup_by_length.get(rounded_length, [])
        best = pick_best(entries)
        if best:
            best['qty_remaining'] -= 1
            logger.debug(f"[TRACE]   ✓ Найдено по длине, КП #{best.get('kp_id', '?')}")
            return best.copy()
        
        # 4. Fuzzy по длине в by_length
        best_match = None
        for lookup_length, entries in completion_lookup_by_length.items():
            ll = float(lookup_length) if isinstance(lookup_length, str) else lookup_length
            if abs(ll - rounded_length) <= TOLERANCE:
                best_match = pick_best(entries)
                if best_match:
                    break
        if best_match:
            best_match['qty_remaining'] -= 1
            logger.debug(f"[TRACE]   ✓ Найдено FUZZY по длине, КП #{best_match.get('kp_id', '?')}")
            return best_match.copy()
        
        logger.warning(f"[TRACE]   ❌ НЕ НАЙДЕНО совпадение для: длина={length:.2f}м, ширина={width}мм, load_code={load_code}")
        return {
            'kp_date': 'неизвестно',
            'customer': 'неизвестно',
            'plate_name': '',
            'kp_id': None
        }
    
    # ✅ НОВОЕ: Логируем плиты ДО обработки для списания
    logger.info(f"[TRACE] ===== ШАГ 7: ПЛИТЫ ПЕРЕД СПИСАНИЕМ (День {day_number}) =====")
    logger.info(f"[TRACE] Дорожек для этого дня: {len(tracks_for_day)}")
    
    total_items_before = 0
    for track_idx, track in enumerate(tracks_for_day):
        items_count = len(track.get('items', []))
        total_items_before += items_count
        logger.info(f"[TRACE]   Дорожка #{start_index + track_idx + 1}: {items_count} плит")
    
    logger.info(f"[TRACE] Всего плит в дорожках: {total_items_before}")
    
    # Собираем плиты по дорожкам (каждая дорожка отдельно)
    day_plates_by_track = []
    total_qty = 0
    
    for track_idx, track in enumerate(tracks_for_day):
        track_number = start_index + track_idx + 1
        track_plates = []
        # #region agent log H8: items в track
        _track_items = track.get('items', [])
        _split_items = [i for i in _track_items if i and i.get('mode') == 'split' and i.get('secondary_cuts')]
        if _split_items and track_idx == 0:  # Только первая дорожка для краткости
            _d8 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
            import json as _j8
            with open(_d8, 'a', encoding='utf-8') as _f8:
                _f8.write(_j8.dumps({"hypothesisId": "H8", "location": "production_completion:track_items", "message": "Items с secondary_cuts в track", "data": {"track_idx": track_idx, "items_count": len(_track_items), "split_with_sec_count": len(_split_items), "day": day_number}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        # #endregion
        for item in track.get('items', []):
            if item is None:
                continue
            length = item.get('length', 0)
            
            # Определяем ширину в зависимости от режима плиты
            mode = item.get('mode', 'solid')
            if mode == 'transverse' and item.get('width'):
                width = round(item['width'] * 1000)  # round для корректного округления float
            elif mode == 'split' and item.get('main_w'):
                width = round(item['main_w'] * 1000)  # round для корректного округления
            else:
                # solid mode - берём width из поля (ИСПРАВЛЕНИЕ: учитываем реальную ширину)
                width = round(item.get('width', 1.2) * 1000)  # round для корректного округления
            
            if not length:
                continue
            
            item_kp_id = item.get('kp_id')
            item_plate_name = item.get('plate_name')
            item_load_code = cfg.normalize_load_code(item.get('load_code', 8))

            plate_info = get_plate_info_smart(length, width, expected_kp_id=item_kp_id, load_code=item_load_code)
            plate_name = plate_info.get('plate_name', '') or (item_plate_name or '')
            kp_id = plate_info.get('kp_id') or item_kp_id
            
            # Если нет имени плиты — формируем его
            if not plate_name and kp_id:
                # Пытаемся взять точное имя из БД по kp_id+размерам
                try:
                    conn = None
                    db_path = str(PROJECT_ROOT / "plita.db")
                    conn = sqlite3.connect(db_path)
                    cur = conn.cursor()
                    # Сначала пробуем точный поиск по длине и ширине
                    cur.execute('''
                        SELECT plate_name, load_class, width_m
                        FROM kp_plates
                        WHERE kp_id = ?
                          AND status = 'в плане'
                          AND ABS(length_m - ?) < 0.02
                          AND ABS(width_m - ?) < 0.01
                        LIMIT 1
                    ''', (kp_id, length, width / 1000.0))
                    row = cur.fetchone()
                    
                    # Если не нашли — ищем ТОЛЬКО по длине (для вторичных резов с неправильной шириной)
                    if not row:
                        cur.execute('''
                            SELECT plate_name, load_class, width_m
                            FROM kp_plates
                            WHERE kp_id = ?
                              AND status = 'в плане'
                              AND ABS(length_m - ?) < 0.02
                            LIMIT 1
                        ''', (kp_id, length))
                        row = cur.fetchone()
                    
                    if row:
                        plate_name = row[0]
                        item_load_code = cfg.normalize_load_code((row[1] or 800) / 100)
                        # Обновляем width на правильное значение из БД
                        width = row[2] * 1000  # width_m → мм
                finally:
                    try:
                        conn.close()
                    except Exception:
                        pass

            if not plate_name:
                plate_name = cfg.make_plate_name(length, width / 1000.0, load_code=item_load_code)
            
            # Получаем дату и заказчика для группировки
            kp_date = plate_info.get('kp_date', 'неизвестно') or item.get('kp_date', 'неизвестно')
            customer = plate_info.get('customer', 'неизвестно') or item.get('customer', 'неизвестно')
            
            # Ищем такую же плиту в списке текущей дорожки
            # Группируем по: plate_name + kp_id + kp_date + customer + width
            found = False
            width_m = width / 1000.0
            for existing in track_plates:
                if (existing['plate_name'] == plate_name and 
                    existing['kp_id'] == kp_id and
                    existing['kp_date'] == kp_date and
                    existing['customer'] == customer and
                    abs(existing['width_m'] - width_m) < 0.01 and
                    not existing.get('is_secondary', False)):  # Только с другими основными!
                    existing['qty'] += 1
                    found = True
                    break
            
            if not found:
                track_plates.append({
                    'plate_name': plate_name,
                    'length_m': length,
                    'width_m': width_m,
                    'load_class': int(cfg.normalize_load_code(item_load_code) * 100),
                    'qty': 1,
                    'kp_id': kp_id,
                    'kp_date': kp_date,
                    'customer': customer,
                    'is_secondary': False  # Флаг: это основная плита
                })
            
            # НОВОЕ: Обрабатываем плиты из вторичных резов (остатков)
            secondary_cuts = item.get('secondary_cuts', []) if item else []
            # #region agent log H8: secondary_cuts в item
            if secondary_cuts:
                _d8 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
                import json as _j8
                with open(_d8, 'a', encoding='utf-8') as _f8:
                    _f8.write(_j8.dumps({"hypothesisId": "H8", "location": "production_completion:secondary_cuts_check", "message": "secondary_cuts в item", "data": {"item_plate_name": item.get('plate_name', ''), "item_mode": item.get('mode', ''), "sec_cuts_count": len(secondary_cuts), "sec_cuts_sample": str(secondary_cuts)[:300], "day": day_number}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            # #endregion
            # #region agent log H12: RELOAD_CHECK — маркер что новый код загружен
            try:
                with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f12:
                    import json as _j12
                    _f12.write(_j12.dumps({"hypothesisId": "H12", "location": "production_completion:before_sec_loop", "message": "RELOAD_CHECK_2026_02_06_v2", "data": {"secondary_cuts_len": len(secondary_cuts) if secondary_cuts else 0, "secondary_cuts_type": str(type(secondary_cuts).__name__), "secondary_cuts_truthy": bool(secondary_cuts), "item_plate_name": item.get('plate_name', ''), "day": day_number}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            except Exception:
                pass
            # #endregion
            for sec_cut in (secondary_cuts or []):
                # #region agent log H11: вход в цикл secondary_cut (ПЕРЕД sec_width_m)
                try:
                    with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f11:
                        import json as _j11
                        _f11.write(_j11.dumps({"hypothesisId": "H11", "location": "production_completion:sec_cut_loop_entry", "message": "Вход в цикл sec_cut", "data": {"sec_cut_keys": list(sec_cut.keys()) if isinstance(sec_cut, dict) else str(type(sec_cut)), "sec_cut_repr": str(sec_cut)[:200], "day": day_number}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                except Exception:
                    pass
                # #endregion
                sec_width_m = sec_cut.get('width', 0)
                if sec_width_m <= 0:
                    continue
                
                sec_width = round(sec_width_m * 1000)  # round для корректного округления float
                # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                sec_length = sec_cut.get('target_length') or length
                
                # ИСПРАВЛЕНИЕ: Для вторичных резов ищем по ВСЕМ КП, без ограничения
                # Сначала пробуем найти с expected_kp_id (если родительский КП совпадает)
                try:
                    sec_plate_info = get_plate_info_smart(sec_length, sec_width, expected_kp_id=item_kp_id, load_code=item_load_code)
                except Exception as _exc11:
                    with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f11e:
                        import json as _j11e
                        _f11e.write(_j11e.dumps({"hypothesisId": "H11_ERR", "location": "production_completion:get_plate_info_smart_error", "message": str(_exc11), "data": {"sec_length": sec_length, "sec_width": sec_width, "day": day_number, "label": sec_cut.get('label', '')}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                    sec_plate_info = {'kp_date': 'неизвестно', 'customer': 'неизвестно', 'plate_name': '', 'kp_id': None}
                sec_kp_id = sec_plate_info.get('kp_id')
                
                # Если не нашли с ограничением — ищем по ВСЕМ КП
                if not sec_kp_id:
                    sec_plate_info = get_plate_info_smart(sec_length, sec_width, expected_kp_id=None, load_code=item_load_code)
                    sec_kp_id = sec_plate_info.get('kp_id')
                
                sec_plate_name = sec_plate_info.get('plate_name', '') or (sec_cut.get('label', '') or '').replace('О ', '').strip()
                # Fallback на item_kp_id только если вообще ничего не нашли
                if not sec_kp_id:
                    sec_kp_id = item_kp_id
                
                # Если нет имени плиты — берём из label
                if not sec_plate_name:
                    sec_plate_name = cfg.make_plate_name(sec_length, sec_width / 1000.0, load_code=item_load_code)
                
                sec_kp_date = sec_plate_info.get('kp_date', 'неизвестно') or item.get('kp_date', 'неизвестно')
                sec_customer = sec_plate_info.get('customer', 'неизвестно') or item.get('customer', 'неизвестно')
                sec_width_m = sec_width / 1000.0
                
                # Ищем такую же плиту в списке (только среди вторичных!)
                sec_found = False
                for existing in track_plates:
                    if (existing['plate_name'] == sec_plate_name and 
                        existing['kp_id'] == sec_kp_id and
                        existing['kp_date'] == sec_kp_date and
                        existing['customer'] == sec_customer and
                        abs(existing['width_m'] - sec_width_m) < 0.01 and
                        existing.get('is_secondary', False) == True):  # Только с другими вторичными!
                        existing['qty'] += 1
                        sec_found = True
                        break
                # #region agent log H10: все вторичные резы (для диагностики пропущенных)
                if sec_width_m < 0.5 or '40,3' in sec_plate_name or '42,2' in sec_plate_name or '51-3,2' in sec_plate_name:
                    _d10 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
                    import json as _j10
                    with open(_d10, 'a', encoding='utf-8') as _f10:
                        _f10.write(_j10.dumps({"hypothesisId": "H10", "location": "production_completion:secondary_cut_trace", "message": "Вторичный рез обработан", "data": {"sec_plate_name": sec_plate_name, "sec_kp_id": sec_kp_id, "sec_length": sec_length, "sec_width_m": sec_width_m, "sec_width_mm": sec_width, "day": day_number, "merged": sec_found, "item_kp_id": item_kp_id, "parent_plate": item.get('plate_name', ''), "label": sec_cut.get('label', ''), "plate_info_kp_id": sec_plate_info.get('kp_id'), "plate_info_name": sec_plate_info.get('plate_name', '')}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                # #endregion
                # #region agent log H7: secondary cuts 51,8-5 / 58,5-5,3
                if '51,8-5' in sec_plate_name or '58,5-5,3' in sec_plate_name:
                    _d7 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
                    import json as _j7
                    with open(_d7, 'a', encoding='utf-8') as _f7:
                        _f7.write(_j7.dumps({"hypothesisId": "H7", "location": "production_completion:secondary_cut_added", "message": "Вторичный рез в track_plates", "data": {"sec_plate_name": sec_plate_name, "sec_kp_id": sec_kp_id, "sec_length": sec_length, "sec_width_m": sec_width_m, "day": day_number, "merged": sec_found}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                # #endregion
                if not sec_found:
                    track_plates.append({
                        'plate_name': sec_plate_name,
                        'length_m': sec_length,
                        'width_m': sec_width_m,
                        'load_class': int(cfg.normalize_load_code(item_load_code) * 100),
                        'qty': 1,
                        'kp_id': sec_kp_id,
                        'kp_date': sec_kp_date,
                        'customer': sec_customer,
                        'is_secondary': True,  # Флаг: это вторичный рез
                        'parent_kp_id': item_kp_id  # КП родительской плиты (для сохранения остатка)
                    })
        
        if track_plates:
            day_plates_by_track.append({
                'track_number': track_number,
                'plates': track_plates
            })
            total_qty += sum(p['qty'] for p in track_plates)
    
    # === ДОБАВЛЯЕМ ПЛИТЫ ИЗ ОСТАТКОВ (если есть) ===
    # Плиты из остатков добавляются как отдельная "дорожка" с номером 0
    plates_from_rests = data.get('plates_from_rests', [])
    rests_for_this_day = []  # Плиты из остатков для текущего дня
    
    if plates_from_rests and day_number == 1:
        # Плиты из остатков показываем только в День 1
        # (так как они не требуют производства на дорожках)
        for rest_plate in plates_from_rests:
            rests_for_this_day.append({
                'plate_name': rest_plate.get('plate_name', ''),
                'length_m': rest_plate.get('length_m', 0),
                'width_m': rest_plate.get('width_mm', 0) / 1000.0,
                'load_class': rest_plate.get('load_code', 8) * 100,
                'qty': rest_plate.get('qty', 1),
                'kp_id': rest_plate.get('kp_id'),
                'kp_date': rest_plate.get('kp_date', 'неизвестно'),
                'customer': rest_plate.get('customer', 'неизвестно'),
                # Флаги для отличия от обычных плит
                'from_rest': True,
                'rest_id': rest_plate.get('rest_id'),
                'match_type': rest_plate.get('match_type', 'exact'),
                'cut_cost': rest_plate.get('cut_cost', 0)
            })
        
        if rests_for_this_day:
            # Добавляем как "Дорожку 0: Из остатков" в начало списка
            day_plates_by_track.insert(0, {
                'track_number': 0,  # Номер 0 = из остатков
                'plates': rests_for_this_day,
                'is_from_rests': True  # Флаг для отображения в клавиатуре
            })
            total_qty += sum(p['qty'] for p in rests_for_this_day)
    
    if not day_plates_by_track:
        plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
        completed_days = data.get('completed_days', [])
        days_info = data.get('days_info', {})
        from_saved_plan = data.get('from_saved_plan', False)
        
        await callback.message.answer(
            f"❌ Не удалось найти плиты для Дня {day_number}.",
            reply_markup=calendar_days_kb(
                data.get('total_days', 1), 
                plan_start_date, 
                completed_days,
                days_info,
                show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
            )
        )
        await callback.answer()
        return
    
    # Подсчитываем общее количество позиций
    total_positions = sum(len(track['plates']) for track in day_plates_by_track)
    
    # ✅ НОВОЕ: Логируем результаты формирования списка для списания
    logger.info(f"[TRACE] ===== ШАГ 8: ПЛИТЫ ДЛЯ СПИСАНИЯ (day_plates_by_track) =====")
    logger.info(f"[TRACE] Дорожек: {len(day_plates_by_track)}")
    logger.info(f"[TRACE] Позиций (уникальных плит): {total_positions}")
    logger.info(f"[TRACE] Всего плит (с количеством): {total_qty}")
    
    for track_data in day_plates_by_track:
        track_num = track_data['track_number']
        plates_count = len(track_data['plates'])
        qty_sum = sum(p['qty'] for p in track_data['plates'])
        logger.info(f"[TRACE]   Дорожка #{track_num}: {plates_count} позиций, {qty_sum} шт")
        
        for plate in track_data['plates']:
            is_sec = " [ВТОРИЧНЫЙ]" if plate.get('is_secondary') else ""
            logger.info(f"[TRACE]     - {plate['plate_name']} × {plate['qty']} (ширина={plate['width_m']*1000:.0f}мм, КП #{plate.get('kp_id', '?')}){is_sec}")
    
    # Сохраняем данные для завершения дня
    await state.update_data(
        completing_day=day_number,
        day_plates_by_track=day_plates_by_track,
        rejected_quantities={},  # Словарь {(track_idx, plate_idx): количество_брака}
        active_plate_id=None  # Какая плита сейчас открыта для редактирования
    )
    
    # Формируем сообщение
    await callback.message.answer(
        f"📋 День {day_number} — завершение производства\n\n"
        f"Всего плит: {total_qty} шт ({total_positions} позиций)\n"
        f"Дорожек: {len(day_plates_by_track)}\n\n"
        f"❗ Отметьте плиты, которые ушли в БРАК:\n"
        f"(Бракованные плиты останутся на следующий день)\n\n"
        f"Нажмите на плиту, чтобы выбрать количество брака:",
        reply_markup=plates_completion_kb(day_plates_by_track, {}, None)
    )
    
    await state.set_state(ProductionStates.marking_completion)
    await callback.answer()


@router.callback_query(F.data.startswith("plate_open_"), ProductionStates.marking_completion)
async def open_plate_counter(callback: CallbackQuery, state: FSMContext):
    """
    Открывает счетчик брака под плитой.
    Позволяет выбрать количество бракованных плит кнопками +/-.
    """
    # Парсим формат: plate_open_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Устанавливаем активную плиту
    active_plate_id = (track_idx, plate_idx)
    await state.update_data(active_plate_id=active_plate_id)
    
    # Обновляем клавиатуру
    await callback.message.edit_reply_markup(
        reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_plus_"), ProductionStates.marking_completion)
async def increase_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Увеличивает количество брака на 1.
    Максимум = общему количеству плит.
    """
    # Парсим формат: reject_plus_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    active_plate_id = data.get('active_plate_id')
    
    # Получаем максимальное количество плит
    plate = day_plates_by_track[track_idx]['plates'][plate_idx]
    max_qty = plate.get('qty', 1)
    
    # Увеличиваем брак (макс = max_qty)
    plate_id = (track_idx, plate_idx)
    current = rejected_quantities.get(plate_id, 0)
    if current < max_qty:
        rejected_quantities[plate_id] = current + 1
        await state.update_data(rejected_quantities=rejected_quantities)
        
        # Обновляем клавиатуру
        await callback.message.edit_reply_markup(
            reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
        )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_minus_"), ProductionStates.marking_completion)
async def decrease_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Уменьшает количество брака на 1.
    Минимум = 0.
    """
    # Парсим формат: reject_minus_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    active_plate_id = data.get('active_plate_id')
    
    # Уменьшаем брак (мин = 0)
    plate_id = (track_idx, plate_idx)
    current = rejected_quantities.get(plate_id, 0)
    if current > 0:
        rejected_quantities[plate_id] = current - 1
        # Если стало 0, удаляем из словаря
        if rejected_quantities[plate_id] == 0:
            del rejected_quantities[plate_id]
        await state.update_data(rejected_quantities=rejected_quantities)
        
        # Обновляем клавиатуру
        await callback.message.edit_reply_markup(
            reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
        )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_reset_"), ProductionStates.marking_completion)
async def reset_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Сбрасывает брак в 0 и закрывает счетчик.
    """
    # Парсим формат: reject_reset_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Сбрасываем брак и закрываем счетчик
    plate_id = (track_idx, plate_idx)
    if plate_id in rejected_quantities:
        del rejected_quantities[plate_id]
    
    await state.update_data(
        rejected_quantities=rejected_quantities,
        active_plate_id=None
    )
    
    # Обновляем клавиатуру
    await callback.message.edit_reply_markup(
        reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, None)
    )
    await callback.answer("Брак сброшен")


@router.callback_query(F.data.startswith("reject_info_"), ProductionStates.marking_completion)
async def reject_info_click(callback: CallbackQuery):
    """
    Обработчик для центральной кнопки 'Брак: X/Y' (без действия).
    """
    await callback.answer()


@router.callback_query(F.data.startswith("track_header_"), ProductionStates.marking_completion)
async def track_header_click(callback: CallbackQuery):
    """
    Обработчик для заголовков дорожек.
    Не делает ничего, просто отвечает на callback.
    """
    await callback.answer("Это заголовок дорожки")


@router.callback_query(F.data == "confirm_completion", ProductionStates.marking_completion)
async def confirm_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Подтверждение завершения дня.
    Переносит выполненные плиты в completed_plates.
    Сохраняет информацию об остатках в plate_rests.
    """
    data = await state.get_data()
    day_number = data.get('completing_day', 1)
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Разделяем на выполненные и бракованные с учетом частичного брака
    completed_plates = []
    rejected_plates = []
    
    for track_idx, track_data in enumerate(day_plates_by_track):
        plates = track_data.get('plates', [])
        track_number = track_data.get('track_number', track_idx + 1)
        
        for plate_idx, plate in enumerate(plates):
            plate_id = (track_idx, plate_idx)
            reject_qty = rejected_quantities.get(plate_id, 0)
            total_qty = plate.get('qty', 1)
            
            if reject_qty == 0:
                # Все плиты выполнены
                completed_plates.append(plate)
            elif reject_qty >= total_qty:
                # Все плиты в браке
                plate_with_track = plate.copy()
                plate_with_track['track_number'] = track_number
                rejected_plates.append(plate_with_track)
            else:
                # ЧАСТИЧНЫЙ БРАК: разделяем плиту на 2 части
                # 1) Выполненная часть
                completed_part = plate.copy()
                completed_part['qty'] = total_qty - reject_qty
                completed_plates.append(completed_part)
                
                # 2) Бракованная часть
                rejected_part = plate.copy()
                rejected_part['qty'] = reject_qty
                rejected_part['track_number'] = track_number
                rejected_plates.append(rejected_part)
    
    # Импортируем функции БД
    db_path = str(PROJECT_ROOT / "plita.db")
    
    # plan_ids для Fallback-поиска плит в других КП того же плана
    source_plans = data.get('current_day_source_plans', [])
    active_plan_id = data.get('active_plan_id') or get_active_plan_id()
    plan_ids = list(source_plans) if source_plans else ([active_plan_id] if active_plan_id else [])
    
    # ИСПРАВЛЕНИЕ: Если plan_ids пуст, получаем все сохранённые планы из metadata
    # Это обеспечивает работу fallback-механизмов даже при отсутствии source_plans
    if not plan_ids:
        plan_ids = get_all_active_plan_ids()
        logger.info(f"[COMPLETION] plan_ids был пуст, загружено {len(plan_ids)} планов из metadata")
    
    # Группируем по kp_id
    plates_by_kp = defaultdict(list)
    plates_without_kp = []  # Плиты без kp_id (не найдены в lookup)
    # #region agent log
    import json as _json3
    _debug_log3 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    # #endregion
    
    # #region agent log H7: сводка по completed_plates (в т.ч. 5 типов, которые не списались)
    _d7 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    import json as _j7
    _key_substrings = (
        '51,8-5',
        '58,5-5,3',
        '61,1-12-10',
        '36,6-6,65',
        '64,8-12-12,5',
        '78,1-12-10',
        '78,1-12-12,5',
        # Плиты из текущего кейса пользователя
        '25,4-12-8п',
        '43-12-8п',
        '63,9-12-8п',
    )
    _targets = [p for p in completed_plates if any(s in p.get('plate_name', '') for s in _key_substrings)]
    with open(_d7, 'a', encoding='utf-8') as _f7:
        _f7.write(_j7.dumps({"hypothesisId": "H7", "location": "production_completion:completed_plates_summary", "message": "Целевые плиты в completed_plates", "data": {"day": day_number, "target_count": len(_targets), "targets": [{"name": p.get("plate_name"), "kp_id": p.get("kp_id"), "qty": p.get("qty"), "length_m": p.get("length_m"), "load_class": p.get("load_class")} for p in _targets], "total_plates": len(completed_plates)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
    # #endregion
    # #region agent log H15: ВСЕ вторичные резы в confirm_day_completion
    _sec_plates = [p for p in completed_plates if p.get('is_secondary')]
    _sec_total_qty = sum(p.get('qty', 1) for p in _sec_plates)
    _no_kp = [p for p in _sec_plates if not p.get('kp_id')]
    with open(_d7, 'a', encoding='utf-8') as _f15:
        _f15.write(_j7.dumps({"hypothesisId": "H15", "location": "production_completion:secondary_in_confirm", "message": "Вторичные резы в confirm", "data": {"day": day_number, "secondary_positions": len(_sec_plates), "secondary_qty": _sec_total_qty, "without_kp": len(_no_kp), "total_plates": len(completed_plates), "secondary_names": [{"name": p.get("plate_name"), "kp_id": p.get("kp_id"), "qty": p.get("qty"), "width_m": p.get("width_m")} for p in _sec_plates[:20]]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
    # #endregion
    for plate in completed_plates:
        kp_id = plate.get('kp_id')
        plate_name = plate.get('plate_name', '')
        width_m = plate.get('width_m', 0)
        # #region agent log H3/H5: плиты с/без kp_id
        if not kp_id or width_m in [0.46, 0.53, 0.5, 0.665] or '4,6' in plate_name or '5,3' in plate_name or '6,65' in plate_name or '51,8-5' in plate_name or '58,5-5,3' in plate_name:
            with open(_debug_log3, 'a', encoding='utf-8') as _f3:
                _f3.write(_json3.dumps({"hypothesisId": "H3", "location": "production_completion:confirm_day_completion", "message": "Плита при списании", "data": {"plate_name": plate_name, "kp_id": kp_id, "width_m": width_m, "has_kp_id": bool(kp_id), "qty": plate.get('qty', 1)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        # #endregion
        if kp_id:
            plates_by_kp[kp_id].append(plate)
        else:
            plates_without_kp.append(plate)
    
    total_moved = 0
    completed_kps = []
    
    # === ОБРАБОТКА ПЛИТ ИЗ ОСТАТКОВ ===
    # Сначала обрабатываем плиты из остатков (помечаем остатки как использованные)
    rests_used_count = 0
    for plate in completed_plates:
        if plate.get('from_rest') and plate.get('rest_id'):
            rest_id = plate['rest_id']
            # Помечаем остаток как использованный
            if kp_db.mark_rest_as_used(rest_id, db_path):
                rests_used_count += 1
                logger.info(f"[COMPLETION] Остаток #{rest_id} помечен как использованный")
            
            # Переносим плиту в completed_plates
            kp_id = plate.get('kp_id')
            if kp_id:
                moved = kp_db.move_plates_to_completed(kp_id, [plate], day_number, db_path, plan_ids=plan_ids)
                total_moved += moved
                
                if kp_db.check_and_update_kp_completion(kp_id, db_path):
                    if kp_id not in completed_kps:
                        completed_kps.append(kp_id)
    
    # НОВОЕ: Счётчик плит, сохранённых как остатки (вторичные резы без заказа)
    secondary_as_rests = 0
    
    # Переносим плиты С kp_id (стандартная логика, исключая плиты из остатков)
    # Вторичные резы с kp_id тоже отправляются на списание — они могут быть заказанными плитами
    # (узкие плиты из того же трека). Width-check в find_one_row не даст списать чужие.
    for kp_id, plates in plates_by_kp.items():
        # Фильтруем плиты из остатков (они уже обработаны выше)
        plates_not_from_rests = [p for p in plates if not p.get('from_rest')]
        if not plates_not_from_rests:
            continue
        
        moved = kp_db.move_plates_to_completed(kp_id, plates_not_from_rests, day_number, db_path, plan_ids=plan_ids)
        total_moved += moved
        
        # Проверяем, завершён ли КП полностью
        if kp_db.check_and_update_kp_completion(kp_id, db_path):
            if kp_id not in completed_kps:
                completed_kps.append(kp_id)
    
    # ========== НОВОЕ: Обрабатываем плиты БЕЗ kp_id ==========
    # Эти плиты не были найдены в lookup-таблицах (возможно из-за изменения ширины после реза).
    # Ищем их в БД по длине И по классу нагрузки (load_class), иначе находили плиту 8п вместо 10п/12,5п
    # и move_plates_to_completed не списывал — там везде фильтр по load_class.
    
    for plate in plates_without_kp:
        length_m = plate.get('length_m', 0)
        plate_name = plate.get('plate_name', '')
        load_class = plate.get('load_class', 800)
        is_secondary = plate.get('is_secondary', False)
        
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        # Ищем по длине И по классу нагрузки среди КП "в работе" и плит "в плане"
        cur.execute('''
            SELECT p.kp_id, p.plate_name 
            FROM kp_plates p
            JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE ABS(p.length_m - ?) < 0.02 
              AND p.load_class = ?
              AND p.qty > 0
              AND m.status = 'в работе'
              AND p.status = 'в плане'
            ORDER BY p.kp_id
            LIMIT 1
        ''', (length_m, load_class))
        row = cur.fetchone()
        conn.close()
        
        if row:
            found_kp_id = row[0]
            db_plate_name = row[1]
            plate['plate_name'] = db_plate_name  # Используем имя из БД для корректного списания
            
            moved = kp_db.move_plates_to_completed(found_kp_id, [plate], day_number, db_path, plan_ids=plan_ids)
            total_moved += moved
            logger.info(f"[COMPLETION] Плита найдена по длине+классу: {plate_name} ({length_m}м, {load_class}) → КП #{found_kp_id}")
            
            if kp_db.check_and_update_kp_completion(found_kp_id, db_path):
                if found_kp_id not in completed_kps:
                    completed_kps.append(found_kp_id)
        else:
            # ИСПРАВЛЕНИЕ: Если это вторичный рез (is_secondary=True) и его нет в БД,
            # сохраняем как остаток, а не выдаём предупреждение
            if is_secondary:
                # Это плита из вторичного реза (остатка), которую не заказывали в КП
                # Сохраняем её как остаток для будущего использования
                width_m = plate.get('width_m', 0)
                width_mm = int(width_m * 1000)
                
                # Берём kp_id родительской плиты (если есть в данных)
                parent_kp_id = plate.get('parent_kp_id') or 1  # fallback на КП #1
                
                kp_db.create_plate_rest(
                    kp_id=parent_kp_id,
                    source_plate_name=f"Вторичный рез: {plate_name}",
                    rest_width_mm=width_mm,
                    length_m=length_m,
                    production_day=day_number,
                    db_path=db_path
                )
                secondary_as_rests += 1
                logger.info(f"[COMPLETION] Вторичный рез сохранён как остаток: {plate_name} ({length_m}м × {width_mm}мм)")
            else:
                logger.warning(f"[COMPLETION] Плита не найдена в БД: {plate_name} ({length_m}м, load_class={load_class})")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========
    
    # ========== ВОЗВРАТ БРАКОВАННЫХ ПЛИТ В ПРОИЗВОДСТВО ==========
    # Бракованные плиты возвращаются в статус 'в производстве',
    # чтобы попасть в следующее планирование
    rejected_returned = 0
    for plate in rejected_plates:
        kp_id = plate.get('kp_id')
        plate_name = plate.get('plate_name')
        qty = plate.get('qty', 1)
        
        if kp_id and plate_name and qty > 0:
            success = kp_db.return_plates_to_production(
                kp_id=kp_id,
                plate_name=plate_name,
                qty=qty,
                db_path=db_path
            )
            if success:
                rejected_returned += 1
                logger.info(f"[COMPLETION] Брак: {plate_name} x{qty} возвращена в производство (КП #{kp_id})")
    
    if rejected_returned > 0:
        logger.info(f"[COMPLETION] Всего возвращено в производство: {rejected_returned} позиций (брак)")
    # ========== КОНЕЦ ВОЗВРАТА БРАКА ==========
    
    # ========== СОХРАНЕНИЕ ОСТАТКОВ ==========
    # Получаем метаданные об остатках из результата оптимизации
    optimization_result = data.get('optimization_result', {})
    rests_created = optimization_result.get('rests_created', [])
    rests_used = optimization_result.get('rests_used', [])
    
    # Подсчитываем неиспользованные остатки
    unused_rests_count = 0
    
    # Для каждого созданного остатка проверяем, использован ли он
    for rest in rests_created:
        is_used = any(
            abs(r['source_length'] - rest['length']) < 0.01 and 
            r['source_rest_mm'] == rest['rest_width_mm']
            for r in rests_used
        )
        
        if not is_used:
            # Остаток не использован - сохраняем в plate_rests
            # Находим kp_id для этого остатка через plate_lookup
            plate_lookup_exact = data.get('plate_lookup_exact', {})
            plate_lookup_by_length = data.get('plate_lookup_by_length', {})
            
            # Ищем kp_id по длине и ширине
            key = (round(rest['length'], 2), rest['source_width_mm'])
            plate_info_list = plate_lookup_exact.get(key, [])
            # plate_lookup_exact тоже возвращает СПИСОК, берём первый элемент
            plate_info = plate_info_list[0] if plate_info_list else None
            
            if not plate_info:
                length_key = round(rest['length'], 2)
                plate_info_list = plate_lookup_by_length.get(length_key, [])
                # plate_lookup_by_length возвращает СПИСОК, берём первый элемент
                plate_info = plate_info_list[0] if plate_info_list else None
            
            if plate_info and plate_info.get('kp_id'):
                kp_id = plate_info['kp_id']
                source_plate_name = plate_info.get('plate_name', f"ПБ {rest['length']}м-{rest['source_width_mm']}мм")
                
                kp_db.create_plate_rest(
                    kp_id=kp_id,
                    source_plate_name=source_plate_name,
                    rest_width_mm=rest['rest_width_mm'],
                    length_m=rest['length'],
                    production_day=day_number,
                    db_path=db_path
                )
                unused_rests_count += 1
    # ========== КОНЕЦ СОХРАНЕНИЯ ОСТАТКОВ ==========
    
    # Формируем отчёт
    report = f"✅ День {day_number} завершён!\n\n"
    report += f"📦 Выполнено плит: {total_moved} шт\n"
    
    # Информация о плитах из остатков
    if rests_used_count > 0:
        report += f"💰 Из остатков: {rests_used_count} шт (чистая прибыль!)\n"
    
    # Информация о вторичных резах, сохранённых как остатки
    if secondary_as_rests > 0:
        report += f"🔄 Вторичные резы (сохранены как остатки): {secondary_as_rests} шт\n"
    
    if rejected_plates:
        rejected_qty = sum(p['qty'] for p in rejected_plates)
        report += f"\n❌ В браке: {rejected_qty} шт ({len(rejected_plates)} позиций)\n"
        report += "Эти плиты останутся на следующий день:\n"
        
        # Группируем бракованные плиты по дорожкам для отчета
        rejected_by_track = defaultdict(list)
        for plate in rejected_plates:
            track_num = plate.get('track_number', '?')
            rejected_by_track[track_num].append(plate)
        
        for track_num in sorted(rejected_by_track.keys()):
            report += f"\n  Дорожка {track_num}:\n"
            for plate in rejected_by_track[track_num]:
                report += f"   • {plate['plate_name']} × {plate['qty']}\n"
    
    # Информация об остатках
    if unused_rests_count > 0:
        report += f"\n📦 Остатков на складе: {unused_rests_count} шт\n"
        report += "(Доступны для следующих заказов)"
    
    if completed_kps:
        report += f"\n🎉 Полностью выполнены КП: {', '.join(map(str, completed_kps))}\n"
        report += "Статус этих КП изменён на «выполнено»"
    
    # Показываем календарь с обновлённой галочкой
    total_days = data.get('total_days', 1)
    start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    days_info = data.get('days_info', {})
    completed_days = data.get('completed_days', [])
    
    # Обновляем days_info - помечаем текущий день как завершённый
    completing_day = day_number
    current_day_date = None
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        current_day_date = (start_dt + timedelta(days=completing_day - 1)).strftime('%Y-%m-%d')
    except:
        pass
    
    if current_day_date and current_day_date in days_info:
        days_info[current_day_date]['completed'] = True
    
    # ВАЖНО: Сохраняем изменения в файл плана!
    active_plan_id = data.get('active_plan_id') or get_active_plan_id()
    source_plans = data.get('current_day_source_plans', [])
    
    # Если есть список планов-источников (мультиплан), отмечаем день во всех планах
    if source_plans:
        for plan_id in source_plans:
            if mark_day_completed(plan_id, current_day_date):
                logger.info(f"День {current_day_date} отмечен как завершённый в плане {plan_id}")
            else:
                logger.warning(f"Не удалось отметить день {current_day_date} как завершённый в плане {plan_id}")
    elif active_plan_id and current_day_date:
        # Если работаем с одним активным планом
        if mark_day_completed(active_plan_id, current_day_date):
            logger.info(f"День {current_day_date} отмечен как завершённый в плане {active_plan_id}")
        else:
            logger.warning(f"Не удалось отметить день {current_day_date} как завершённый")
    
    # Обновляем completed_days в state
    if day_number not in completed_days:
        completed_days.append(day_number)
        completed_days.sort()
        await state.update_data(completed_days=completed_days)
    
    # Проверяем, работаем ли мы с сохранённым планом
    from_saved_plan = data.get('from_saved_plan', False)
    
    await callback.message.answer(report)
    await callback.message.answer(
        "📅 Календарь производства обновлён!\n"
        "Выберите следующий день:",
        reply_markup=calendar_days_kb(
            total_days, 
            start_date, 
            completed_days,
            days_info,
            show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
        )
    )
    await state.set_state(ProductionStates.viewing_calendar)
    await callback.answer("✅ День завершён!")


@router.callback_query(F.data == "cancel_completion", ProductionStates.marking_completion)
async def cancel_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Отмена завершения дня.
    Возвращаемся к выбору дней.
    """
    data = await state.get_data()
    total_days = data.get('total_days', 1)
    
    # Очищаем данные завершения, но оставляем остальные данные
    await state.update_data(
        completing_day=None,
        day_plates_by_track=None,
        rejected_quantities=None,
        active_plate_id=None
    )
    
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    
    # Проверяем, работаем ли мы с сохранённым планом
    from_saved_plan = data.get('from_saved_plan', False)
    
    await callback.message.answer(
        "❌ Завершение дня отменено.\n\n"
        "Выберите день для просмотра:",
        reply_markup=calendar_days_kb(
            total_days, 
            plan_start_date, 
            completed_days,
            days_info,
            show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
        )
    )
    await state.set_state(ProductionStates.viewing_calendar)
    await callback.answer()
