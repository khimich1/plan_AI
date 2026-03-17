"""Основная логика планирования производства - оптимизация и распределение по дням"""
import asyncio
import json
import logging
import sqlite3
import math
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict, Counter

logger = logging.getLogger(__name__)

from aiogram import Router
from aiogram.types import Message
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.db_config import PB_DB_PATH, PLITA_DB_PATH
from core.reinforcement_db import get_reinforcement
from core.optimization import optimize_with_cascading_longitudinal_cuts
import core.config_and_data as cfg
from core.config_and_data import PlateOrder, canonical_plate_key
import core.optimization as optimization

from ..keyboards import main_menu_kb, calendar_days_kb
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    get_global_day_occupancy,
    MAX_TRACKS_PER_DAY
)

# Допуски при привязке ключа трека к заказу для РЕСКЬЮ
RESCUE_TOL_LEN_M = 0.02
RESCUE_TOL_W_MM = 10
RESCUE_EXHAUSTED_RANK = 999999

# Путь к NDJSON-логу для отладки плит/ключей
_DEBUG_LOG = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
_DEBUG_SESSION_LOG = r"c:\Users\Роман\Desktop\Шишов\debug-d7e22e.log"
_DEBUG_RUNTIME_LOG = PROJECT_ROOT / "debug-73ca51.log"
_DEBUG_RUNTIME_SESSION_ID = "73ca51"


def _debug_write(hypothesis_id, location, message, data):
    """Пишет строку NDJSON в debug.log."""
    try:
        import time
        line = json.dumps({
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": time.time() * 1000
        }, ensure_ascii=False) + "\n"
        with open(_DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _debug_session_write(run_id, hypothesis_id, location, message, data):
    """Пишет NDJSON в debug-d7e22e.log для Debug Mode."""
    try:
        line = json.dumps({
            "sessionId": "d7e22e",
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(__import__("time").time() * 1000),
        }, ensure_ascii=False) + "\n"
        with open(_DEBUG_SESSION_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _debug_runtime_write(run_id, hypothesis_id, location, message, data):
    """Пишет NDJSON в debug-73ca51.log для текущей debug-сессии."""
    try:
        line = json.dumps({
            "sessionId": _DEBUG_RUNTIME_SESSION_ID,
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(__import__("time").time() * 1000),
        }, ensure_ascii=False) + "\n"
        with open(_DEBUG_RUNTIME_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

router = Router()


async def load_and_plan_production(message: Message, state: FSMContext):
    """
    Универсальная функция загрузки КП и планирования производства.
    Работает с разными способами фильтрации: date, kp, all, customer.
    """
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    filter_method = data.get('filter_method', 'date')
    
    # === ЗАГРУЗКА КП В ЗАВИСИМОСТИ ОТ ФИЛЬТРА ===
    db_path = PLITA_DB_PATH
    pb_db_path = PB_DB_PATH
    kp_db.init_schema(str(db_path))
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    kp_list = []
    
    if filter_method == 'date':
        # Логика по дате
        target_date_str = data.get('target_date')
        target_date = datetime.fromisoformat(target_date_str)
        date_description = data.get('date_description', target_date.strftime('%d.%m.%Y'))
        
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
        """)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            if not exec_terms:
                continue
            
            try:
                exec_date = datetime.strptime(exec_terms, '%d.%m.%Y')
                if exec_date <= target_date:
                    kp_list.append({
                        'kp_id': kp_id,
                        'date': exec_date,
                        'day': exec_date.day,
                        'customer': customer
                    })
            except:
                continue
    
    elif filter_method == 'kp':
        # Логика по номерам КП
        kp_ids = data.get('kp_ids', [])
        
        placeholders = ','.join('?' * len(kp_ids))
        cur.execute(f"""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE kp.kp_id IN ({placeholders})
              AND meta.status = 'в работе'
        """, kp_ids)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    elif filter_method == 'all':
        # Логика "все КП в работе"
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
        """)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    elif filter_method == 'customer':
        # Логика по заказчику
        customer_name = data.get('customer_name', '')
        
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
              AND kp.customer_name = ?
        """, (customer_name,))
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    if not kp_list:
        conn.close()
        await message.answer(
            "❌ Нет подходящих КП для производства.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        return
    
    kp_list.sort(key=lambda x: x['date'])
    
    await message.answer(f"✅ Найдено КП: {len(kp_list)}\nЗагружаю плиты...")
    
    # Опциональный фильтр по id плит (при выборе «По КП» с выбором плит)
    kp_plate_ids_raw = data.get('kp_plate_ids') or {}
    kp_plate_ids = {}
    if isinstance(kp_plate_ids_raw, dict):
        for k, v in kp_plate_ids_raw.items():
            sk = str(k)
            if v is not None and not isinstance(v, list):
                v = list(v) if hasattr(v, '__iter__') and not isinstance(v, str) else []
            kp_plate_ids[sk] = v
    # === ДАЛЬШЕ ВСЯ ТЕКУЩАЯ ЛОГИКА ===
    try:
        # === ШАГ 2: СОБИРАЕМ ПЛИТЫ ===
        plates_by_date_and_reinforcement = defaultdict(lambda: defaultdict(list))
        
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            kp_date = kp_info['date']
            plate_ids_for_kp = kp_plate_ids.get(str(kp_id)) if kp_plate_ids else None
            if plate_ids_for_kp is not None and len(plate_ids_for_kp) == 0:
                continue
            if plate_ids_for_kp and len(plate_ids_for_kp) > 0:
                placeholders = ','.join('?' * len(plate_ids_for_kp))
                cur.execute(f"""
                    SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве' AND id IN ({placeholders})
                    ORDER BY position_number, id
                """, (kp_id,) + tuple(plate_ids_for_kp))
            else:
                cur.execute("""
                    SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве'
                """, (kp_id,))
            
            for row in cur.fetchall():
                # length_dm_raw может отсутствовать в старых БД — берём по индексу
                plate_name, length_m, width_m, load_class, qty = row[0], row[1], row[2], row[3], row[4]
                length_dm_raw = (row[5] or '') if len(row) > 5 else ''
                load_code = cfg.normalize_load_code(load_class // 100)
                # Коррекция ширины по названию: "-12-8п" / "-12-" = 12 дм = 1200 мм; если в БД ошибочно 0.665 — подставляем 1200
                width_mm = round((width_m or 0) * 1000)
                if (plate_name and ("-12-8п" in plate_name or "-12-" in plate_name)) and (width_m is None or width_m < 0.9 or width_m > 1.5):
                    width_mm = 1200
                reinforcement_value = get_reinforcement(
                    length_m=length_m,
                    load_code=load_code,
                    source='series',
                    db_path=pb_db_path,
                    allow_fallback=True
                )
                
                if reinforcement_value is None:
                    reinforcement_value = 999.0
                
                plates_by_date_and_reinforcement[kp_date][reinforcement_value].append({
                    'plate_name': plate_name,
                    'length': length_m,
                    'width': width_mm,  # с учётом коррекции по названию для -12-8п
                    'load_code': load_code,
                    'qty': qty,
                    'reinforcement': reinforcement_value,
                    'kp_id': kp_id,
                    'kp_date': kp_date.strftime('%d.%m.%Y'),
                    'customer': kp_info['customer'],
                    'length_dm_raw': length_dm_raw,
                })
        
        # Создаём lookup-таблицы
        plate_to_kp_info = {}
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            plate_ids_for_kp = kp_plate_ids.get(str(kp_id)) if kp_plate_ids else None
            if plate_ids_for_kp is not None and len(plate_ids_for_kp) == 0:
                continue
            if plate_ids_for_kp and len(plate_ids_for_kp) > 0:
                placeholders = ','.join('?' * len(plate_ids_for_kp))
                cur.execute(f"""
                    SELECT plate_name, length_m, width_m
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве' AND id IN ({placeholders})
                """, (kp_id,) + tuple(plate_ids_for_kp))
            else:
                cur.execute("""
                    SELECT plate_name, length_m, width_m
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве'
                """, (kp_id,))
            for row in cur.fetchall():
                plate_name, length_m, width_m = row
                key = (round(length_m, 2), round(width_m * 1000))  # round для корректного округления
                if key not in plate_to_kp_info:
                    plate_to_kp_info[key] = {
                        'kp_id': kp_id,
                        'kp_date': kp_info['date'].strftime('%d.%m.%Y'),
                        'customer': kp_info['customer'],
                        'plate_name': plate_name
                    }
        
        conn.close()
        
        # === ШАГ 3: ФОРМИРУЕМ СПИСОК ПЛИТ ===
        selected_plates = []
        sorted_dates = sorted(plates_by_date_and_reinforcement.keys())
        
        for kp_date in sorted_dates:
            reinforcement_dict = plates_by_date_and_reinforcement[kp_date]
            sorted_reinforcements = sorted(reinforcement_dict.keys())
            
            for reinforcement in sorted_reinforcements:
                plates = reinforcement_dict[reinforcement]
                for plate_data in plates:
                    selected_plates.append(plate_data)
        
        if not selected_plates:
            await message.answer(
                "❌ Не найдено плит для планирования.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        # #region agent log H_366_loaded: плиты КП 2 после загрузки (до остатков)
        _sel_kp2 = [{"plate_name": p.get("plate_name"), "kp_id": p.get("kp_id"), "length": p.get("length"), "width": p.get("width"), "qty": p.get("qty")} for p in selected_plates if p.get("kp_id") == 2]
        try:
            with open(_DEBUG_LOG, "a", encoding="utf-8") as _fl:
                _fl.write(json.dumps({"hypothesisId": "H_366_loaded", "location": "production_execution:after_load_plates", "message": "Плиты КП №2 в selected_plates до остатков", "data": {"kp2_plates": _sel_kp2, "filter_method": filter_method, "kp_plate_ids_keys": list((kp_plate_ids or {}).keys())}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # === ШАГ 3.5: ПРОВЕРКА ОСТАТКОВ НА СКЛАДЕ ===
        plates_from_rests = []
        plates_for_optimizer = []
        
        plita_db_path = str(PROJECT_ROOT / 'plita.db')
        for plate_data in selected_plates:
            length_m = plate_data['length']
            width_mm = plate_data['width']
            qty_needed = plate_data['qty']
            
            matching_rests = kp_db.find_matching_rests(
                length_m=length_m,
                width_mm=width_mm,
                qty_needed=qty_needed,
                db_path=plita_db_path
            )
            
            qty_from_rests = 0
            
            if matching_rests:
                for rest_info in matching_rests:
                    qty_to_use = rest_info['qty_to_use']
                    qty_from_rests += qty_to_use
                    
                    plates_from_rests.append({
                        'plate_name': plate_data.get('plate_name', ''),
                        'length_m': length_m,
                        'width_mm': width_mm,
                        'qty': qty_to_use,
                        'kp_id': plate_data.get('kp_id'),
                        'kp_date': plate_data.get('kp_date', 'неизвестно'),
                        'customer': plate_data.get('customer', 'неизвестно'),
                        'load_code': cfg.normalize_load_code(plate_data.get('load_code', 8)),
                        'reinforcement': plate_data.get('reinforcement', 0),
                        'rest_id': rest_info['rest_id'],
                        'rest_length': rest_info['rest_length'],
                        'rest_width_mm': rest_info['rest_width_mm'],
                        'match_type': rest_info['match_type'],
                        'cut_cost': rest_info['cut_cost'],
                        'source_plate_name': rest_info['source_plate_name'],
                        'source_kp_id': rest_info['source_kp_id'],
                        'from_rest': True
                    })
            
            qty_remaining = qty_needed - qty_from_rests
            if qty_remaining > 0:
                plate_for_opt = plate_data.copy()
                plate_for_opt['qty'] = qty_remaining
                plates_for_optimizer.append(plate_for_opt)
        
        if plates_from_rests:
            total_from_rests = sum(p['qty'] for p in plates_from_rests)
            rests_msg = f"📦 Плиты из остатков: {total_from_rests} шт\n\n"
            
            for p in plates_from_rests:
                rests_msg += f"✅ {p['plate_name']} × {p['qty']} (КП #{p['kp_id']})\n"
                rests_msg += f"   Из остатка: {p['rest_length']}м × {p['rest_width_mm']}мм\n"
                
                if p['match_type'] == 'exact':
                    rests_msg += f"   Точное совпадение, себестоимость: 0 руб.\n"
                elif p['match_type'] == 'width_cut':
                    rests_msg += f"   Резы: продольный ({p['length_m']}м)\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                elif p['match_type'] == 'length_cut':
                    rests_msg += f"   Резы: поперечный\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                else:
                    rests_msg += f"   Резы: продольный ({p['length_m']}м), поперечный\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                
                rests_msg += "\n"
            
            rests_msg += "💰 Эти плиты уже оплачены - чистая прибыль!"
            await message.answer(rests_msg)
        
        if not plates_for_optimizer:
            await message.answer(
                "✅ Все плиты можно взять из остатков!\n"
                "Оптимизация не требуется.",
                reply_markup=main_menu_kb()
            )
            await state.update_data(
                plates_from_rests=plates_from_rests,
                all_from_rests=True
            )
            await state.clear()
            return
        
        selected_plates = plates_for_optimizer
        await message.answer("⏳ Запускаю оптимизацию раскроя...")
        
        # === ШАГ 4: ОПТИМИЗАЦИЯ ===
        orders_2d = []
        for plate_data in selected_plates:
            orders_2d.append({
                'length': plate_data['length'],
                'width': plate_data['width'],
                'qty': plate_data['qty'],
                'load_code': cfg.normalize_load_code(plate_data['load_code']),
                'reinforcement': plate_data['reinforcement'],
                'kp_date': plate_data.get('kp_date', 'неизвестно'),
                'customer': plate_data.get('customer', 'неизвестно'),
                'plate_name': plate_data.get('plate_name', ''),
                'kp_id': plate_data.get('kp_id'),
                'length_dm_raw': plate_data.get('length_dm_raw', '') or '',
            })
        # #region agent log
        _targets = []
        for _o in orders_2d:
            _n = (_o.get("plate_name") or "")
            if any(_k in _n for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п")):
                _targets.append({
                    "plate_name": _n,
                    "length": _o.get("length"),
                    "width": _o.get("width"),
                    "load_code": _o.get("load_code"),
                    "qty": _o.get("qty", 1),
                    "kp_id": _o.get("kp_id"),
                })
        _debug_session_write(
            "run1",
            "H1",
            "production_execution:orders_2d_built",
            "Target plates after parsing and loading",
            {
                "targets_count": len(_targets),
                "targets": _targets,
                "orders_total_qty": sum(int(x.get("qty", 0) or 0) for x in orders_2d),
            },
        )
        # #endregion
        # #region agent log H_366: плита 36,6-6,65 в orders_2d при загрузке плана
        _log_366 = [{"plate_name": o.get("plate_name"), "kp_id": o.get("kp_id"), "length": o.get("length"), "width": o.get("width"), "qty": o.get("qty", 1)} for o in orders_2d if o.get("kp_id") == 2 and ("36,6" in (o.get("plate_name") or "") or "6,65" in (o.get("plate_name") or "") or (abs(float(o.get("length", 0)) - 3.66) < 0.01 and o.get("width") == 665))]
        if _log_366 or any(o.get("kp_id") == 2 for o in orders_2d):
            try:
                with open(_DEBUG_LOG, "a", encoding="utf-8") as _f366:
                    _f366.write(json.dumps({"hypothesisId": "H_366", "location": "production_execution:orders_2d_after_build", "message": "КП №2 и/или плита 36,6-6,65 в orders_2d", "data": {"kp2_orders": _log_366, "all_kp2_count": sum(1 for o in orders_2d if o.get("kp_id") == 2), "total_orders": len(orders_2d)}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        # #region agent log
        _orders_by_key = Counter(
            (
                round(float(o.get('length', 0)), 2),
                int(round(float(o.get('width', 0) or 0))),
                cfg.normalize_load_code(o.get('load_code', 8)),
            )
            for o in orders_2d
            for _ in range(int(o.get('qty', 1) or 0))
        )
        _debug_runtime_write(
            "run1",
            "H1_input_orders",
            "production_execution:orders_2d_built",
            "Demand snapshot before optimizer",
            {
                "orders_total_qty": int(sum(int(o.get('qty', 0) or 0) for o in orders_2d)),
                "orders_total_lines": len(orders_2d),
                "orders_by_key": {str(list(k)): int(v) for k, v in _orders_by_key.items()},
            },
        )
        # #endregion
        # Логируем уникальные load_code в orders_2d (план: этап 2.3)
        unique_loads = set(o['load_code'] for o in orders_2d)
        logger.info(f"[DEMAND] Уникальные load_code в orders_2d: {sorted(unique_loads)}")
        if 12 in unique_loads and 12.5 not in unique_loads:
            kp_ids_in_production = list({p.get('kp_id') for p in selected_plates if p.get('kp_id')})
            if kp_ids_in_production:
                try:
                    with kp_db._connect(kp_db.DEFAULT_DB) as conn:
                        cur = conn.cursor()
                        placeholders = ','.join('?' * len(kp_ids_in_production))
                        cur.execute(
                            f"SELECT 1 FROM kp_plates WHERE kp_id IN ({placeholders}) AND load_class = 1250 LIMIT 1",
                            kp_ids_in_production
                        )
                        if cur.fetchone():
                            logger.warning("[DEMAND] ⚠️ В БД есть плиты 12.5п, но в orders_2d только 12п!")
                except Exception:
                    pass
        # #region agent log: 59,9-12-10п trace (H1,H2)
        _log_59_10 = [{"length": o["length"], "width": o["width"], "plate_name": o.get("plate_name", ""), "kp_id": o.get("kp_id"), "qty": o.get("qty", 1)} for o in orders_2d if 5.98 <= float(o.get("length", 0)) <= 6.0 and cfg.normalize_load_code(o.get("load_code", 8)) == 10]
        if _log_59_10:
            try:
                with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8") as _f:
                    _f.write(__import__("json").dumps({"hypothesisId": "H_59_10_source", "location": "production_execution:orders_2d_built", "message": "orders_2d: плиты 5.98-6м 10п (ширина для 59,9-12-10п)", "data": {"orders": _log_59_10}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        # #region agent log H_orders: плиты 61,2 и 59,8 в orders_2d при планировании
        _log_61_59 = [{"plate_name": o.get("plate_name", ""), "kp_id": o.get("kp_id"), "length": o.get("length"), "width": o.get("width"), "qty": o.get("qty", 1)} for o in orders_2d if ("61,2" in (o.get("plate_name") or "") or "59,8" in (o.get("plate_name") or ""))]
        try:
            with open(_DEBUG_LOG, "a", encoding="utf-8") as _f:
                _f.write(__import__("json").dumps({"hypothesisId": "H_orders", "location": "production_execution:orders_2d_built", "message": "Плиты 61,2 и 59,8 в orders_2d", "data": {"count": len(_log_61_59), "entries": _log_61_59}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # ✅ НОВОЕ: Логируем все плиты ДО оптимизации
        logger.info(f"[TRACE] ===== ШАГ 1: ПЛИТЫ ДО ОПТИМИЗАЦИИ =====")
        logger.info(f"[TRACE] Всего плит: {sum(p['qty'] for p in orders_2d)}")
        for p in orders_2d:
            logger.info(f"[TRACE]   {p['plate_name']} × {p['qty']} (длина={p['length']:.2f}м, ширина={p['width']}мм, КП #{p.get('kp_id', '?')})")
        
        # Lookup-таблицы: ключ строго (length, width), без слияния длин (5.7 и 5.71 — разные ключи).
        plate_lookup_exact = {}
        plate_lookup_by_length = {}
        for order in orders_2d:
            L = round(order['length'], 2)
            W = order['width']
            entry = {
                'kp_date': order.get('kp_date', 'неизвестно'),
                'customer': order.get('customer', 'неизвестно'),
                'plate_name': order.get('plate_name', ''),
                'reinforcement': order.get('reinforcement', 0),
                'load_code': cfg.normalize_load_code(order.get('load_code', 8)),
                'qty_remaining': order.get('qty', 1),
                'kp_id': order.get('kp_id'),
            }
            key = (L, W)
            if key not in plate_lookup_exact:
                plate_lookup_exact[key] = []
            plate_lookup_exact[key].append(entry)
        
        for order in orders_2d:
            length_key = round(order['length'], 2)
            length_entry = {
                'kp_date': order.get('kp_date', 'неизвестно'),
                'customer': order.get('customer', 'неизвестно'),
                'plate_name': order.get('plate_name', ''),
                'reinforcement': order.get('reinforcement', 0),
                'qty_remaining': order.get('qty', 1),
                'kp_id': order.get('kp_id'),
            }
            if length_key not in plate_lookup_by_length:
                plate_lookup_by_length[length_key] = []
            plate_lookup_by_length[length_key].append(length_entry)
        
        # Сортируем списки по дате КП
        def parse_date_for_sort(date_str):
            if not date_str or date_str == 'неизвестно':
                return datetime.max
            try:
                return datetime.strptime(date_str, '%d.%m.%Y')
            except:
                return datetime.max
        
        for key in plate_lookup_exact:
            plate_lookup_exact[key].sort(key=lambda x: parse_date_for_sort(x.get('kp_date', '')))
        
        for key in plate_lookup_by_length:
            plate_lookup_by_length[key].sort(key=lambda x: parse_date_for_sort(x.get('kp_date', '')))
        
        # Запуск оптимизации
        optimization_result = await asyncio.to_thread(
            optimize_with_cascading_longitudinal_cuts,
            orders_2d=orders_2d
        )
        if not optimization_result or optimization_result.get('total_plates', 0) == 0:
            await message.answer(
                "❌ Оптимизация не дала результатов.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        await message.answer(f"✅ Оптимизация завершена! Исходных плит: {optimization_result.get('total_plates', 0)}")
        
        # ✅ НОВОЕ: Логируем результаты оптимизации
        logger.info(f"[TRACE] ===== ШАГ 2: РЕЗУЛЬТАТЫ ОПТИМИЗАЦИИ =====")
        logger.info(f"[TRACE] Первичных резов: {len(optimization_result.get('primary_cuts', []))}")
        logger.info(f"[TRACE] Вторичных резов: {len(optimization_result.get('secondary_cuts', []))}")
        
        # Подсчитываем плиты по типам
        primary_plates_count = 0
        for cut in optimization_result.get('primary_cuts', []):
            primary_plates_count += cut.get('qty', 0)
            logger.info(f"[TRACE]   Первичный: {cut['width']}мм + {cut['rest']}мм × {cut['qty']} (длины={cut.get('lengths', [])})")
        
        secondary_plates_count = 0
        for cut in optimization_result.get('secondary_cuts', []):
            secondary_plates_count += cut.get('qty', 0) * cut.get('pieces', 1)
            logger.info(f"[TRACE]   Вторичный: {cut['source']}мм → {cut['cuts']} × {cut['qty']} (тип={cut.get('type', '?')})")
        
        logger.info(f"[TRACE] Плит из первичных резов: {primary_plates_count}")
        logger.info(f"[TRACE] Плит из вторичных резов: {secondary_plates_count}")
        logger.info(f"[TRACE] Всего плит после оптимизации: {primary_plates_count + secondary_plates_count}")
        
        # === ПРОВЕРКА: вход vs выход оптимизатора (только первичные резы) ===
        # plate_assignments содержит первичные + вторичные; сравниваем только с первичными,
        # вторичные — бонусные плиты из остатков, их больше, чем заказано, и это ожидаемо.
        input_plates = sum(p['qty'] for p in orders_2d)
        all_assignments = optimization_result.get('plate_assignments', [])
        output_plates_primary = sum(1 for p in all_assignments if p.get('source') == 'primary')
        output_plates_total = len(all_assignments)
        if input_plates != output_plates_primary:
            logger.error(
                f"[OPT_CHECK] ❌ РАСХОЖДЕНИЕ! Запрошено плит: {input_plates}, "
                f"первичных в результате: {output_plates_primary}, "
                f"всего в plate_assignments (вкл. вторичные): {output_plates_total}, "
                f"потеряно первичных: {input_plates - output_plates_primary}"
            )
            ordered = Counter(
                canonical_plate_key(o['length'], o['width'], o.get('load_code', 8))
                for o in orders_2d for _ in range(o['qty'])
            )
            def _plate_key(p):
                return canonical_plate_key(
                    p.get('length', 0), p.get('width', 0), p.get('load_code', 8)
                )
            produced = Counter(
                _plate_key(p) for p in all_assignments if p.get('source') == 'primary'
            )
            missing = ordered - produced
            if missing:
                logger.error(f"[OPT_CHECK] Не хватает в результате: {dict(missing)}")
        else:
            logger.info(
                f"[OPT_CHECK] ✅ Совпадение: запрошено {input_plates} плит, "
                f"первичных получено {output_plates_primary}, "
                f"вторичных дополнительно: {output_plates_total - output_plates_primary}"
            )
        # #region agent log
        _ordered_debug = Counter(
            (
                round(float(o.get('length', 0)), 2),
                int(round(float(o.get('width', 0) or 0))),
                cfg.normalize_load_code(o.get('load_code', 8)),
            )
            for o in orders_2d
            for _ in range(int(o.get('qty', 1) or 0))
        )
        _produced_debug = Counter(
            (
                round(float(p.get('length', 0)), 2),
                int(round(float(p.get('width', 0) or 0))),
                cfg.normalize_load_code(p.get('load_code', 8)),
            )
            for p in all_assignments
            if p.get('source') == 'primary'
        )
        _debug_runtime_write(
            "run1",
            "H2_optimizer_output",
            "production_execution:after_optimizer",
            "Optimizer demand vs produced(primary)",
            {
                "input_plates": int(input_plates),
                "output_primary": int(output_plates_primary),
                "output_total_assignments": int(output_plates_total),
                "missing_primary_keys": {str(list(k)): int(v) for k, v in (_ordered_debug - _produced_debug).items()},
                "extra_primary_keys": {str(list(k)): int(v) for k, v in (_produced_debug - _ordered_debug).items()},
            },
        )
        # #endregion
        
        # #region agent log (73b708) H_WHERE_OPT: сколько плит по целевым ключам вышло из оптимизатора
        try:
            _target_keys = [(5.08, 320, 8), (5.98, 665, 8)]
            _pa = optimization_result.get('plate_assignments', []) or []
            _opt_counts = {}
            for _k in _target_keys:
                _opt_counts[str(_k)] = 0
            for _p in _pa:
                _pl = round(float(_p.get('length', 0)), 2)
                _pw = int(round(float(_p.get('width', 0))))
                _plc = cfg.normalize_load_code(_p.get('load_code', 8))
                for _tk in _target_keys:
                    if abs(_pl - _tk[0]) < 0.02 and _pw == _tk[1] and _plc == _tk[2]:
                        _opt_counts[str(_tk)] = _opt_counts.get(str(_tk), 0) + 1
                        break
            _agent_log = PROJECT_ROOT / "debug-73b708.log"
            with open(_agent_log, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_WHERE_OPT", "location": "production_execution:after_optimizer", "message": "Plates by key at optimizer output", "data": {"target_keys": _target_keys, "counts": _opt_counts, "total_plate_assignments": len(_pa)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # === ШАГ 4.5: ДОПОЛНЕНИЕ LOOKUP ДЛЯ ВТОРИЧНЫХ РЕЗОВ ===
        if optimization_result.get('secondary_cuts'):
            orders_dict = {}
            for order in orders_2d:
                key = (
                    round(order['length'], 2),
                    order['width'],
                    cfg.normalize_load_code(order.get('load_code', 8))
                )
                if key not in orders_dict:
                    orders_dict[key] = []
                orders_dict[key].append(order)
            
            for sec_cut in optimization_result['secondary_cuts']:
                target_key = sec_cut.get('target_order_key')
                if not target_key:
                    continue
                
                target_length, target_width, target_load_code = target_key
                original_orders = orders_dict.get((round(target_length, 2), target_width, target_load_code), [])
                
                if not original_orders:
                    continue
                
                result_lengths = sec_cut.get('lengths', [])
                result_width = sec_cut['cuts'][0]
                
                for result_length in result_lengths:
                    key_result = (round(result_length, 2), result_width)
                    
                    original_order = None
                    for order in original_orders:
                        original_key = (round(order['length'], 2), order['width'])
                        if original_key in plate_lookup_exact:
                            for entry in plate_lookup_exact[original_key]:
                                if entry.get('qty_remaining', 0) > 0:
                                    original_order = order
                                    break
                        if original_order:
                            break
                    
                    if not original_order:
                        continue
                    
                    if key_result not in plate_lookup_exact:
                        plate_lookup_exact[key_result] = []
                    
                    plate_lookup_exact[key_result].append({
                        'kp_date': original_order.get('kp_date', 'неизвестно'),
                        'customer': original_order.get('customer', 'неизвестно'),
                        'plate_name': original_order.get('plate_name', ''),
                        'reinforcement': original_order.get('reinforcement', 0),
                        'load_code': cfg.normalize_load_code(original_order.get('load_code', 8)),
                        'qty_remaining': 1,
                        'kp_id': original_order.get('kp_id'),
                        'is_from_secondary': True
                    })
                    
                    length_key = round(result_length, 2)
                    if length_key not in plate_lookup_by_length:
                        plate_lookup_by_length[length_key] = []
                    
                    plate_lookup_by_length[length_key].append({
                        'kp_date': original_order.get('kp_date', 'неизвестно'),
                        'customer': original_order.get('customer', 'неизвестно'),
                        'plate_name': original_order.get('plate_name', ''),
                        'reinforcement': original_order.get('reinforcement', 0),
                        'qty_remaining': 1,
                        'kp_id': original_order.get('kp_id'),
                        'is_from_secondary': True
                    })
            
            logger.info(
                f"[PRODUCTION] Дополнено lookup-таблиц: exact={len(plate_lookup_exact)}, by_length={len(plate_lookup_by_length)}"
            )
        
        # === ШАГ 5: ПОДГОТОВКА ДЛЯ ВИЗУАЛИЗАЦИИ ===
        optimization.OPT_CASCADING_PLAN = optimization_result
        
        all_loads = set(p['load_code'] for p in orders_2d)
        optimization_result['loads_in_group'] = sorted(all_loads)
        optimization.OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}
        optimization.LOAD_TO_REINFORCEMENT_MAP = {
            load_code: ['all'] for load_code in all_loads
        }
        
        PlateOrder.from_orders_2d(orders_2d).apply_to_globals()
        
        # #region agent log (95694e) количество 5.98/665 в результате оптимизации (primary_cuts)
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            _n_opt = 0
            for _c in optimization_result.get('primary_cuts', []) or []:
                _L = round(float((_c.get('lengths') or [6.0])[0]), 2)
                _w = _c.get('width') or 1200
                if abs(_L - 5.98) < 0.02 and _w == 665:
                    _n_opt += _c.get('qty', 1)
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_opt_598665", "location": "production_execution:after_optimization", "message": "count 5.98/665 in optimization_result primary_cuts", "data": {"count_598_665": _n_opt}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) количество 5.08/320 и 5.98/530 в primary_cuts
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            _n_508320 = _n_598530 = 0
            for _c in optimization_result.get('primary_cuts', []) or []:
                _L = round(float((_c.get('lengths') or [6.0])[0]), 2)
                _w = _c.get('width') or 1200
                if abs(_L - 5.08) < 0.02 and _w == 320:
                    _n_508320 += _c.get('qty', 1)
                if abs(_L - 5.98) < 0.02 and _w == 530:
                    _n_598530 += _c.get('qty', 1)
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_opt_rescue", "location": "production_execution:after_optimization", "message": "count 5.08/320 and 5.98/530 in primary_cuts", "data": {"count_508_320": _n_508320, "count_598_530": _n_598530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # === ШАГ 6: ПОДСЧЕТ ДОРОЖЕК ===
        await message.answer("⏳ Подсчитываю дорожки...")
        
        from viz_modules.layout_sequence import build_layout_sequence
        from core.visualization import split_sequence_into_tracks
        from core.plate_audit import PlateAudit as _PlateAudit

        # Подхватываем audit из оптимизатора, если он там был создан
        _handler_audit: _PlateAudit | None = optimization_result.get('_plate_audit')
        if _handler_audit is None:
            _handler_audit = _PlateAudit(orders_2d)

        seq = build_layout_sequence()
        # #region agent log (73b708) H_LAYOUT_IN: целостность после build_layout_sequence
        try:
            _pa_count = len(optimization_result.get('plate_assignments', []) or [])
            if isinstance(seq, list) and seq and isinstance(seq[0], dict) and seq[0].get('load_code') is not None:
                _seq_total = sum(len(g.get('sequence', [])) for g in seq)
                _format = "grouped"
            else:
                _seq_total = len(seq) if seq else 0
                _format = "flat"
            _agent_log = PROJECT_ROOT / "debug-73b708.log"
            with open(_agent_log, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_LAYOUT_IN", "location": "production_execution:after_build_layout_sequence", "message": "Sequence vs plate_assignments count", "data": {"sequence_total": _seq_total, "plate_assignments_count": _pa_count, "format": _format, "diff": _seq_total - _pa_count}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) количество 5.98/665 в последовательности до split
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            def _count_598_665_in_seq(s):
                n = 0
                if isinstance(s, list) and s and isinstance(s[0], dict) and s[0].get('load_code') is not None:
                    for g in s:
                        for it in g.get('sequence', []):
                            L = round(float(it.get('length', 0) or it.get('target_length', 0)), 2)
                            w = it.get('width') or it.get('main_w') or 1.2
                            w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                            if abs(L - 5.98) < 0.02 and w_mm == 665:
                                n += 1
                return n
            _n_seq = _count_598_665_in_seq(seq)
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_seq_598665", "location": "production_execution:after_build_layout", "message": "count 5.98/665 in sequence before split", "data": {"count_598_665": _n_seq}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) количество 5.08/320 и 5.98/530 в последовательности до split
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            _n508, _n598 = 0, 0
            if isinstance(seq, list) and seq and isinstance(seq[0], dict) and seq[0].get('load_code') is not None:
                for g in seq:
                    for it in g.get('sequence', []):
                        L = round(float(it.get('length', 0) or it.get('target_length', 0)), 2)
                        w = it.get('width') or it.get('main_w') or 1.2
                        w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                        if abs(L - 5.08) < 0.02 and w_mm == 320:
                            _n508 += 1
                        if abs(L - 5.98) < 0.02 and w_mm == 530:
                            _n598 += 1
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_seq_rescue", "location": "production_execution:after_build_layout", "message": "count 5.08/320 and 5.98/530 in sequence before split", "data": {"count_508_320": _n508, "count_598_530": _n598}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # PlateAudit: checkpoint после build_layout_sequence
        _handler_audit.checkpoint("layout_sequence", seq)

        all_tracks_list = split_sequence_into_tracks(seq)

        # PlateAudit: checkpoint после split_sequence_into_tracks
        _handler_audit.checkpoint("tracks", all_tracks_list)

        # #region agent log
        def _count_keys_in_seq_and_tracks(order_keys_set):
            seq_counts = Counter()
            tracks_counts = Counter()
            if isinstance(seq, list) and seq and isinstance(seq[0], dict) and seq[0].get('load_code') is not None:
                for g in seq:
                    for it in g.get('sequence', []):
                        l = round(float(it.get('length', 0) or it.get('target_length', 0) or 0), 2)
                        wv = it.get('width') if it.get('width') is not None else it.get('main_w')
                        w = int(round(float(wv) * 1000)) if wv is not None and float(wv) < 20 else int(round(float(wv or 0)))
                        lc = cfg.normalize_load_code(it.get('load_code', 8))
                        k = (l, w, lc)
                        if k in order_keys_set:
                            seq_counts[k] += 1
            for tr in all_tracks_list or []:
                for it in tr.get('items', []) or []:
                    l = round(float(it.get('target_length') if it.get('mode') == 'transverse' and it.get('target_length') is not None else it.get('length', 0) or 0), 2)
                    wv = it.get('main_w') if it.get('mode') == 'split' else it.get('width')
                    w = int(round(float(wv) * 1000)) if wv is not None and float(wv) < 20 else int(round(float(wv or 0)))
                    lc = cfg.normalize_load_code(it.get('load_code', 8))
                    k = (l, w, lc)
                    if k in order_keys_set:
                        tracks_counts[k] += 1
                    for sc in it.get('secondary_cuts', []) or []:
                        sl = round(float(sc.get('target_length') or l), 2)
                        sw = int(round(float(sc.get('width', 0))))
                        slc = cfg.normalize_load_code(sc.get('load_code', it.get('load_code', 8)))
                        sk = (sl, sw, slc)
                        if sk in order_keys_set:
                            tracks_counts[sk] += 1
            return seq_counts, tracks_counts

        _order_keys_set = set(
            (
                round(float(o.get('length', 0)), 2),
                int(round(float(o.get('width', 0) or 0))),
                cfg.normalize_load_code(o.get('load_code', 8)),
            )
            for o in orders_2d
        )
        _seq_counts, _tracks_counts = _count_keys_in_seq_and_tracks(_order_keys_set)
        _debug_runtime_write(
            "run1",
            "H3_layout_split",
            "production_execution:after_split_tracks",
            "Counts by ordered keys in sequence and tracks",
            {
                "ordered_keys_count": len(_order_keys_set),
                "sequence_counts": {str(list(k)): int(v) for k, v in _seq_counts.items()},
                "tracks_counts": {str(list(k)): int(v) for k, v in _tracks_counts.items()},
                "lost_on_split_keys": {str(list(k)): int(v) for k, v in (_seq_counts - _tracks_counts).items()},
            },
        )
        # #endregion
        # #region agent log (95694e) количество 5.98/665 в дорожках после split
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            _n_tracks = 0
            for tr in (all_tracks_list or []):
                for it in tr.get('items', []) or []:
                    L = round(float(it.get('length', 0) or it.get('target_length', 0)), 2)
                    w = it.get('width') or it.get('main_w') or 1.2
                    w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                    if abs(L - 5.98) < 0.02 and w_mm == 665:
                        _n_tracks += 1
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_tracks_598665", "location": "production_execution:after_split", "message": "count 5.98/665 in tracks after split", "data": {"count_598_665": _n_tracks}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) количество 5.08/320 и 5.98/530 в дорожках после split
        try:
            _log_95694e = PROJECT_ROOT / "debug-95694e.log"
            _n508, _n598 = 0, 0
            for tr in (all_tracks_list or []):
                for it in tr.get('items', []) or []:
                    L = round(float(it.get('length', 0) or it.get('target_length', 0)), 2)
                    w = it.get('width') or it.get('main_w') or 1.2
                    w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                    if abs(L - 5.08) < 0.02 and w_mm == 320:
                        _n508 += 1
                    if abs(L - 5.98) < 0.02 and w_mm == 530:
                        _n598 += 1
            with open(_log_95694e, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_tracks_rescue", "location": "production_execution:after_split", "message": "count 5.08/320 and 5.98/530 in tracks after split", "data": {"count_508_320": _n508, "count_598_530": _n598}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (73b708) H_WHERE_TRACKS: сколько плит по целевым ключам в треках после split_sequence_into_tracks
        try:
            _target_keys = [(5.08, 320, 8), (5.98, 665, 8)]
            def _w_mm(w, default=1.2):
                if w is None:
                    return round(default * 1000)
                try:
                    f = float(w)
                    return round(f * 1000) if f < 20 else round(f)
                except (TypeError, ValueError):
                    return round(default * 1000)
            _track_counts = {str(_k): 0 for _k in _target_keys}
            for _tr in all_tracks_list:
                for _it in _tr.get('items', []) or []:
                    if not _it:
                        continue
                    _mode = _it.get('mode', 'solid')
                    _lc = cfg.normalize_load_code(_it.get('load_code', 8))
                    if _mode == 'split':
                        _wm = _w_mm(_it.get('main_w', 1.2))
                    else:
                        _wm = _w_mm(_it.get('width', 1.2))
                    if _mode == 'transverse' and _it.get('target_length') is not None:
                        _ln = round(float(_it.get('target_length') or 0), 2)
                    else:
                        _ln = round(float(_it.get('length', 0) or 0), 2)
                    if _ln <= 0:
                        continue
                    for _tk in _target_keys:
                        if abs(_ln - _tk[0]) < 0.02 and _wm == _tk[1] and _lc == _tk[2]:
                            _track_counts[str(_tk)] = _track_counts.get(str(_tk), 0) + 1
                            break
                    for _sc in _it.get('secondary_cuts', []) or []:
                        _sw = _w_mm(_sc.get('width', 0), 0)
                        _sl = round(float(_sc.get('target_length') or _ln), 2)
                        if _sw <= 0:
                            continue
                        _slc = cfg.normalize_load_code(_sc.get('load_code', _it.get('load_code', 8)))
                        for _tk in _target_keys:
                            if abs(_sl - _tk[0]) < 0.02 and _sw == _tk[1] and _slc == _tk[2]:
                                _track_counts[str(_tk)] = _track_counts.get(str(_tk), 0) + 1
                                break
            _agent_log = PROJECT_ROOT / "debug-73b708.log"
            with open(_agent_log, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_WHERE_TRACKS", "location": "production_execution:after_split_tracks", "message": "Plates by key in all_tracks_list", "data": {"target_keys": _target_keys, "counts": _track_counts}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # === ШАГ 6.5: ЗАЩИТА ОТ ПОТЕРИ ПЛИТ (РЕСКЬЮ) ===
        # Приводим ключ дорожки к ключу заказа. С учётом remaining заполняем сначала позиции с меньшим спросом (3.79→1 шт).
        def _normalize_key_to_orders(key, order_keys, order_counts, tol_len=RESCUE_TOL_LEN_M, tol_w=RESCUE_TOL_W_MM, remaining=None):
            length, width_mm, load_code = key
            len_r = round(length, 2)
            w_int = int(round(width_mm))
            candidates = [ok for ok in order_keys
                          if abs(round(ok[0], 2) - len_r) <= tol_len and abs(int(round(ok[1])) - w_int) <= tol_w]
            if not candidates:
                return None
            same_load = [ok for ok in candidates if ok[2] == load_code]
            pool = same_load if same_load else candidates
            # Сначала заполняем позиции с меньшим остатком (remaining); при равном rem — с большим need (чтобы не голодал один ключ).
            # Затем по близости длины, затем стабильно по ключу.
            def _rank(ok):
                rem = (remaining.get(ok, order_counts.get(ok, 0)) if remaining is not None else order_counts.get(ok, 0))
                dist = abs(ok[0] - length)
                need = order_counts.get(ok, 0)
                return (RESCUE_EXHAUSTED_RANK if rem <= 0 else rem, -need, dist, ok[0], ok[1], ok[2])
            chosen = min(pool, key=_rank)
            # Лог: несколько кандидатов — выясняем, почему выбран именно этот ключ
            if len(pool) > 1:
                rank_vals = [_rank(ok) for ok in pool]
                _debug_write("H_rescue_norm", "production_execution:_normalize_key_to_orders", "several candidates, chosen",
                             {"input_key": list(key), "candidates": [list(ok) for ok in pool], "chosen": list(chosen),
                              "remaining": {str(k): remaining.get(k, "?") for k in pool} if remaining else {},
                              "rank_values": [(list(ok), r) for ok, r in zip(pool, rank_vals)]})
            return chosen

        def _to_width_mm(w, default_m=1.2):
            """Ширина в мм: если < 20 — считаем метры (×1000), иначе уже мм. При нечисловом w — default_m."""
            if w is None:
                w = default_m
            try:
                f = float(w)
            except (TypeError, ValueError):
                f = default_m
            return round(f * 1000) if f < 20 else round(f)

        def _is_interesting_length(L):
            """Длины, близкие к проблемным: 4.78/4.79, 4.17/4.18, 6.11/6.12."""
            return (4.76 <= L <= 4.82) or (4.15 <= L <= 4.20) or (6.09 <= L <= 6.14)

        def _raw_track_counts(tracks_list):
            """Считает плиты в дорожках по ключу (L, W, load_code) без привязки к заказам — для отладки любых размеров."""
            out = {}
            # #region agent log b80b04 H1: элементы трека с длиной ~6 м и шириной 530 или 1200
            _log_b80 = PROJECT_ROOT / "debug-b80b04.log"
            _interesting_w = (520, 530, 540, 1190, 1200, 1210)
            def _is_6m(L):
                return L is not None and 5.98 <= round(float(L), 2) <= 6.02
            # #endregion
            for track in tracks_list:
                for item in track.get('items', []) or []:
                    if not item:
                        continue
                    length = round(item.get('length', 0) or 0, 2)
                    load_code = cfg.normalize_load_code(item.get('load_code', 8))
                    mode = item.get('mode', 'solid')
                    if mode == 'split':
                        width_mm = _to_width_mm(item.get('main_w', 1.2))
                    else:
                        width_mm = _to_width_mm(item.get('width', 1.2))
                    if mode == 'transverse' and item.get('target_length') is not None:
                        target_length = round(float(item.get('target_length') or 0), 2)
                        if target_length > 0:
                            k = (target_length, width_mm, load_code)
                            out[k] = out.get(k, 0) + 1
                            # #region agent log b80b04 H1
                            if width_mm in _interesting_w and _is_6m(target_length):
                                try:
                                    with open(_log_b80, 'a', encoding='utf-8') as _f:
                                        _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H1", "location": "production_execution:_raw_track_counts:transverse_target", "message": "Item 6m + 530/1200", "data": {"length": length, "width_raw": item.get('width'), "main_w_raw": item.get('main_w'), "mode": mode, "target_length": target_length, "width_mm": width_mm, "key": list(k), "label": item.get('label', '')[:60]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                                except Exception:
                                    pass
                            # #endregion
                        remainder = round(float(item.get('remainder', 0) or 0), 2)
                        if remainder > 0.1:
                            kr = (remainder, width_mm, load_code)
                            out[kr] = out.get(kr, 0) + 1
                            # #region agent log b80b04 H1
                            if width_mm in _interesting_w and _is_6m(remainder):
                                try:
                                    with open(_log_b80, 'a', encoding='utf-8') as _f:
                                        _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H1", "location": "production_execution:_raw_track_counts:transverse_remainder", "message": "Remainder 6m + 530/1200", "data": {"length": length, "width_raw": item.get('width'), "width_mm": width_mm, "remainder": remainder, "key": list(kr)}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                                except Exception:
                                    pass
                            # #endregion
                    else:
                        if length <= 0:
                            continue
                        k = (length, width_mm, load_code)
                        out[k] = out.get(k, 0) + 1
                        # #region agent log b80b04 H1
                        if width_mm in _interesting_w and _is_6m(length):
                            try:
                                with open(_log_b80, 'a', encoding='utf-8') as _f:
                                    _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H1", "location": "production_execution:_raw_track_counts:solid", "message": "Solid 6m + 530/1200", "data": {"length": length, "width_raw": item.get('width'), "main_w_raw": item.get('main_w'), "mode": mode, "width_mm": width_mm, "key": list(k), "label": item.get('label', '')[:60]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                            except Exception:
                                pass
                        # #endregion
                    for sec_cut in item.get('secondary_cuts', []) or []:
                        sw = _to_width_mm(sec_cut.get('width', 0), default_m=0)
                        sl = sec_cut.get('target_length') or length
                        if sw > 0 and sl > 0:
                            sec_lc = cfg.normalize_load_code(sec_cut.get('load_code', item.get('load_code', 8)))
                            sk = (round(float(sl), 2), sw, sec_lc)
                            out[sk] = out.get(sk, 0) + 1
                            # #region agent log b80b04 H1
                            if sw in _interesting_w and _is_6m(sl):
                                try:
                                    with open(_log_b80, 'a', encoding='utf-8') as _f:
                                        _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H1", "location": "production_execution:_raw_track_counts:secondary", "message": "Secondary 6m + 530/1200", "data": {"parent_length": length, "sec_length": sl, "sec_width_mm": sw, "key": list(sk), "parent_label": item.get('label', '')[:60]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                                except Exception:
                                    pass
                            # #endregion
            return out

        def _count_tracks_for_rescue(tracks_list, order_keys, order_counts):
            # BUG-5 FIX: двухпроходная стратегия.
            # Проход 1 — строгое совпадение (tol_len=0, tol_w=0): сначала засчитываем
            # элементы, у которых ключ совпадает точно. Это предотвращает «кражу» плиты
            # у близкого ключа (5.98 → 6.0) раньше, чем будет обработан точный ключ.
            # Проход 2 — нечёткое совпадение: засчитываем оставшиеся элементы с допуском.
            remaining = {k: order_counts.get(k, 0) for k in order_keys}
            counts = {k: 0 for k in order_keys}

            # Собираем плоский список элементов треков для двух проходов
            _all_items = []
            for _tr in tracks_list:
                for _it in (_tr.get('items') or []):
                    if _it:
                        _all_items.append(_it)

            def _process_item(item, tol_len, tol_w):
                """Обрабатывает один элемент трека с заданными допусками.
                Возвращает True если элемент был засчитан."""
                length = round(item.get('length', 0) or 0, 2)
                load_code = cfg.normalize_load_code(item.get('load_code', 8))
                mode = item.get('mode', 'solid')
                if mode == 'split':
                    width_mm = _to_width_mm(item.get('main_w', 1.2))
                elif mode == 'transverse':
                    width_mm = _to_width_mm(item.get('width', 1.2))
                else:
                    width_mm = _to_width_mm(item.get('width', 1.2))
                credited = False

                if mode == 'transverse' and item.get('target_length') is not None:
                    target_length = round(float(item.get('target_length') or 0), 2)
                    if target_length > 0:
                        key_target = (target_length, width_mm, load_code)
                        norm_key = _normalize_key_to_orders(key_target, order_keys, order_counts,
                                                            tol_len=tol_len, tol_w=tol_w, remaining=remaining)
                        if _is_interesting_length(target_length):
                            _debug_write("H_rescue_count", "production_execution:_count_tracks:transverse_target",
                                         "plate assigned", {"target_length": target_length, "length": length,
                                         "width_mm": width_mm, "load_code": load_code, "norm_key": list(norm_key) if norm_key else None,
                                         "tol_len": tol_len,
                                         "assigned": norm_key is not None and norm_key in remaining and remaining.get(norm_key, 0) > 0})
                        if norm_key is not None and norm_key in remaining and remaining[norm_key] > 0:
                            remaining[norm_key] -= 1
                            counts[norm_key] = counts.get(norm_key, 0) + 1
                            credited = True
                    remainder = round(float(item.get('remainder', 0) or 0), 2)
                    if remainder > 0.1:
                        key_rem = (remainder, width_mm, load_code)
                        norm_rem = _normalize_key_to_orders(key_rem, order_keys, order_counts,
                                                            tol_len=tol_len, tol_w=tol_w, remaining=remaining)
                        if norm_rem is not None and norm_rem in remaining and remaining[norm_rem] > 0:
                            remaining[norm_rem] -= 1
                            counts[norm_rem] = counts.get(norm_rem, 0) + 1
                            credited = True
                else:
                    if length <= 0:
                        return False
                    key = (length, width_mm, load_code)
                    norm_key = _normalize_key_to_orders(key, order_keys, order_counts,
                                                        tol_len=tol_len, tol_w=tol_w, remaining=remaining)
                    if _is_interesting_length(length):
                        _debug_write("H_rescue_count", "production_execution:_count_tracks:solid",
                                     "plate assigned", {"length": length, "width_mm": width_mm, "load_code": load_code,
                                     "norm_key": list(norm_key) if norm_key else None,
                                     "tol_len": tol_len,
                                     "assigned": norm_key is not None and norm_key in remaining and remaining.get(norm_key, 0) > 0})
                    if norm_key is not None and norm_key in remaining and remaining[norm_key] > 0:
                        remaining[norm_key] -= 1
                        counts[norm_key] = counts.get(norm_key, 0) + 1
                        credited = True
                return credited

            # Проход 1: строгое совпадение (tol_len=0, tol_w=0)
            # Элементы, не засчитанные строго, переходят в проход 2
            _unmatched = []
            for _item in _all_items:
                _credited = _process_item(_item, tol_len=0, tol_w=0)
                if not _credited:
                    _unmatched.append(_item)

            # Проход 2: нечёткое совпадение для незасчитанных элементов
            for _item in _unmatched:
                _process_item(_item, tol_len=RESCUE_TOL_LEN_M, tol_w=RESCUE_TOL_W_MM)

            # Secondary cuts: обрабатываем отдельно — у них есть target_order_key (BUG-4)
            # Сначала точные по target_order_key, потом fuzzy для оставшихся
            for item in _all_items:
                item_length = round(item.get('length', 0) or 0, 2)
                for sec_cut in (item.get('secondary_cuts') or []):
                    sec_width_mm = _to_width_mm(sec_cut.get('width', 0), default_m=0)
                    sec_length = sec_cut.get('target_length') or item_length
                    if sec_width_mm <= 0 or sec_length <= 0:
                        continue
                    sec_load_code = cfg.normalize_load_code(sec_cut.get('load_code', item.get('load_code', 8)))
                    # BUG-4 FIX: если есть точный target_order_key от оптимизатора,
                    # используем его напрямую вместо fuzzy-поиска по ключу трека
                    _tok = sec_cut.get('target_order_key')
                    if _tok and len(_tok) == 3:
                        _tok_canon = canonical_plate_key(_tok[0], _tok[1], _tok[2])
                        if _tok_canon in remaining and remaining[_tok_canon] > 0:
                            remaining[_tok_canon] -= 1
                            counts[_tok_canon] = counts.get(_tok_canon, 0) + 1
                            continue
                    # Fallback: fuzzy match
                    sec_key = (round(sec_length, 2), sec_width_mm, sec_load_code)
                    sec_norm = _normalize_key_to_orders(sec_key, order_keys, order_counts, remaining=remaining)
                    if sec_norm is not None and sec_norm in remaining and remaining[sec_norm] > 0:
                        remaining[sec_norm] -= 1
                        counts[sec_norm] = counts.get(sec_norm, 0) + 1
            return counts

        def _merge_to_canonical_order_keys(raw_order_counts, tol_len=0.02):
            """Объединяет ключи заказов: длины в пределах tol_len считаются одним ключом (минимум по длине).
            Чтобы плиты из БД (3.8) и из раскладки (3.79) не дублировались в РЕСКЬЮ и все списывались."""
            if not raw_order_counts:
                return {}, lambda k: k
            keys_list = list(raw_order_counts.keys())
            def canonical_key(key):
                L, W, LC = key
                best_L = L
                for (L2, W2, LC2) in keys_list:
                    if W2 == W and LC2 == LC and abs(L2 - L) <= tol_len and L2 < best_L:
                        best_L = L2
                return (best_L, W, LC)
            merged = {}
            for k, qty in raw_order_counts.items():
                ck = canonical_key(k)
                merged[ck] = merged.get(ck, 0) + qty
            return merged, canonical_key

        def _build_order_info_map(orders_list, canonical_key_fn=None):
            info_map = {}
            for order in orders_list:
                L = round(float(order.get('length', 0)), 2)
                W = order.get('width', 1200)
                W_canon = int(round(float(W))) if W is not None else 1200
                key = (L, W_canon, cfg.normalize_load_code(order.get('load_code', 8)))
                if canonical_key_fn:
                    key = canonical_key_fn(key)
                if key not in info_map:
                    info_map[key] = []
                info_map[key].append({
                    'kp_id': order.get('kp_id'),
                    'customer': order.get('customer'),
                    'kp_date': order.get('kp_date'),
                    'plate_name': order.get('plate_name', ''),
                    'qty_remaining': order.get('qty', 1),
                    'length_dm_raw': order.get('length_dm_raw', '') or '',
                })
            return info_map

        def _create_rescue_tracks(missing_counts, info_map):
            rescue_tracks = []
            current_track = []
            current_len = 0.0
            max_len = 101.0

            def _flush_track():
                nonlocal current_track, current_len
                if current_track:
                    rescue_tracks.append({
                        'items': current_track,
                        'length': current_len,
                        'load_code': 0,
                        'label': 'РЕСКЬЮ',
                        'max_reinforcement': 0.0
                    })
                current_track = []
                current_len = 0.0

            for key, qty_missing in missing_counts.items():
                length, width_mm, load_code = key
                width_m = width_mm / 1000.0
                for _ in range(qty_missing):
                    if current_track and current_len + length > max_len:
                        _flush_track()
                    # Достаём информацию о КП (если есть)
                    order_info = None
                    for entry in info_map.get(key, []):
                        if entry.get('qty_remaining', 0) > 0:
                            entry['qty_remaining'] -= 1
                            order_info = entry
                            break
                    if order_info and order_info.get('plate_name'):
                        plate_name = order_info['plate_name']
                        display_name = plate_name
                        rescue_order_missing = False
                    else:
                        display_name = cfg.make_plate_name(
                            length, width_m, load_code=load_code,
                            length_dm_raw=order_info.get('length_dm_raw') if order_info else None
                        )
                        plate_name = None
                        rescue_order_missing = True
                        logger.warning(
                            "[RESCUE] Не найден exact order_info для ключа %s. "
                            "Создаю rescue-элемент только для визуализации с label=%s",
                            key,
                            display_name,
                        )
                    current_track.append({
                        'length': length,
                        'mode': 'solid',
                        'width': width_m,
                        'load_code': load_code,
                        'label': display_name,
                        'reinforcement': 0,
                        'kp_id': order_info.get('kp_id') if order_info else None,
                        'customer': order_info.get('customer') if order_info else None,
                        'kp_date': order_info.get('kp_date') if order_info else None,
                        'plate_name': plate_name,
                        'rescue_order_missing': rescue_order_missing,
                    })
                    current_len += length
            _flush_track()
            return rescue_tracks

        # Канонизация: длина round(.,2), ширина int — чтобы не было расхождения float/int в dict.
        # BUG-6 FIX: tol_len=0.02 (вместо 0) позволяет слить ключи, расходящиеся на ≤2см
        # из-за округления в парсере или оптимизаторе (напр. 3.79 vs 3.8 из БД).
        # 5.7 и 5.71 как ПБ 57 / ПБ 57,1 остаются разными — разница 0.01 < 0.02 означает
        # слияние, что корректно: оба попадут под одну rescue-проверку суммарным qty.
        raw_order_counts = {}
        for order in orders_2d:
            key = canonical_plate_key(
                order.get('length', 0),
                order.get('width', 1200),
                order.get('load_code', 8),
            )
            raw_order_counts[key] = raw_order_counts.get(key, 0) + order.get('qty', 1)
            # #region agent log b80b04 H3: заказы 60-12 и 60-5,3 — ключ и width в orders_2d
            _pn = (order.get('plate_name') or '') or ''
            if '60-12' in _pn or '60-5,3' in _pn or '51-3,2' in _pn:
                try:
                    _log_b80 = PROJECT_ROOT / "debug-b80b04.log"
                    with open(_log_b80, 'a', encoding='utf-8') as _f:
                        _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H3", "location": "production_execution:raw_order_counts", "message": "Order line 60-12 / 60-5,3 / 51-3,2", "data": {"plate_name": _pn[:80], "length": L, "width_raw": W, "W_canon": W_canon, "key": list(key), "qty": order.get('qty', 1)}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
        order_counts, canonical_key_fn = _merge_to_canonical_order_keys(raw_order_counts, tol_len=0.02)
        order_keys = list(order_counts.keys())
        competing = []
        for i, k1 in enumerate(order_keys):
            for k2 in order_keys[i+1:]:
                if k1[1] == k2[1] and k1[2] == k2[2] and 0 < abs(k1[0] - k2[0]) <= 0.03:
                    competing.append((list(k1), list(k2), order_counts.get(k1, 0), order_counts.get(k2, 0)))
        _debug_write("H_rescue_keys", "production_execution:before_count", "raw_order_counts, order_keys, competing",
                     {"raw_order_counts": {str(list(k)): v for k, v in raw_order_counts.items()},
                      "order_keys": [list(k) for k in order_keys],
                      "order_counts": {str(list(k)): v for k, v in order_counts.items()},
                      "competing_keys": competing})
        track_counts = _count_tracks_for_rescue(all_tracks_list, order_keys, order_counts)
        raw_track_by_key = _raw_track_counts(all_tracks_list)
        _debug_write("H_rescue_tracks_raw", "production_execution:after_count", "track items by (length, width, load_code) before normalization",
                     {"raw_by_key": {str(list(k)): v for k, v in raw_track_by_key.items()}, "total": sum(raw_track_by_key.values())})
        missing_counts = {}
        for key, qty_need in order_counts.items():
            qty_have = track_counts.get(key, 0)
            if qty_have < qty_need:
                missing_counts[key] = qty_need - qty_have
        # #region agent log
        _debug_runtime_write(
            "run1",
            "H4_missing_calc",
            "production_execution:missing_counts",
            "Need/have/raw before rescue creation",
            {
                "order_counts": {str(list(k)): int(v) for k, v in order_counts.items()},
                "track_counts_norm": {str(list(k)): int(v) for k, v in track_counts.items()},
                "track_counts_raw": {str(list(k)): int(v) for k, v in raw_track_by_key.items()},
                "missing_counts": {str(list(k)): int(v) for k, v in missing_counts.items()},
            },
        )
        # #endregion
        # Защита от расхождения «в боте 60-12 vs в визуализации 60-5,3»: если по raw в треках
        # не хватает (6, 1200, 8), но недостача попала в (6, 530, 8) — перераспределяем одну
        # недостачу на ключ (6, 1200, 8), чтобы в РЕСКЬЮ добавилась плита 60-12-8п.
        _key_6_530 = (6.0, 530, 8)
        _key_6_1200 = (6.0, 1200, 8)
        need_1200 = order_counts.get(_key_6_1200, 0)
        raw_1200 = raw_track_by_key.get(_key_6_1200, 0)
        missing_530 = missing_counts.get(_key_6_530, 0)
        if need_1200 > raw_1200 and missing_530 > 0:
            transfer = min(missing_530, need_1200 - raw_1200)
            missing_counts[_key_6_530] = missing_530 - transfer
            if missing_counts[_key_6_530] == 0:
                del missing_counts[_key_6_530]
            missing_counts[_key_6_1200] = missing_counts.get(_key_6_1200, 0) + transfer
        # #region agent log b80b04 H2/H5: need vs have vs raw для (6,530,8) и (6,1200,8)
        try:
            _log_b80 = PROJECT_ROOT / "debug-b80b04.log"
            _keys_rescue = [(6.0, 530, 8), (6.0, 1200, 8)]
            _trace = []
            for _k in _keys_rescue:
                _trace.append({"key": list(_k), "need": order_counts.get(_k, 0), "have_norm": track_counts.get(_k, 0), "raw_tracks": raw_track_by_key.get(_k, 0), "missing": missing_counts.get(_k, 0)})
            with open(_log_b80, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "b80b04", "hypothesisId": "H2_H5", "location": "production_execution:missing_counts", "message": "6x530 vs 6x1200 need/have/raw/missing", "data": {"trace": _trace, "total_missing": sum(missing_counts.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (2d5c43) H1-H4: need/have/raw before RESCUE, primary_cuts and track totals
        try:
            _log_2d5c43 = PROJECT_ROOT / "debug-2d5c43.log"
            _target_keys_2d5c43 = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
            _trace_2d5c43 = [{"key": list(k), "need": order_counts.get(k, 0), "have_norm": track_counts.get(k, 0), "raw_tracks": raw_track_by_key.get(k, 0), "missing": missing_counts.get(k, 0)} for k in _target_keys_2d5c43]
            _prim_count = len(optimization_result.get('primary_cuts', []))
            _track_items_total = sum(len(t.get('items', [])) for t in all_tracks_list)
            with open(_log_2d5c43, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "2d5c43", "hypothesisId": "H1_H4", "location": "production_execution:before_rescue", "message": "need/have/raw and totals", "data": {"trace": _trace_2d5c43, "total_missing": sum(missing_counts.values()), "primary_cuts_count": _prim_count, "track_items_total": _track_items_total, "tracks_count": len(all_tracks_list)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (73b708) H_WHERE_RESCUE: need / have (normalized) / raw для целевых ключей перед rescue
        try:
            _trace_keys = [(5.08, 320, 8), (5.98, 665, 8)]
            _rescue_trace = []
            for _tk in _trace_keys:
                _rescue_trace.append({
                    "key": list(_tk),
                    "need": order_counts.get(_tk, 0),
                    "have_normalized": track_counts.get(_tk, 0),
                    "raw_in_tracks": raw_track_by_key.get(_tk, 0),
                    "missing": missing_counts.get(_tk, 0),
                })
            _agent_log = PROJECT_ROOT / "debug-73b708.log"
            with open(_agent_log, 'a', encoding='utf-8') as _f:
                _f.write(json.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_WHERE_RESCUE", "location": "production_execution:before_rescue", "message": "Need vs have vs raw for target keys", "data": {"trace": _rescue_trace}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log
        _target_keys = {(5.98, 1200, 8), (5.08, 530, 8), (5.08, 320, 8)}
        _target_missing = []
        for _k in _target_keys:
            _target_missing.append({
                "key": list(_k),
                "need": int(order_counts.get(_k, 0) or 0),
                "have": int(track_counts.get(_k, 0) or 0),
                "missing": int(missing_counts.get(_k, 0) or 0),
                "raw_track_this_key": int(raw_track_by_key.get(_k, 0) or 0),
            })
        _debug_session_write(
            "run1",
            "H2",
            "production_execution:missing_counts",
            "Need/have/missing for target keys before RESCUE",
            {
                "target_missing": _target_missing,
                "total_missing_all_keys": int(sum(missing_counts.values()) if missing_counts else 0),
            },
        )
        # #endregion
        interesting_missing = {str(list(k)): (order_counts.get(k, 0), track_counts.get(k, 0), v)
                              for k, v in missing_counts.items() if _is_interesting_length(k[0])}
        _debug_write("H_rescue_result", "production_execution:after_count", "need/have/missing for interesting keys",
                     {"interesting_missing": interesting_missing, "total_missing": sum(missing_counts.values())})
        # Полный список недостающих + сырые счётчики по ключу и по (L,W) — для отладки любых размеров.
        all_missing = {}
        for k, v in missing_counts.items():
            need = order_counts.get(k, 0)
            have = track_counts.get(k, 0)
            raw_this = raw_track_by_key.get(k, 0)
            L, W, LC = k
            raw_same_lw = sum(q for (l2, w2, _), q in raw_track_by_key.items() if abs(l2 - L) <= 0.02 and abs(w2 - W) <= 25)
            all_missing[str(list(k))] = {
                "need": need, "have": have, "missing": v,
                "raw_this_key": raw_this,
                "raw_same_LW": raw_same_lw
            }
        _debug_write("H_rescue_all_missing", "production_execution:after_count", "all missing keys need/have/missing + raw track counts",
                     {"all_missing": all_missing, "total_missing": sum(missing_counts.values())})
        # Для проверки допуска 10 мм: смотри в logs/bot.log — покрытие заказа дорожками и недостаток
        logger.info(
            f"[CHECK] РЕСКЬЮ: позиций в заказе {len(order_counts)}, "
            f"в плане покрыто {sum(track_counts.values())} плит, недостаёт {sum(missing_counts.values())} (непокрытых позиций: {len(missing_counts)})"
        )
        rescue_tracks_added = 0
        if missing_counts:
            info_map = _build_order_info_map(orders_2d, canonical_key_fn=canonical_key_fn)
            rescue_tracks = _create_rescue_tracks(missing_counts, info_map)
            all_tracks_list.extend(rescue_tracks)
            rescue_tracks_added = len(rescue_tracks)
            # #region agent log
            _rescue_missing_exact = []
            for _tr in rescue_tracks:
                for _it in _tr.get('items', []) or []:
                    if _it.get('rescue_order_missing'):
                        _rescue_missing_exact.append({
                            "label": _it.get('label'),
                            "length": _it.get('length'),
                            "width": _it.get('width'),
                            "load_code": _it.get('load_code'),
                            "kp_id": _it.get('kp_id'),
                        })
            _debug_runtime_write(
                "run1",
                "H5_rescue_fallback",
                "production_execution:rescue_tracks_created",
                "Rescue tracks with missing exact order binding",
                {
                    "rescue_tracks_count": len(rescue_tracks),
                    "rescue_items_without_exact_order_count": len(_rescue_missing_exact),
                    "rescue_items_without_exact_order_sample": _rescue_missing_exact[:30],
                },
            )
            # #endregion
            # #region agent log
            _rescue_targets = []
            for _tr in rescue_tracks:
                for _it in _tr.get("items", []) or []:
                    _n = (_it.get("plate_name") or _it.get("label") or "")
                    if any(_k in _n for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п")):
                        _rescue_targets.append({
                            "plate_name": _n,
                            "length": _it.get("length"),
                            "width": _it.get("width"),
                            "load_code": _it.get("load_code"),
                            "kp_id": _it.get("kp_id"),
                        })
            _debug_session_write(
                "run1",
                "H3",
                "production_execution:rescue_tracks_created",
                "Target plates present in generated RESCUE tracks",
                {
                    "rescue_tracks_count": len(rescue_tracks),
                    "target_rescue_entries": _rescue_targets,
                    "target_rescue_count": len(_rescue_targets),
                },
            )
            # #endregion
            logger.warning(
                f"[RESCUE] Добавлены дополнительные дорожки: {len(rescue_tracks)}. "
                f"Потеряно плит: {sum(missing_counts.values())}"
            )

        # PlateAudit: финальный checkpoint после rescue
        _handler_audit.checkpoint("final", all_tracks_list)
        if _handler_audit.has_losses("input", "final"):
            logger.error("[AUDIT] Итоговые потери плит:\n%s", _handler_audit.summary())
        else:
            logger.info("[AUDIT] Все плиты учтены:\n%s", _handler_audit.summary())

        # #region agent log (session 73b708) H_E: резюме недостачи и рескью + какие заказы дали недостающие ключи
        try:
            _agent_log = PROJECT_ROOT / "debug-73b708.log"
            _total_missing = sum(missing_counts.values()) if missing_counts else 0
            _missing_sample = [list(k) + [v] for k, v in list(missing_counts.items())[:10]] if missing_counts else []
            _missing_to_orders = []
            _log_err = None
            if missing_counts:
                try:
                    for mk in missing_counts.keys():
                        _ml, _mw, _mlc = mk
                        _orders_for_key = [{"plate_name": str(o.get("plate_name") or ""), "length": float(o.get("length", 0)), "width": int(round(float(o.get("width", 1200)))), "qty": int(o.get("qty", 1))} for o in orders_2d if abs(round(float(o.get("length", 0)), 2) - _ml) < 0.02 and int(round(float(o.get("width", 1200)))) == _mw and cfg.normalize_load_code(o.get("load_code", 8)) == _mlc]
                        _missing_to_orders.append({"key": list(mk), "orders": _orders_for_key[:5]})
                except Exception as _e:
                    _log_err = str(_e)
            _payload = {"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_E", "location": "production_execution:rescue_summary", "message": "Rescue/missing summary", "data": {"total_missing": _total_missing, "missing_keys_count": len(missing_counts) if missing_counts else 0, "missing_sample": _missing_sample, "rescue_tracks_added": rescue_tracks_added, "missing_key_to_orders": _missing_to_orders}, "timestamp": __import__('time').time()}
            if _log_err:
                _payload["data"]["missing_key_to_orders_error"] = _log_err
            with open(_agent_log, 'a', encoding='utf-8') as _fe:
                _fe.write(json.dumps(_payload, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        total_tracks_count = len(all_tracks_list)
        total_days = math.ceil(total_tracks_count / tracks_count)
        
        # Проверка: все плиты плана есть в БД (этап 3 — не придумываем плиты)
        plan_plates = Counter(_raw_track_counts(all_tracks_list))
        kp_ids_in_production = list({p.get('kp_id') for p in selected_plates if p.get('kp_id')})
        if kp_ids_in_production and plan_plates:
            try:
                db_plates = Counter()
                with kp_db._connect(kp_db.DEFAULT_DB) as conn:
                    cur = conn.cursor()
                    placeholders = ','.join('?' * len(kp_ids_in_production))
                    cur.execute(f"""
                        SELECT length_m, width_m, load_class, SUM(qty)
                        FROM kp_plates
                        WHERE kp_id IN ({placeholders})
                        AND status IN ('в производстве', 'в плане')
                        GROUP BY length_m, width_m, load_class
                    """, kp_ids_in_production)
                    for row in cur.fetchall():
                        length_m, width_m, load_class, qty = row
                        lc = cfg.normalize_load_code(load_class // 100)
                        width_mm = int(round(width_m * 1000))
                        db_plates[(round(length_m, 2), width_mm, lc)] += qty
                extra_in_plan = plan_plates - db_plates
                missing_in_plan = db_plates - plan_plates
                if extra_in_plan:
                    logger.error(f"[ПЛАН] В плане есть плиты, которых нет в БД (придуманные): {dict(extra_in_plan)}")
                if missing_in_plan:
                    logger.warning(f"[ПЛАН] В БД есть плиты, не попавшие в план (остатки?): {dict(missing_in_plan)}")
            except Exception as e:
                logger.exception(f"[ПЛАН] Проверка план vs БД: {e}")
        
        # === СОХРАНЯЕМ ВСЕ ДАННЫЕ ===
        target_date_str = data.get('target_date')
        plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
        completed_days = data.get('completed_days', [])
        
        # Форматируем дату начала для читаемости
        plan_start_display = plan_start_date
        try:
            plan_start_display = datetime.strptime(plan_start_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            pass
        
        await message.answer(
            f"✅ План готов!\n\n"
            f"📊 Параметры:\n"
            f"  • Дата начала: {plan_start_display}\n"
            f"  • Всего дорожек: {total_tracks_count}\n"
            f"  • Дорожек в день: {tracks_count}\n"
            f"  • Потребуется дней: {total_days}\n\n"
            f"💡 Что дальше?\n"
            f"1️⃣ Просмотрите дни ниже 👇\n"
            f"2️⃣ Чтобы посмотреть диаграмму ДО сохранения — нажмите «📈 Диаграмма этого плана»\n"
            f"3️⃣ Нажмите «💾 Сохранить план» когда всё готово\n\n"
            f"⚠️ ВАЖНО: План сохраняется только после нажатия кнопки!\n"
            f"Без сохранения он останется только в памяти.\n\n"
            f"«📈 Диаграмма этого плана» — по текущему расчёту (даже без сохранения)."
        )
        
        # Рассчитываем days_info с глобальной загруженностью для новых дат
        global_occupancy = get_global_day_occupancy()
        
        # Создаём days_info для каждого дня нового плана
        days_info = {}
        try:
            start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
        except:
            start_dt = datetime.now()
        
        # Проверяем, не превышает ли план доступные слоты
        overloaded_days = []  # Дни, где будет превышение
        
        for day_num in range(1, total_days + 1):
            day_date = start_dt + timedelta(days=day_num - 1)
            date_key = day_date.strftime('%Y-%m-%d')
            date_display = day_date.strftime('%d.%m')
            
            current_occupied = global_occupancy.get(date_key, 0)
            free_slots = MAX_TRACKS_PER_DAY - current_occupied
            
            # Проверяем превышение
            if tracks_count > free_slots:
                overloaded_days.append({
                    'date': date_display,
                    'occupied': current_occupied,
                    'free': free_slots,
                    'want': tracks_count,
                    'excess': tracks_count - free_slots
                })
            
            days_info[date_key] = {
                'occupied': current_occupied,
                'max': MAX_TRACKS_PER_DAY,
                'completed': False,
                'day_number': day_num
            }
        
        # Если есть превышение - показываем предупреждение
        if overloaded_days:
            warning_lines = ["⚠️ ВНИМАНИЕ! Превышение лимита дорожек!\n"]
            warning_lines.append(f"Вы хотите планировать {tracks_count} дорожек/день,")
            warning_lines.append(f"но на некоторых датах не хватает места:\n")
            
            for day in overloaded_days[:5]:  # Показываем максимум 5 проблемных дней
                warning_lines.append(
                    f"  • {day['date']}: занято {day['occupied']}/5, "
                    f"свободно {day['free']}, нужно {day['want']}"
                )
            
            if len(overloaded_days) > 5:
                warning_lines.append(f"  ... и ещё {len(overloaded_days) - 5} дней")
            
            warning_lines.append(f"\n💡 Решения:")
            warning_lines.append(f"1️⃣ Уменьшите количество дорожек в день")
            warning_lines.append(f"2️⃣ Выберите другую дату начала")
            warning_lines.append(f"3️⃣ Удалите или отредактируйте другие планы")
            
            await message.answer('\n'.join(warning_lines))
        
        await state.update_data(
            total_tracks_count=total_tracks_count,
            total_days=total_days,
            tracks_count=tracks_count,
            all_tracks_list=all_tracks_list,
            orders_2d=orders_2d,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            plate_to_kp_info=plate_to_kp_info,
            optimization_result=optimization_result,
            target_date=target_date_str,
            plates_from_rests=plates_from_rests,
            plan_start_date=plan_start_date,
            completed_days=completed_days,
            days_info=days_info
        )
        
        # === ПОКАЗЫВАЕМ КНОПКИ С ДАТАМИ ===
        await message.answer(
            "Выберите день производства:",
            reply_markup=calendar_days_kb(
                total_days, 
                plan_start_date, 
                completed_days,
                days_info
            )
        )
        
        await state.set_state(ProductionStates.waiting_day_selection)
        
    except Exception as e:
        logger.exception(f"Ошибка при планировании производства: {e}")
        await message.answer(
            "❌ Ошибка при планировании производства.\n\n"
            "Попробуйте позже. Если повторяется — смотри logs/bot.log.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
