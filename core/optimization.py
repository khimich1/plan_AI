#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль оптимизации раскроя плит:
- Оптимизация резов по ширине (PuLP)
- Оптимизация с учётом длин
- Полная оптимизация с narrowing
- Каскадные продольные резы (вторичное использование остатков)

Терминология:
- ПРОДОЛЬНЫЙ РЕЗ: режет вдоль длины, уменьшает ШИРИНУ (1.2м → 0.32м + 0.88м)
- ПОПЕРЕЧНЫЙ РЕЗ: режет поперёк, уменьшает ДЛИНУ (6.0м → 3.0м + 3.0м)
"""
# Относительные импорты внутри core/
from . import config_and_data as cfg
from .price_db import get_price
from dataclasses import dataclass


# ==================== КОНФИГУРАЦИЯ ОПТИМИЗАЦИИ ====================

# Ширина пропила (в мм) - НЕ используется в расчётах
# Пропил косвенно учитывается в таблице NARROWING_TABLE через значения narrowing
# Формально для совместимости оставляем константу, но не применяем в расчётах
KERF_WIDTH_MM = 0

@dataclass
class OptimizationConfig:
    """
    Конфигурация параметров оптимизации.
    Позволяет экспериментировать с разными коэффициентами штрафов и бонусов.
    """
    # Коэффициент штрафа за неиспользованные остатки
    # OLD: 0.5 (50% стоимости остатка)
    # NEW: 0.15 (15% стоимости остатка)
    unused_rest_penalty_coeff: float = 0.15
    
    # Бонус за использование вторичных резов (отрицательное значение = бонус)
    # OLD: -500 (экономический стимул использовать остатки)
    # NEW: 0 (нет бонуса, остатки используются только если это выгодно)
    secondary_reuse_bonus: float = 0.0


# Дефолтная конфигурация (NEW поведение)
DEFAULT_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.15,
    secondary_reuse_bonus=0.0
)

# Старая конфигурация (OLD поведение, для экспериментов)
OLD_CONFIG = OptimizationConfig(
    unused_rest_penalty_coeff=0.5,
    secondary_reuse_bonus=-500.0
)

# ==================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ОПТИМИЗАЦИИ ====================

OPT_WIDTH_PRIORITY = []  # приоритет ширин: ['0_32','0_46','0_70','0_72','0_86']
OPT_PLAN = {}  # результат полной оптимизации: как закрывать спрос
OPT_CASCADING_PLAN = {}  # результат каскадной оптимизации с вторичными резами
OPT_CASCADING_PLAN_BY_LOAD = {}  # результат каскадной оптимизации, СГРУППИРОВАННЫЙ ПО НАГРУЗКЕ
LOAD_TO_REINFORCEMENT_MAP = {}  # маппинг: load_code → [reinforcement_keys] для поиска плана по нагрузке


# ==================== ЛЕГАСИ-АДАПТЕРЫ ====================
from collections import Counter


def _group_plate_lengths(plates: list[float]) -> dict[float, int]:
    """Группируем длины (в метрах) -> количество."""
    return Counter(round(float(length), 2) for length in (plates or []))


def _append_actions(actions: list, width_mm: int, lengths: dict, long_cuts: int, src_type: str):
    """Добавляем aggregated записи в OPT_PLAN['actions']."""
    rest_mm = max(0, 1200 - width_mm)
    for length_m, qty in lengths.items():
        actions.append((
            src_type,
            width_mm,
            rest_mm if src_type != 'solid' else 0,
            length_m,
            qty,
            long_cuts,
            0  # поперечные резы в legacy-плане не моделируем
        ))


def apply_width_optimization() -> dict:
    """
    Упрощённый наследуемый оптимизатор ширин.
    Наполняет OPT_WIDTH_PRIORITY и OPT_PLAN['actions'] данными из cfg.PLATES_*.
    """
    priority = []
    actions = []

    priority_groups = [
        ('0_32', 320, cfg.PLATES_0_32),
        ('0_46', 460, cfg.PLATES_0_46),
        ('0_70', 700, cfg.PLATES_0_70),
        ('0_72', 720, cfg.PLATES_0_72),
        ('0_86', 860, cfg.PLATES_0_86),
        ('0_74', 740, cfg.PLATES_0_74),
        ('0_88', 880, cfg.PLATES_0_88),
        ('0_48', 480, cfg.PLATES_0_48),
        ('0_50', 500, cfg.PLATES_0_50),
        ('0_34', 340, cfg.PLATES_0_34),
    ]

    for code, width_mm, plate_list in priority_groups:
        if not plate_list:
            continue
        priority.append(code)
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    solid_groups = [
        (1200, cfg.PLATES_1_2),
    ]
    split_groups = [
        (1080, cfg.PLATES_1_08),
        (1000, cfg.PLATES_1_0),
    ]

    for width_mm, plate_list in solid_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=0, src_type='solid')

    for width_mm, plate_list in split_groups:
        if not plate_list:
            continue
        lengths = _group_plate_lengths(plate_list)
        _append_actions(actions, width_mm, lengths, long_cuts=1, src_type='split')

    OPT_WIDTH_PRIORITY.clear()
    OPT_WIDTH_PRIORITY.extend(priority)

    OPT_PLAN.clear()
    OPT_PLAN.update({
        'orders': {width: len(lst) for (_, width, lst) in priority_groups if lst},
        'actions': actions
    })

    return OPT_PLAN


def optimize_cuts_pulp(orders: dict | None = None) -> dict:
    """
    Legacy-обёртка над новой каскадной оптимизацией (для совместимости с визуализатором).
    """
    if orders is None:
        orders = {}
        for width_mm, plates in [
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34)
        ]:
            if plates:
                orders[width_mm] = len(plates)

    result = optimize_with_cascading_longitudinal_cuts(orders=orders) if orders else {}
    if result:
        OPT_PLAN.clear()
        OPT_PLAN.update(result)
    return result


# ==================== СОВРЕМЕННЫЕ ФУНКЦИИ ОПТИМИЗАЦИИ ====================

def _get_next_order_info(order_info_list: dict, key: tuple) -> dict:
    """
    Возвращает информацию о следующем КП с qty_remaining > 0 и уменьшает счётчик.
    
    Простыми словами:
    - Ищет в списке записей для данного (length, width, load_code) первую запись, 
      у которой ещё есть неназначенные плиты (qty_remaining > 0)
    - Уменьшает счётчик qty_remaining на 1
    - Возвращает копию информации о КП
    
    Args:
        order_info_list: словарь {(length, width, load_code): [список записей КП]}
        key: кортеж (length, width, load_code)
    
    Returns:
        dict: информация о КП (kp_id, customer, kp_date, plate_name, load_code) или пустой словарь
    """
    entries = order_info_list.get(key, [])
    for entry in entries:
        if entry.get('qty_remaining', 0) > 0:
            entry['qty_remaining'] -= 1
            # Возвращаем копию без qty_remaining (он служебный)
            return {
                'kp_id': entry.get('kp_id'),
                'customer': entry.get('customer'),
                'kp_date': entry.get('kp_date'),
                'plate_name': entry.get('plate_name'),
                'load_code': entry.get('load_code'),
                'reinforcement': entry.get('reinforcement')
            }
    # Fallback по (length, width) без load_code — ищем любой ключ с теми же длиной и шириной
    if len(key) == 3:
        length, width, load_code = key
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) >= 2 and candidate_key[0] == length and candidate_key[1] == width:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        try:
                            _req_lc = key[2] if len(key) >= 3 else None
                            _found_lc = candidate_key[2] if len(candidate_key) >= 3 else None
                            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f:
                                _f.write(__import__('json').dumps({
                                    "hypothesisId": "H2_fallback",
                                    "location": "optimization.py:_get_next_order_info",
                                    "message": "fallback used (length, width)",
                                    "data": {
                                        "requested_key": list(key),
                                        "found_key": list(candidate_key),
                                        "fallback_reason": "load_code_mismatch",
                                        "requested_load_code": _req_lc,
                                        "found_load_code": _found_lc,
                                        "kp_id": entry.get('kp_id'),
                                        "plate_name": (entry.get('plate_name') or '')[:50]
                                    },
                                    "timestamp": __import__('time').time()
                                }, ensure_ascii=False) + '\n')
                        except Exception:
                            pass
                        return {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement')
                        }
        # Fallback по «соседней» длине (±0.02 м), та же ширина и load_code (61,2↔61,1; 59,8↔59,9)
        # Иначе при конкурирующих длинах решатель даёт общий объём, список по точной длине кончается —
        # плиты получают kp_id из opt (первый КП), а в БД они в другом КП и не списываются.
        LEN_TOL = 0.02
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) < 3:
                continue
            c_len, c_width, c_lc = candidate_key[0], candidate_key[1], candidate_key[2]
            if abs(c_len - length) <= LEN_TOL and c_width == width and c_lc == load_code:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement')
                        }
    return {}


def _build_proportional_slot_lists(
    orders_2d: list,
    demand_2d: dict,
) -> tuple:
    """
    Строит пропорциональные слоты атрибуции по ключу (length, width, load_code).
    Возвращает (slot_lists, slot_cursors).
    slot_lists[key] — список из demand_2d[key] атрибуций, пропорционально qty заказов
    (floor + остаток по убыванию qty). Курсоры инициализированы в 0.
    """
    groups: dict = {}
    for order in orders_2d:
        key = (order['length'], order['width'], order.get('load_code', 800))
        groups.setdefault(key, []).append(order)

    slot_lists: dict = {}
    for key, need in demand_2d.items():
        entries = groups.get(key, [])
        total_qty = sum(o.get('qty', 1) for o in entries)
        if total_qty == 0:
            slot_lists[key] = []
            continue
        shares = [int(need * o.get('qty', 1) / total_qty) for o in entries]
        remainder = need - sum(shares)
        for idx in sorted(range(len(entries)), key=lambda i: -entries[i].get('qty', 1)):
            if remainder <= 0:
                break
            shares[idx] += 1
            remainder -= 1
        slots = []
        for entry, share in zip(entries, shares):
            info = {
                k: entry.get(k)
                for k in ('kp_id', 'customer', 'kp_date', 'plate_name', 'load_code', 'reinforcement')
            }
            slots.extend([info] * share)
        slot_lists[key] = slots

    cursors = {key: 0 for key in slot_lists}
    return slot_lists, cursors


def _next_slot_info(
    slot_lists: dict,
    slot_cursors: dict,
    key: tuple,
) -> dict:
    """
    Возвращает следующую атрибуцию по ключу из предрасчитанных слотов и сдвигает курсор.
    При исчерпании слотов (LP over-deliver) возвращает последний слот.
    """
    slots = slot_lists.get(key, [])
    if not slots:
        return {}
    idx = slot_cursors.get(key, 0)
    entry = slots[min(idx, len(slots) - 1)]
    slot_cursors[key] = idx + 1
    return entry


def _peek_order_info(order_info_list: dict, key: tuple) -> dict:
    """
    Возвращает информацию о первом КП с qty_remaining > 0 БЕЗ уменьшения счётчика.
    
    Используется для получения информации при создании primary_options,
    когда ещё не известно, будет ли опция использована.
    
    Args:
        order_info_list: словарь {(length, width, load_code): [список записей КП]}
        key: кортеж (length, width, load_code)
    
    Returns:
        dict: информация о КП (включая load_code) или пустой словарь
    """
    entries = order_info_list.get(key, [])
    for entry in entries:
        if entry.get('qty_remaining', 0) > 0:
            return {
                'kp_id': entry.get('kp_id'),
                'customer': entry.get('customer'),
                'kp_date': entry.get('kp_date'),
                'plate_name': entry.get('plate_name'),
                'load_code': entry.get('load_code'),
                'reinforcement': entry.get('reinforcement')
            }
    # Fallback по (length, width) без load_code
    if len(key) == 3:
        length, width, _ = key
        for candidate_key, candidate_entries in order_info_list.items():
            if len(candidate_key) >= 2 and candidate_key[0] == length and candidate_key[1] == width:
                for entry in candidate_entries:
                    if entry.get('qty_remaining', 0) > 0:
                        return {
                            'kp_id': entry.get('kp_id'),
                            'customer': entry.get('customer'),
                            'kp_date': entry.get('kp_date'),
                            'plate_name': entry.get('plate_name'),
                            'load_code': entry.get('load_code'),
                            'reinforcement': entry.get('reinforcement')
                        }
    return {}


def _optimize_2d_with_lengths(orders_2d: list, plate_width: int = 1200,
                               min_useful_width: int = 200,
                               opt_config: OptimizationConfig = None) -> dict:
    """
    ПРИВАТНАЯ функция: Полная 2D оптимизация с длинами в ILP модели.
    Минимизирует СТОИМОСТЬ (не просто количество плит!) с учётом ОБЕИХ размерностей (длина + ширина).
    
    УЛУЧШЕНИЯ (версия 2.0):
    ✅ Учёт реальных цен плит из базы данных (get_price)
    ✅ Учёт стоимости продольных и поперечных резов
    ✅ Фильтрация бесполезных вариантов (скорость ↑ в 2-3 раза)
    ✅ Настраиваемые штрафы и бонусы через OptimizationConfig
    ✅ Сохранение kp_id и customer для каждой плиты (для диаграммы Ганта)
    
    Args:
        orders_2d: [{'length': 5.6, 'width': 320, 'qty': 11, 'kp_id': 1, 'customer': 'Роман'}, ...] — спрос по (длина, ширина)
        plate_width: ширина исходной плиты в мм (1200)
        min_useful_width: минимальная полезная ширина остатка
        opt_config: конфигурация параметров оптимизации (штрафы, бонусы)
    
    Returns:
        {
            'primary_cuts': [{'width', 'rest', 'qty', 'lengths': [5.6, ...]}, ...],
            'secondary_cuts': [{'source', 'cuts', 'qty', 'pieces', 'lengths': [...], 'type'}, ...],
            'total_plates': int,
            'plate_assignments': [{'length', 'width', 'source', 'kp_id', 'customer', ...}, ...]
        }
    """
    # Используем дефолтную конфигурацию, если не передана
    if opt_config is None:
        opt_config = DEFAULT_CONFIG
    try:
        from pulp import LpProblem, LpMinimize, LpVariable, LpInteger, lpSum, value, PULP_CBC_CMD, LpStatus
    except ImportError:
        print('[OPT_2D] PuLP не установлен.')
        return {}
    
    if not orders_2d:
        return {}
    
    print(f"\n[OPT_2D] === ПОЛНАЯ 2D ОПТИМИЗАЦИЯ ===")
    print(f"[OPT_2D] Заказ:")
    for order in orders_2d:
        print(f"  {order['qty']}x {order['length']}м x {order['width']}мм")
    
    # 1. ПОДГОТОВКА: Группируем спрос по (length, width, load_code)
    # ИСПРАВЛЕНИЕ: Добавляем load_code в ключ для различения плит с одинаковыми размерами, но разной нагрузкой
    demand_2d = {}  # {(length, width, load_code): qty}
    for order in orders_2d:
        load_code = order.get('load_code', 800)
        key = (order['length'], order['width'], load_code)
        demand_2d[key] = demand_2d.get(key, 0) + order['qty']
    # #region agent log: demand keys 59/10 (H3,H4)
    _demand_59_10 = [(list(k), v) for k, v in demand_2d.items() if abs(k[0] - 5.99) < 0.02 and (k[2] == 10 or abs(float(k[2]) - 10) < 0.01)]
    if _demand_59_10:
        try:
            open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8").write(__import__("json").dumps({"hypothesisId": "H_59_10_demand", "location": "optimization.py:demand_2d_built", "message": "demand_2d: ключи 5.99м 10п (length, width, load_code)", "data": {"keys": _demand_59_10}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    # 1.5 НОВОЕ: Создаём маппинг (length, width, load_code) -> СПИСОК информации о КП
    # ИСПРАВЛЕНИЕ: Теперь ключ включает load_code для различения плит с разной нагрузкой
    order_info_list = {}  # {(length, width, load_code): [список записей для каждого КП]}
    for order in orders_2d:
        load_code = order.get('load_code', 800)
        key = (order['length'], order['width'], load_code)
        if key not in order_info_list:
            order_info_list[key] = []
        # Добавляем запись для КАЖДОГО заказа с qty_remaining
        order_info_list[key].append({
            'kp_id': order.get('kp_id'),
            'customer': order.get('customer', 'неизвестно'),
            'kp_date': order.get('kp_date', 'неизвестно'),
            'plate_name': order.get('plate_name', ''),
            'load_code': load_code,
            'reinforcement': order.get('reinforcement', 0),
            'qty_remaining': order.get('qty', 1)  # Сколько плит этого КП осталось назначить
        })

    slot_lists, slot_cursors = _build_proportional_slot_lists(orders_2d, demand_2d)

    tolerance_length = 0  # Строгое совпадение длины (после правок length_dm_to_m)
    tolerance_width = 20     # ±20мм по ширине (генерация опций, напр. поперечный рез остаток→цель)
    demand_tolerance_width = 10  # ±10мм при сопоставлении спроса с источником (по письму: допуск реза при работе)

    # 2. ГЕНЕРАЦИЯ ОПЦИЙ ПЕРВИЧНЫХ РЕЗОВ (с длинами!)
    primary_options = []
    option_id = 0
    
    # Таблица сужений (из таблицы допустимых резов)
    # Формат: (исходная_ширина_остатка, целевая_ширина, отход)
    # ВАЖНО: Значения остатков рассчитаны БЕЗ явного учёта пропила в коде
    # Пропил косвенно учтён через значения narrowing (разница между source_rest и target_w)
    NARROWING_TABLE = [
        (480, 460, 20),   # Остаток 480мм → 460мм (из реза 720+480)
        (500, 460, 40),   # Остаток 500мм → 460мм (из реза 700+500)
        (495, 460, 35),   # Остаток 495мм → 460мм (из реза 720+495 или 700+495)
        (740, 720, 20),   # Остаток 740мм → 720мм (из реза 460+740)
        (690, 660, 30),   # Остаток 690мм → 660мм (из реза 460+690)
        (890, 860, 30),   # Остаток 890мм → 860мм (из реза 320+890)
        (495, 480, 15),   # Остаток 495мм → 480мм
    ]
    
    # Создаём обратный индекс: для каждой целевой ширины -> список (основная_ширина, ширина_остатка, отход)
    target_to_sources = {}  # {460: [(720, 480, 20), (700, 500, 40), ...]}
    for source_rest, target_w, waste in NARROWING_TABLE:
        if target_w not in target_to_sources:
            target_to_sources[target_w] = []
        # Ищем, из какого первичного реза получается source_rest
        main_w = plate_width - source_rest
        if 200 <= main_w <= 1000:  # Разумный диапазон для основной части
            target_to_sources[target_w].append((main_w, source_rest, waste))
    
    print(f"[OPT_2D] Таблица narrowing создана: {len(NARROWING_TABLE)} правил")
    for target_w, sources in target_to_sources.items():
        print(f"  {target_w}мм можно получить через: {sources}")
    
    solid_widths = sorted(set([plate_width, 1080]))

    for (length, width, load_code), qty in demand_2d.items():
        # Получаем информацию о КП для этой плиты (без уменьшения счётчика)
        # ИСПРАВЛЕНИЕ: Ключ теперь включает load_code
        order_info = _peek_order_info(order_info_list, (length, width, load_code))
        
        # Вариант 1: Плита БЕЗ реза (ширины из списка solid_widths)
        # Эти ширины НЕ РЕЖУТСЯ и используются как есть
        if width in solid_widths:
            primary_options.append({
                'id': option_id,
                'length': length,
                'main': width,
                'rest': 0,
                'type': 'solid',  # Без резов
                'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                'kp_id': order_info.get('kp_id'),
                'customer': order_info.get('customer'),
                'kp_date': order_info.get('kp_date'),
                'plate_name': order_info.get('plate_name')
            })
            option_id += 1

        # Вариант 2: Плита С ПРЯМЫМ резом (ширина < исходной плиты)
        elif width < plate_width:
            # Пропил косвенно учтён в таблице NARROWING_TABLE
            rest = plate_width - width
            # Создаём вариант для ЛЮБОЙ ширины
            # Если rest < min_useful_width, остаток просто пойдёт в отход
            primary_options.append({
                'id': option_id,
                'length': length,
                'main': width,
                'rest': rest,
                'type': 'direct',  # Прямой рез
                'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                'kp_id': order_info.get('kp_id'),
                'customer': order_info.get('customer'),
                'kp_date': order_info.get('kp_date'),
                'plate_name': order_info.get('plate_name')
            })
            option_id += 1
            
            # НОВОЕ! Вариант 3: Плита через НЕПРЯМОЙ рез (с narrowing остатка)
            # Ищем, можно ли получить эту ширину через сужение остатка от ДРУГОГО реза
            if width in target_to_sources:
                for main_w, rest_w, waste in target_to_sources[width]:
                    # Создаём первичный рез main_w + rest_w
                    # Остаток rest_w потом автоматически сузится до width
                    if main_w != width and rest_w >= min_useful_width:  # Не дублируем прямой рез
                        primary_options.append({
                            'id': option_id,
                            'length': length,
                            'main': main_w,           # Например, 720мм (основная часть)
                            'rest': rest_w,           # Например, 480мм (остаток)
                            'type': 'indirect',       # Непрямой рез через narrowing
                            'target_width': width,    # Целевая ширина: 460мм (что нужно)
                            'narrowing_waste': waste, # Отход при сужении: 20мм
                            'load_code': order_info.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                            'kp_id': order_info.get('kp_id'),
                            'customer': order_info.get('customer'),
                            'kp_date': order_info.get('kp_date'),
                            'plate_name': order_info.get('plate_name')
                        })
                        option_id += 1
    
    print(f"[OPT_2D] Опций первичных резов (до фильтрации): {len(primary_options)}")
    # #region agent log (2d5c43) Plan B: опции для 5.1/320 и 6/530 до фильтра
    try:
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type'), "load_code": o.get('load_code')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with open(_log, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_gen", "location": "optimization:primary_options_after_build", "message": "options for 320 and 530 before filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530, "solid_widths": solid_widths}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 2.5 ФИЛЬТРАЦИЯ ПЕРВИЧНЫХ ОПЦИЙ (Улучшение 4: убираем заведомо невыгодные)
    filtered_primary = []
    for opt in primary_options:
        # Правило 1 УДАЛЕНО: теперь плиты с маленьким остатком (< min_useful_width) 
        # тоже создаются, а остаток просто идёт в отход
        
        # Правило 2: Пропускаем indirect, если есть direct с тем же результатом
        if opt.get('type') == 'indirect':
            target_w = opt.get('target_width')
            has_direct = any(
                o['type'] == 'direct' and 
                o['main'] == target_w and
                abs(o['length'] - opt['length']) <= tolerance_length
                for o in primary_options
            )
            if has_direct:
                continue
        
        filtered_primary.append(opt)
    
    primary_options = filtered_primary
    print(f"[OPT_2D] После фильтрации осталось: {len(primary_options)} первичных опций")
    # #region agent log (2d5c43) Plan B: опции для 320/530 после фильтра
    try:
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _opts_320 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 320 or o.get('target_width') == 320]
        _opts_530 = [{"id": o['id'], "length": o['length'], "main": o['main'], "type": o.get('type')} for o in primary_options if o.get('main') == 530 or o.get('target_width') == 530]
        with open(_log, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_opt_filter", "location": "optimization:primary_options_after_filter", "message": "options for 320 and 530 after filter", "data": {"opts_320": _opts_320, "opts_530": _opts_530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 3. ГЕНЕРАЦИЯ ОПЦИЙ ВТОРИЧНЫХ РЕЗОВ (2D: длина + ширина!)
    secondary_options = []
    
    # Собираем все возможные остатки (length, rest_width)
    possible_rests = {}  # {(length, rest_width): [source_option_ids]}
    for opt in primary_options:
        key = (opt['length'], opt['rest'])
        if key not in possible_rests:
            possible_rests[key] = []
        possible_rests[key].append(opt['id'])
    
    sec_id = 0
    for (source_length, source_width), source_ids in possible_rests.items():
        # Пропускаем остатки нулевой ширины (плиты 1200мм без резов)
        if source_width < min_useful_width:
            continue
        
        # Для каждого остатка проверяем все целевые (length, width, load_code)
        # ИСПРАВЛЕНИЕ: Теперь ключ включает load_code
        for (target_length, target_width, target_load_code), qty in demand_2d.items():
            
            # Вариант A: Множественная резка по ширине (одинаковая длина)
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            if abs(target_length - source_length) <= tolerance_length and target_width <= source_width:
                # РАНЬШЕ: брали только максимум кусков (pieces = source_width // target_width)
                # ТЕПЕРЬ: перебираем все варианты от 1 до max_pieces, чтобы можно было
                # получать 1, 2, ... плит из одного остатка (например, 0.88 → 1×0.32 с хвостом 0.56)
                max_pieces = source_width // target_width
                for pieces in range(1, max_pieces + 1):
                    waste = source_width - (pieces * target_width)
                    # Отбрасываем варианты с слишком большим отходом.
                    # Для случаев с ОДНОЙ плитой позволяем больше отхода (до 80%),
                    # чтобы не выкидывать схему 0.88 → 0.32 + 0.56.
                    max_waste_fraction = 0.8 if pieces == 1 else 0.5
                    if waste <= source_width * max_waste_fraction:
                        secondary_options.append({
                            'id': sec_id,
                            'source_length': source_length,
                            'source_rest': source_width,
                            'output_length': target_length,
                            'output_width': target_width,
                            'pieces': pieces,
                            'waste': waste,
                            'type': 'multiple',
                            'source_ids': source_ids,
                            'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                        })
                        sec_id += 1

            
            # Вариант A2: Комбинированная резка (множественная по ширине + поперечная по длине)
            # Это позволяет резать остаток 5.6м × 880мм → 2× 3.31м × 320мм
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            if target_length < source_length - 0.1 and target_width <= source_width:  # Целевая длина КОРОЧЕ остатка
                pieces = source_width // target_width
                if pieces >= 1:
                    # Проверяем, что целевая длина влезает хотя бы раз
                    waste_width = source_width - (pieces * target_width)
                    waste_length = (source_length - target_length) * 1000  # в мм
                    
                    if waste_width < source_width * 0.5:
                        secondary_options.append({
                            'id': sec_id,
                            'source_length': source_length,
                            'source_rest': source_width,
                            'output_length': target_length,  # КОРОЧЕ остатка!
                            'output_width': target_width,
                            'pieces': pieces,  # Кусков по ширине
                            'waste': waste_width,
                            'length_waste': waste_length,
                            'type': 'multiple_transverse',  # Комбинированный тип
                            'source_ids': source_ids,
                            'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                        })
                        sec_id += 1
            
            # Вариант B: Сужение (narrowing)
            if (abs(target_length - source_length) <= tolerance_length and
                target_width < source_width <= target_width + 100):
                waste = source_width - target_width
                if waste <= 100:
                    secondary_options.append({
                        'id': sec_id,
                        'source_length': source_length,
                        'source_rest': source_width,
                        'output_length': target_length,
                        'output_width': target_width,
                        'pieces': 1,
                        'waste': waste,
                        'type': 'narrowing',
                        'source_ids': source_ids,
                        'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                    })
                    sec_id += 1
            
            # Вариант C: Поперечный рез (transverse cut)
            # ВАЖНО: Нельзя получить больше, чем есть! target_width <= source_width
            # (раньше проверяли abs(target_width - source_width) <= tolerance_width, 
            #  что позволяло target_width > source_width на 20мм — это баг!)
            if (target_length < source_length - 0.1 and
                target_width <= source_width and
                source_width - target_width <= tolerance_width):
                length_waste = (source_length - target_length) * 1000  # в мм
                secondary_options.append({
                    'id': sec_id,
                    'source_length': source_length,
                    'source_rest': source_width,
                    'output_length': target_length,
                    'output_width': target_width,
                    'pieces': 1,
                    'waste': 0,
                    'length_waste': length_waste,
                    'type': 'transverse',
                    'source_ids': source_ids,
                    'target_order_key': (target_length, target_width, target_load_code)  # ИСПРАВЛЕНИЕ: Добавляем load_code
                })
                sec_id += 1
    
    print(f"[OPT_2D] Опций вторичных резов (до фильтрации): {len(secondary_options)}")
    
    # 3.5 ФИЛЬТРАЦИЯ ВТОРИЧНЫХ ОПЦИЙ (Улучшение 4: убираем дубликаты и невыгодные)
    filtered_secondary = []
    seen_combinations = set()
    
    for opt in secondary_options:
        # Правило 3: Убираем дубликаты (одинаковые варианты).
        # ВАЖНО: учитываем также количество кусков (pieces), чтобы варианты
        # 0.88 → 1×0.32 и 0.88 → 2×0.32 не считались одинаковыми.
        key = (
            opt['source_length'], 
            opt['source_rest'], 
            opt['output_length'], 
            opt['output_width'], 
            opt['type'],
            opt.get('pieces', 1)
        )
        
        if key in seen_combinations:
            continue
        seen_combinations.add(key)
        
        # Правило 4: Пропускаем варианты с огромными отходами.
        # Для случаев с ОДНОЙ плитой (pieces == 1) позволяем до 80% площади отхода,
        # иначе — как раньше, 30%.
        waste_width = opt.get('waste', 0)
        waste_length = opt.get('length_waste', 0)
        
        source_area = opt['source_length'] * opt['source_rest']
        waste_area = (waste_width * opt['source_length']) + (waste_length * opt['source_rest'] / 1000.0)
        
        max_waste_fraction_area = 0.8 if opt.get('pieces', 1) == 1 else 0.3
        if opt['type'] != 'multiple_transverse' and waste_area > source_area * max_waste_fraction_area:
            continue
        
        # Правило 5: Пропускаем transverse с отходами > 50% длины
        if opt['type'] == 'transverse':
            waste_fraction = waste_length / (opt['source_length'] * 1000) if opt['source_length'] > 0 else 0
            if waste_fraction > 0.5:
                continue
        
        filtered_secondary.append(opt)
    
    secondary_options = filtered_secondary
    print(f"[OPT_2D] После фильтрации осталось: {len(secondary_options)} вторичных опций")
    
    # 4. СОЗДАНИЕ ILP МОДЕЛИ
    prob = LpProblem("2D_Optimization", LpMinimize)
    
    # Переменные
    x_prim = {opt['id']: LpVariable(f"prim_{opt['id']}", lowBound=0, cat=LpInteger) 
              for opt in primary_options}
    x_sec = {opt['id']: LpVariable(f"sec_{opt['id']}", lowBound=0, cat=LpInteger) 
             for opt in secondary_options}
    
    # 5. ОГРАНИЧЕНИЯ: Покрытие спроса по (length, width, load_code)
    # ИСПРАВЛЕНИЕ: Теперь ключ включает load_code для различения плит с разной нагрузкой
    BIG_M = 1e6
    surplus_vars = []
    # #region agent log: total demand (H_demand)
    _debug_log = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    _total_demand = sum(demand_2d.values())
    try:
        open(_debug_log, 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H_demand", "location": "optimization.py:before_demand_loop", "message": "total demand and keys", "data": {"total_demand": _total_demand, "demand_keys_count": len(demand_2d), "demand_keys": [(round(k[0], 2), k[1], k[2], v) for k, v in demand_2d.items()]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) спрос по плитам РЕСКЬЮ: 50.8-3.2-8П, 59.8-5.3-8П
    try:
        _d508320 = sum(q for k, q in demand_2d.items() if abs(k[0] - 5.08) < 0.02 and k[1] == 320 and (k[2] == 8 or k[2] == '8'))
        _d598530 = sum(q for k, q in demand_2d.items() if abs(k[0] - 5.98) < 0.02 and k[1] == 530 and (k[2] == 8 or k[2] == '8'))
        _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
        with open(_log_p, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_demand_rescue", "location": "optimization:demand_2d", "message": "demand for 5.08/320 and 5.98/530 before solver", "data": {"demand_508_320": _d508320, "demand_598_530": _d598530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    no_sources_keys = []  # [(key, qty), ...] for H1
    opt_to_demands = {}  # opt_id -> [(demand_key, qty), ...] — какой спрос использует этот opt
    # Группируем спросы по множеству источников: один constraint на группу (lpSum >= sum(qty)),
    # чтобы 4.78+4.79 производили 25 плит, а не max(17,8)=17.
    sources_to_demands = {}  # (frozenset(prim_ids), frozenset(sec_id_pieces)) -> (sources_list, [(dk, qty), ...])
    for (target_length, target_width, target_load_code), qty in demand_2d.items():
        sources = []
        prim_ids = set()
        sec_id_pieces = set()  # (opt_id, pieces)
        dk = (round(target_length, 2), target_width, target_load_code)
        
        # Источник 1a: Первичные резы ПРЯМЫЕ (type='direct' или 'solid')
        for opt in primary_options:
            if (abs(opt['length'] - target_length) <= tolerance_length and 
                abs(opt['main'] - target_width) <= demand_tolerance_width and
                opt.get('type') in ['direct', 'solid'] and
                opt.get('load_code', 800) == target_load_code):
                sources.append(x_prim[opt['id']])
                prim_ids.add(opt['id'])
                opt_to_demands.setdefault(opt['id'], []).append((dk, qty))
        
        # Источник 1b: Первичные резы НЕПРЯМЫЕ (type='indirect', через narrowing)
        for opt in primary_options:
            if (abs(opt['length'] - target_length) <= tolerance_length and
                opt.get('type') == 'indirect' and
                abs(opt.get('target_width', 0) - target_width) <= demand_tolerance_width and
                opt.get('load_code', 800) == target_load_code):
                sources.append(x_prim[opt['id']])
                prim_ids.add(opt['id'])
                opt_to_demands.setdefault(opt['id'], []).append((dk, qty))
        
        # Источник 2: Вторичные резы
        for opt in secondary_options:
            opt_target_key = opt.get('target_order_key', (0, 0, 800))
            if len(opt_target_key) == 3:
                opt_target_load = opt_target_key[2]
            else:
                opt_target_load = 800
            if (abs(opt['output_length'] - target_length) <= tolerance_length and 
                abs(opt['output_width'] - target_width) <= demand_tolerance_width and
                opt_target_load == target_load_code):
                sources.append(x_sec[opt['id']] * opt['pieces'])
                sec_id_pieces.add((opt['id'], opt['pieces']))
        
        if sources:
            key = (frozenset(prim_ids), frozenset(sec_id_pieces))
            if key not in sources_to_demands:
                sources_to_demands[key] = (list(sources), [])
            sources_to_demands[key][1].append((dk, qty))
        else:
            # === КРИТИЧЕСКОЕ ПРЕДУПРЕЖДЕНИЕ: НЕТ ИСТОЧНИКОВ ДЛЯ СПРОСА ===
            no_sources_keys.append(((round(target_length, 2), target_width, target_load_code), qty))
            # #region agent log: no sources (H1) + same L/W other load (H3)
            _same_lw_other_load = [opt.get('load_code') for opt in primary_options if abs(opt['length'] - target_length) <= tolerance_length and abs(opt.get('main', opt.get('target_width', 0)) - target_width) <= demand_tolerance_width and opt.get('load_code', 800) != target_load_code]
            try:
                open(_debug_log, 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H1", "location": "optimization.py:no_sources", "message": "demand key has no sources", "data": {"target_length": target_length, "target_width": target_width, "target_load_code": target_load_code, "qty": qty, "same_LW_other_load_codes": _same_lw_other_load[:5]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
            import logging
            logger = logging.getLogger(__name__)
            logger.error(f"[OPT_2D] ❌ НЕТ ИСТОЧНИКОВ для плиты: {target_length}м x {target_width}мм (load={target_load_code}) x{qty}шт!")
            logger.error(f"[OPT_2D]    Эта плита НЕ БУДЕТ включена в план!")
            logger.error(f"[OPT_2D]    Возможные причины:")
            logger.error(f"[OPT_2D]      - Нестандартная ширина {target_width}мм не поддерживается NARROWING_TABLE")
            logger.error(f"[OPT_2D]      - Нет подходящих исходных плит для реза")
            logger.error(f"[OPT_2D]      - Несовместимый load_code={target_load_code}")
            print(f"[OPT_2D] ⚠️ ВНИМАНИЕ: Плита {target_length}м x {target_width}мм НЕ МОЖЕТ быть произведена!")
    
    # #region agent log (2d5c43) Plan B: ключи без источников (no_sources_keys)
    try:
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _no_src_ser = [[list(k), q] for k, q in no_sources_keys]
        with open(_log, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_no_sources", "location": "optimization:after_demand_loop", "message": "demand keys with no sources in model", "data": {"no_sources_keys": _no_src_ser, "count": len(no_sources_keys)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # Сохраняем ссылку на группу с opt 105 для проверки после solve (95694e)
    _dbg_grp105_sources = None
    _dbg_grp105_total_qty = None
    _dbg_rescue_groups = {}  # (length_key, width) -> (sources_list, total_qty) для 5.08/320 и 5.98/530
    _opt_598_for_dbg = next((o for o in primary_options if abs(o.get('length', 0) - 5.98) < 0.02 and o.get('main') == 665), None)
    # Ограничения по группам: явно требуем lpSum(sources) >= total_qty (исправление потери плиты при ==)
    for grp_idx, (key, (sources_list, demands_list)) in enumerate(sources_to_demands.items()):
        total_qty = sum(q for _, q in demands_list)
        surplus = LpVariable(f"surplus_grp_{grp_idx}", lowBound=0, cat=LpInteger)
        surplus_vars.append(surplus)
        prob += lpSum(sources_list) >= total_qty, f"demand_grp_{grp_idx}_min"
        prob += surplus == lpSum(sources_list) - total_qty, f"demand_grp_{grp_idx}_surplus"
        if _opt_598_for_dbg and key[0] and _opt_598_for_dbg['id'] in key[0]:
            _dbg_grp105_sources = list(sources_list)
            _dbg_grp105_total_qty = total_qty
        for _dk, _q in demands_list:
            _len_w = (round(float(_dk[0]), 2), _dk[1]) if _dk else (0, 0)
            if abs(_len_w[0] - 5.08) < 0.02 and _len_w[1] == 320:
                _dbg_rescue_groups[(5.08, 320)] = (list(sources_list), total_qty)
            if abs(_len_w[0] - 5.98) < 0.02 and _len_w[1] == 530:
                _dbg_rescue_groups[(5.98, 530)] = (list(sources_list), total_qty)
    
    # Явный нижний предел для 5.98/665: гарантирует минимум по спросу (обход возможной потери в групповом ограничении)
    _demand_598665 = sum(q for k, q in demand_2d.items() if abs(k[0] - 5.98) < 0.02 and k[1] == 665)
    if _demand_598665 > 0:
        _opts_598 = [o for o in primary_options if abs(o.get('length', 0) - 5.98) < 0.02 and o.get('main') == 665]
        if _opts_598:
            _vars_598 = [x_prim[o['id']] for o in _opts_598]
            prob += lpSum(_vars_598) >= _demand_598665, "demand_598665_min"
    
    # #region agent log (95694e) группа ограничений, содержащая opt 5.98/665
    try:
        _opt_598 = next((o for o in primary_options if abs(o.get('length', 0) - 5.98) < 0.02 and o.get('main') == 665), None)
        if _opt_598:
            _oid = _opt_598['id']
            for _key, (_sources_list, _demands_list) in sources_to_demands.items():
                _prim_ids = _key[0] if isinstance(_key[0], (set, frozenset)) else set(_key[0]) if _key else set()
                if _oid in _prim_ids:
                    _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
                    with open(_log_p, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_grp_598665", "location": "optimization:demand_group", "message": "constraint group for 5.98/665 opt", "data": {"opt_id": _oid, "total_qty": sum(q for _, q in _demands_list), "demands_list": [[list(d[0]), d[1]] for d in _demands_list], "prim_ids": list(_prim_ids)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
                    break
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) группа ограничений для 5.08/320 и 5.98/530
    try:
        for _lk, _lw in [(5.08, 320), (5.98, 530)]:
            _opt_r = next((o for o in primary_options if abs(o.get('length', 0) - _lk) < 0.02 and o.get('main') == _lw), None)
            if not _opt_r:
                continue
            _oid = _opt_r['id']
            for _key, (_sources_list, _demands_list) in sources_to_demands.items():
                _prim_ids = _key[0] if isinstance(_key[0], (set, frozenset)) else set(_key[0]) if _key else set()
                if _oid in _prim_ids:
                    _label = "5.08/320" if _lk == 5.08 else "5.98/530"
                    _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
                    with open(_log_p, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_grp_rescue", "location": "optimization:demand_group", "message": f"constraint group for {_label} opt", "data": {"plate_key": _label, "opt_id": _oid, "total_qty": sum(q for _, q in _demands_list), "prim_ids": list(_prim_ids)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
                    break
    except Exception:
        pass
    # #endregion
    # #region agent log: no_sources summary (H1)
    if no_sources_keys:
        _plates_lost = sum(q for _, q in no_sources_keys)
        try:
            open(_debug_log, 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H1", "location": "optimization.py:no_sources_summary", "message": "no_sources summary", "data": {"keys_count": len(no_sources_keys), "total_plates_lost": _plates_lost, "keys": [(list(k), q) for k, q in no_sources_keys]}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #region agent log (session 73b708) H_A: потери на этапе оптимизатора
        try:
            _agent_log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-73b708.log"
            open(_agent_log, 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_A", "location": "optimization.py:no_sources_summary", "message": "Plates lost at optimizer (no sources)", "data": {"keys_count": len(no_sources_keys), "total_plates_lost": _plates_lost, "keys": [(list(k), q) for k, q in no_sources_keys]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
    # #endregion
    # #region agent log: competing demand keys (share same primary opts)
    _competing_opts = []
    for opt_id, demands in opt_to_demands.items():
        lengths = set(d[0][0] for d in demands)
        if len(lengths) > 1:
            opt_info = next((o for o in primary_options if o['id'] == opt_id), None)
            _competing_opts.append({"opt_id": opt_id, "opt_length": opt_info['length'] if opt_info else None,
                                    "demands": [(list(d[0]), d[1]) for d in demands]})
    if _competing_opts:
        try:
            open(_debug_log, 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H_opt_competing", "location": "optimization.py:competing_demands", "message": "primary opt used by multiple demand keys (different lengths)", "data": {"competing_opts": _competing_opts}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
        except Exception:
            pass
    # #endregion
    # 5.5 ПРИОРИТЕТНОЕ ОГРАНИЧЕНИЕ: Solid плиты ОБЯЗАТЕЛЬНЫ для полных ширин (1200/1080мм)!
    # ИСПРАВЛЕНИЕ: Теперь ключ включает load_code
    print(f"[OPT_2D] Проверяем приоритетные ограничения для solid-плит: {solid_widths}")
    for (target_length, target_width, target_load_code), qty in demand_2d.items():
        if target_width in solid_widths:
            solid_sources = []
            for opt in primary_options:
                if (abs(opt['length'] - target_length) <= tolerance_length and 
                    opt['main'] == target_width and
                    opt.get('type') == 'solid' and
                    opt['main'] == target_width and
                    opt.get('load_code', 800) == target_load_code):  # ИСПРАВЛЕНИЕ: проверяем load_code
                    solid_sources.append(x_prim[opt['id']])
            
            if solid_sources:
                # ЖЁСТКОЕ требование: удовлетворить спрос ТОЛЬКО solid-плитами!
                prob += lpSum(solid_sources) >= qty, f"solid_priority_{target_length}m_{target_width}mm_{target_load_code}"
                print(f"[OPT_2D] ✓ ПРИОРИТЕТ: {qty} плит {target_width}мм × {target_length}м (load={target_load_code}) ОБЯЗАТЕЛЬНО solid")
    
    # 6. ОГРАНИЧЕНИЯ: Баланс остатков
    for (source_length, source_width), source_ids in possible_rests.items():
        # Пропускаем остатки нулевой ширины (solid плиты не создают остатков!)
        if source_width == 0:
            continue
            
        # Произведено (ТОЛЬКО из плит С резом, не solid!)
        produced = []
        for opt_id in source_ids:
            opt = next((o for o in primary_options if o['id'] == opt_id), None)
            if opt and opt.get('type') != 'solid':  # ИСКЛЮЧАЕМ SOLID!
                produced.append(x_prim[opt_id])
        
        # Использовано
        consumed = []
        for opt in secondary_options:
            if (abs(opt['source_length'] - source_length) <= tolerance_length and 
                opt['source_rest'] == source_width):
                consumed.append(x_sec[opt['id']])
        
        if produced and consumed:
            prob += lpSum(consumed) <= lpSum(produced), f"balance_{source_length}m_{source_width}mm"
    
    # 7. ЦЕЛЕВАЯ ФУНКЦИЯ (УЛУЧШЕННАЯ: учёт реальных цен + приоритет остатков)
    total_cost = 0
    for surplus in surplus_vars:
        total_cost += BIG_M * surplus

    # Новое: мягкий приоритет на МЕНЬШЕЕ количество исходных плит 1200 мм.
    # total_plates_var = сумма всех первичных резов (каждый такой рез — одна плита 1.2 м).
    total_plates_var = lpSum(x_prim.values())
    
    # 7.1 Стоимость ПЕРВИЧНЫХ РЕЗОВ (плиты + продольные резы)
    print(f"[OPT_2D] Расчёт стоимости первичных резов...")
    for opt in primary_options:
        qty_var = x_prim[opt['id']]
        
        # Получаем РЕАЛЬНУЮ цену плиты из базы данных
        plate_price = get_price(opt['length'], 8, cfg.PRICE_DB_PATH)
        if plate_price is None:
            plate_price = 10000  # Дефолтная цена, если нет в БД
        
        # Стоимость продольного реза (если плита режется)
        if opt['type'] in ['direct', 'indirect']:
            cut_cost = cfg.LONG_CUT_PRICE_PER_M * opt['length']
        else:
            cut_cost = 0  # Solid плита - без реза
        
        # КРИТИЧНО: Плиты solid (1200мм целиком) должны иметь ВЫСОКИЙ приоритет
        # Добавляем огромный штраф, если solid-плита НЕ используется, но есть спрос
        if opt['type'] == 'solid':
            # Бонус за использование solid-плиты (делаем их очень выгодными)
            total_cost += qty_var * (plate_price + cut_cost - 5000)  # БОНУС 5000 руб!
        else:
            # Обычная стоимость для плит с резом
            total_cost += qty_var * (plate_price + cut_cost)
    
    # 7.2 Стоимость ВТОРИЧНЫХ РЕЗОВ (продольные + поперечные)
    for opt in secondary_options:
        qty_var = x_sec[opt['id']]
        
        # Продольный рез (для narrowing и multiple)
        if opt['type'] in ['narrowing', 'multiple', 'multiple_transverse']:
            cut_cost = cfg.LONG_CUT_PRICE_PER_M * opt['source_length']
            total_cost += qty_var * cut_cost
        
        # ПОПЕРЕЧНЫЙ РЕЗ (используем цену за штуку × количество резов)
        if opt['type'] in ['transverse', 'multiple_transverse']:
            # TRANSVERSE_CUT_PRICE = 1200 руб/шт (из конфига)
            total_cost += qty_var * cfg.TRANSVERSE_CUT_PRICE
    
    # 7.3 ШТРАФ ЗА НЕИСПОЛЬЗОВАННЫЕ ОСТАТКИ (настраиваемый через конфиг)
    unused_penalty = 0
    for (source_length, source_width), source_ids in possible_rests.items():
        produced = [x_prim[opt_id] for opt_id in source_ids]
        consumed = []
        for opt in secondary_options:
            if (abs(opt['source_length'] - source_length) <= tolerance_length and 
                opt['source_rest'] == source_width):
                consumed.append(x_sec[opt['id']])
        
        if produced and consumed:
            unused = lpSum(produced) - lpSum(consumed)
            
            # НОВАЯ ФОРМУЛА: Стоимость остатка = цена плиты × (ширина_остатка / ширина исходной плиты)
            base_price = get_price(source_length, 6, cfg.PRICE_DB_PATH)
            if base_price is None:
                base_price = 5000  # Дефолтная цена для остатков
            rest_price = base_price * (source_width / float(plate_width))
            
            # Штраф = N% стоимости остатка (коэффициент берём из конфига)
            # OLD: 0.5 (50% стоимости), NEW: 0.15 (15% стоимости)
            unused_penalty += unused * rest_price * opt_config.unused_rest_penalty_coeff
    
    # 7.4 ШТРАФ ЗА ОТХОДЫ (в рублях, усиленный)
    waste_penalty = 0
    for opt in secondary_options:
        # Отходы по ШИРИНЕ (в мм → в м² → в рубли)
        waste_width_mm = opt.get('waste', 0)
        if waste_width_mm > 0:
            waste_area_m2 = (waste_width_mm / 1000.0) * opt['source_length']
            waste_price = waste_area_m2 * 1000  # ~1000 руб/м² за отход
            waste_penalty += x_sec[opt['id']] * waste_price
        
        # Отходы по ДЛИНЕ (в мм → в м² → в рубли)
        waste_length_mm = opt.get('length_waste', 0)
        if waste_length_mm > 0:
            waste_area_m2 = (waste_length_mm / 1000.0) * (opt['source_rest'] / 1000.0)
            waste_price = waste_area_m2 * 1000
            waste_penalty += x_sec[opt['id']] * waste_price
    
    # 7.5 БОНУС ЗА ИСПОЛЬЗОВАНИЕ ОСТАТКОВ (настраиваемый через конфиг)
    reuse_bonus = 0
    for opt in secondary_options:
        # За каждый вторичный рез - бонус (отрицательное значение уменьшает стоимость)
        # OLD: -500 руб за каждый вторичный рез
        # NEW: 0 (нет экономического бонуса)
        reuse_bonus += x_sec[opt['id']] * opt_config.secondary_reuse_bonus
    
    # 7.6 ИТОГОВАЯ ЦЕЛЕВАЯ ФУНКЦИЯ
    print(f"[OPT_2D] Минимизируем: стоимость плит + резов + штрафы + бонусы")
    print(f"[OPT_2D] Конфиг: unused_penalty={opt_config.unused_rest_penalty_coeff}, reuse_bonus={opt_config.secondary_reuse_bonus}")

    # Дополнительный приоритет: меньше исходных плит лучше.
    # Коэффициент 5000 ~ половина средней цены плиты: не ломает расчёт стоимости,
    # но заставляет оптимизатор избегать лишних плит, если есть альтернатива.
    plates_priority_penalty = total_plates_var * 5000.0

    prob += total_cost + unused_penalty + waste_penalty + reuse_bonus + plates_priority_penalty
    
    # 8. РЕШЕНИЕ
    print(f"[OPT_2D] Запуск решателя...")
    prob.solve(PULP_CBC_CMD(msg=0))
    
    # #region agent log (95694e) статус решателя
    try:
        _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
        with open(_log_p, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_solver_status", "location": "optimization:after_solve", "message": "solver status", "data": {"status": LpStatus[prob.status]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) проверка: фактическая сумма переменных группы 105 vs total_qty
    try:
        if _dbg_grp105_sources is not None and _dbg_grp105_total_qty is not None:
            _actual_sum = sum(value(v) for v in _dbg_grp105_sources)
            _satisfied = _actual_sum >= _dbg_grp105_total_qty
            _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_verify", "location": "optimization:after_solve", "message": "group 105 LHS vs RHS", "data": {"actual_sum": _actual_sum, "total_qty": _dbg_grp105_total_qty, "constraint_satisfied": _satisfied}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
            if not _satisfied:
                _lp_path = _log_p.parent / "debug-95694e-violation.lp"
                try:
                    prob.writeLP(str(_lp_path))
                    with open(_log_p, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_lp_export", "location": "optimization:after_solve", "message": "model written (constraint violated)", "data": {"lp_path": str(_lp_path)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) verify для групп 5.08/320 и 5.98/530
    try:
        _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
        for (_lk, _lw), (_src_list, _tq) in (_dbg_rescue_groups or {}).items():
            _actual = sum(value(v) for v in _src_list)
            _sat = _actual >= _tq
            _label = "5.08/320" if _lk == 5.08 else "5.98/530"
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_verify_rescue", "location": "optimization:after_solve", "message": f"group {_label} LHS vs RHS", "data": {"plate_key": _label, "actual_sum": _actual, "total_qty": _tq, "constraint_satisfied": _sat}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) ограничения, содержащие переменную opt 105
    try:
        if _opt_598_for_dbg:
            _var105 = x_prim[_opt_598_for_dbg['id']]
            _constr_names = []
            for _cname, _c in getattr(prob, 'constraints', {}).items():
                if getattr(_c, 'constants', None) and _var105 in _c.constants:
                    _constr_names.append(_cname)
            _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_constraints", "location": "optimization:after_solve", "message": "constraints involving x_prim[105]", "data": {"constraint_names": _constr_names}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    if LpStatus[prob.status] != 'Optimal':
        print(f"[OPT_2D] ⚠️ Решение не найдено! Статус: {LpStatus[prob.status]}")
        return {}
    
    # #region agent log (95694e) значения x_prim для опций 5.98/665 после решателя
    try:
        _opts_598665 = []
        for opt in primary_options:
            if abs(opt.get('length', 0) - 5.98) < 0.02 and opt.get('main') == 665:
                _val = value(x_prim[opt['id']])
                _opts_598665.append({"opt_id": opt['id'], "value": _val, "qty_rounded": int(round(_val))})
        if _opts_598665:
            _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_xprim_598665", "location": "optimization:after_solve", "message": "x_prim for 5.98/665 options", "data": {"opts": _opts_598665, "total_rounded": sum(o["qty_rounded"] for o in _opts_598665)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log (95694e) значения x_prim для опций 5.08/320 и 5.98/530 после решателя
    try:
        _rescue_opts = []
        for _lk, _lw in [(5.08, 320), (5.98, 530)]:
            _opts_r = [o for o in primary_options if abs(o.get('length', 0) - _lk) < 0.02 and o.get('main') == _lw]
            for o in _opts_r:
                _val = value(x_prim[o['id']])
                _rescue_opts.append({"plate_key": "5.08/320" if _lk == 5.08 else "5.98/530", "opt_id": o['id'], "value": _val, "qty_rounded": int(round(_val))})
        if _rescue_opts:
            _log_p = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-95694e.log"
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_xprim_rescue", "location": "optimization:after_solve", "message": "x_prim for 5.08/320 and 5.98/530 options", "data": {"opts": _rescue_opts}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # #region agent log (2d5c43) Plan B: значения x_prim для опций 320 и 530 после решателя
    try:
        _opts_320 = []
        _opts_530 = []
        for opt in primary_options:
            if opt.get('main') == 320 or opt.get('target_width') == 320:
                _v = value(x_prim[opt['id']])
                _opts_320.append({"opt_id": opt['id'], "length": opt['length'], "main": opt['main'], "type": opt.get('type'), "value": _v, "qty_rounded": int(round(_v))})
            if opt.get('main') == 530 or opt.get('target_width') == 530:
                _v = value(x_prim[opt['id']])
                _opts_530.append({"opt_id": opt['id'], "length": opt['length'], "main": opt['main'], "type": opt.get('type'), "value": _v, "qty_rounded": int(round(_v))})
        _log = __import__('pathlib').Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        with open(_log, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H_solver_val", "location": "optimization:after_solve", "message": "x_prim values for 320 and 530 options", "data": {"opts_320": _opts_320, "opts_530": _opts_530, "sum_320": sum(o["qty_rounded"] for o in _opts_320), "sum_530": sum(o["qty_rounded"] for o in _opts_530)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    # 9. ИЗВЛЕЧЕНИЕ РЕЗУЛЬТАТОВ
    result = {
        'primary_cuts': [],
        'secondary_cuts': [],
        'total_plates': 0,
        'plate_assignments': [],
        # Сохраняем исходный заказ, чтобы визуализация и отчёты точно знали, что просил пользователь
        'orders_requested': [order.copy() for order in orders_2d],
        # Метаданные для отслеживания остатков
        'rests_created': [],  # Остатки, созданные при первичных резах
        'rests_used': []      # Остатки, использованные во вторичных резах
    }
    
    # Первичные резы
    # ИСПРАВЛЕНИЕ: Для каждой плиты получаем свой kp_id из order_info_list!
    # Раньше все плиты с одинаковыми (length, width) получали kp_id первого КП.
    # Теперь каждая плита создаётся отдельно со своим kp_id.
    # #region agent log: summary empty plate_info (H2)
    _empty_primary_keys = []
    # #endregion
    for opt in primary_options:
        qty = int(round(value(x_prim[opt['id']])))
        if qty > 0:
            # Создаём ОТДЕЛЬНУЮ запись для КАЖДОЙ плиты
            for _ in range(qty):
                # Получаем информацию о следующем КП с уменьшением счётчика
                # ИСПРАВЛЕНИЕ: Для indirect типа используем target_width (заказанная ширина),
                # для direct/solid — main (они совпадают с заказом)
                lookup_width = opt.get('target_width', opt['main']) if opt.get('type') == 'indirect' else opt['main']
                # ИСПРАВЛЕНИЕ: Добавляем load_code в ключ поиска
                lookup_load_code = opt.get('load_code', 800)
                plate_info = _next_slot_info(slot_lists, slot_cursors, (opt['length'], lookup_width, lookup_load_code))
                # #region agent log: primary plate kp_id (H1, H2, H4)
                if not plate_info and opt.get('kp_id'):
                    _k = (opt['length'], lookup_width, lookup_load_code)
                    _empty_primary_keys.append(_k)
                    try:
                        with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f:
                            _f.write(__import__('json').dumps({"hypothesisId": "H2", "location": "optimization.py:primary_emit", "message": "primary plate_info empty", "data": {"key": list(_k), "opt_kp_id": opt.get('kp_id'), "opt_plate_name": (opt.get('plate_name') or '')[:50]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                    except Exception:
                        pass
                # #endregion
                result['primary_cuts'].append({
                    'width': opt['main'],
                    'demand_width': lookup_width,
                    'rest': opt['rest'],
                    'qty': 1,  # Каждая плита отдельно!
                    'lengths': [opt['length']],
                    'load_code': plate_info.get('load_code', 800) if plate_info else opt.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                    'kp_id': plate_info.get('kp_id') if plate_info else opt.get('kp_id'),
                    'customer': plate_info.get('customer') if plate_info else opt.get('customer'),
                    'kp_date': plate_info.get('kp_date') if plate_info else opt.get('kp_date'),
                    'plate_name': plate_info.get('plate_name') if plate_info else opt.get('plate_name')
                })
                result['total_plates'] += 1
                
                result['plate_assignments'].append({
                    'length': opt['length'],
                    'width': opt['main'],
                    'source': 'primary',
                    'rest_width': opt['rest'],
                    'load_code': plate_info.get('load_code', 800) if plate_info else opt.get('load_code', 800),  # ИСПРАВЛЕНИЕ: добавляем load_code
                    'kp_id': plate_info.get('kp_id') if plate_info else opt.get('kp_id'),
                    'customer': plate_info.get('customer') if plate_info else opt.get('customer'),
                    'kp_date': plate_info.get('kp_date') if plate_info else opt.get('kp_date'),
                    'plate_name': plate_info.get('plate_name') if plate_info else opt.get('plate_name')
                })
    
    # #region agent log: summary empty plate_info (H2)
    if _empty_primary_keys:
        try:
            from collections import Counter
            _c = Counter(_empty_primary_keys)
            _summary = [{"key": list(k), "count": _c[k]} for k in sorted(_c.keys())]
            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"hypothesisId": "H2", "location": "optimization.py:primary_emit_summary", "message": "plan keys with empty plate_info (summary)", "data": {"by_key": _summary, "total_plates_empty": len(_empty_primary_keys), "unique_keys": len(_c)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        except Exception:
            pass
    # #endregion
    
    # ========== НОВАЯ ЛОГИКА: СОРТИРОВКА ДЛЯ ПРОИЗВОДСТВА ==========
    # Требования завода:
    # 1. Первая плита ДОЛЖНА быть целой (без реза, rest=0)
    # 2. Плиты с одинаковым резом должны идти подряд
    print("[OPT_2D] 🔧 Применяем правила завода для порядка плит...")
    
    # Разделяем primary_cuts на группы
    solid_plates = []    # Целые плиты (rest=0)
    cut_plates = []      # Плиты с резом (rest>0)
    
    for cut in result['primary_cuts']:
        if cut['rest'] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)
    
    # ВАЖНО: Сортируем ЦЕЛЫЕ плиты тоже! (чтобы они не перемешивались с резанными)
    # Сортируем по ширине (целые 1200мм все вместе), потом по первой длине
    solid_plates.sort(key=lambda x: (-x['width'], -x['lengths'][0] if x.get('lengths') else 0))
    
    # Сортируем плиты с резом по (rest, width) — одинаковые резы идут подряд
    # Сортируем по убыванию ширины остатка, чтобы крупные резы шли в начале
    cut_plates.sort(key=lambda x: (-x['rest'], -x['width']))
    
    # Новый порядок: СНАЧАЛА целые (ОТСОРТИРОВАННЫЕ!), ПОТОМ плиты с резом (сгруппированные)
    result['primary_cuts'] = solid_plates + cut_plates
    
    # Универсальная пост-коррекция: для каждого ключа (L, W, lc) из demand_2d
    # добираем плиты, которые LP-решатель недодал.
    # Снимок делается ДО цикла, чтобы только что добавленные плиты не влияли на подсчёт have.
    _existing_cuts_snapshot = list(result['primary_cuts'])
    for (L, W, lc), need in demand_2d.items():
        have = sum(
            1 for c in _existing_cuts_snapshot
            if abs((c.get('lengths') or [0])[0] - L) < 0.02
            and c.get('demand_width', c.get('width')) == W
            and (c.get('load_code') == lc or str(c.get('load_code')) == str(lc))
        )
        if have >= need:
            continue
        rest = (plate_width - W) if W < plate_width else 0
        for _ in range(need - have):
            pi = _next_slot_info(slot_lists, slot_cursors, (L, W, lc))
            result['primary_cuts'].append({
                'width': W,
                'demand_width': W,
                'rest': rest,
                'qty': 1,
                'lengths': [L],
                'load_code': pi.get('load_code', lc) if pi else lc,
                'kp_id': pi.get('kp_id') if pi else None,
                'customer': pi.get('customer') if pi else None,
                'kp_date': pi.get('kp_date') if pi else None,
                'plate_name': pi.get('plate_name') if pi else None,
            })
            result['total_plates'] += 1
    
    # #region agent log (2d5c43) H1,H2,H5: demand vs primary_cuts, 6m 530/1200
    try:
        from pathlib import Path as _Path
        _log_2d5c43 = _Path(__file__).resolve().parent.parent / "debug-2d5c43.log"
        _demand_total = sum(demand_2d.values())
        _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
        _demand_by_key = {}
        for k, q in demand_2d.items():
            for tk in _target_keys:
                if abs(round(k[0], 2) - round(tk[0], 2)) <= 0.02 and k[1] == tk[1] and (k[2] == tk[2] or k[2] == '8'):
                    _demand_by_key[tuple(tk)] = _demand_by_key.get(tuple(tk), 0) + q
                    break
        _prim_total = len(result['primary_cuts'])
        _prim_by_key = {tk: 0 for tk in _target_keys}
        _prim_6_530_1200 = []
        for c in result['primary_cuts']:
            L = round((c.get('lengths') or [0])[0], 2)
            W = c.get('width', 0)
            lc = c.get('load_code', 8)
            for tk in _target_keys:
                if abs(L - tk[0]) <= 0.02 and W == tk[1] and (lc == tk[2] or lc == '8'):
                    _prim_by_key[tk] = _prim_by_key.get(tk, 0) + 1
                    break
            if 5.98 <= L <= 6.02 and W in (530, 1200) and len(_prim_6_530_1200) < 25:
                _prim_6_530_1200.append({"length": L, "width": W, "rest": c.get('rest', 0), "plate_name": (c.get('plate_name') or "")[:60]})
        _demand_by_key_ser = [list(k) + [v] for k, v in _demand_by_key.items()]
        _prim_by_key_ser = [list(k) + [v] for k, v in _prim_by_key.items()]
        with open(_log_2d5c43, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H1_H2_H5", "location": "optimization:_optimize_2d:primary_cuts_done", "message": "demand vs primary_cuts", "data": {"demand_total": _demand_total, "demand_by_key": _demand_by_key_ser, "primary_total": _prim_total, "primary_by_key": _prim_by_key_ser, "primary_6m_530_1200_sample": _prim_6_530_1200}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    print(f"[OPT_2D] ✓ Целых плит в начале: {len(solid_plates)}")
    print(f"[OPT_2D] ✓ Плит с резом (сгруппировано): {len(cut_plates)}")
    
    # Пересоздаём plate_assignments в правильном порядке
    # ИСПРАВЛЕНИЕ: добавляем load_code, иначе в треках подставляется 8п и списание ищет 8п вместо 10п
    result['plate_assignments'] = []
    for cut in result['primary_cuts']:
        for length in cut['lengths']:
            result['plate_assignments'].append({
                'length': length,
                'width': cut['width'],
                'source': 'primary',
                'rest_width': cut['rest'],
                'kp_id': cut.get('kp_id'),
                'customer': cut.get('customer'),
                'kp_date': cut.get('kp_date'),
                'plate_name': cut.get('plate_name'),
                'load_code': cut.get('load_code', 800),
            })
            
            # Сохраняем информацию об остатке для отслеживания
            if cut['rest'] > 0:
                result['rests_created'].append({
                    'length': length,
                    'rest_width_mm': cut['rest'],
                    'source_width_mm': cut['width']
                })
    
    print(f"[OPT_2D] ✓ Создано остатков: {len(result['rests_created'])}")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========
    
    # Вторичные резы
    # ИСПРАВЛЕНИЕ: Для каждой плиты из вторичного реза получаем свой kp_id!
    for opt in secondary_options:
        qty = int(round(value(x_sec[opt['id']])))
        if qty > 0:
            target_key = opt.get('target_order_key')
            
            # Добавляем каждый вторичный рез ОТДЕЛЬНО со своим kp_id
            for _ in range(qty):
                # Отмечаем использованный остаток
                result['rests_used'].append({
                    'source_length': opt['source_length'],
                    'source_rest_mm': opt['source_rest']
                })
                
                # Для каждой плиты (pieces) из этого реза получаем свой kp_id
                for _ in range(opt['pieces']):
                    # Получаем информацию о следующем КП
                    plate_info = _get_next_order_info(order_info_list, target_key) if target_key else {}
                    # #region agent log: secondary plate kp_id (H2, H5)
                    if not plate_info and target_key:
                        try:
                            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8') as _f:
                                _f.write(__import__('json').dumps({"hypothesisId": "H2", "location": "optimization.py:secondary_emit", "message": "secondary plate_info empty", "data": {"target_key": list(target_key) if isinstance(target_key, tuple) else target_key, "output_length": opt.get('output_length'), "output_width": opt.get('output_width')}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                        except Exception:
                            pass
                    # #endregion
                    result['secondary_cuts'].append({
                        'source': opt['source_rest'],
                        'cuts': [opt['output_width']],
                        'qty': 1,  # Каждая плита отдельно!
                        'pieces': 1,
                        'waste': opt.get('waste', 0),
                        'type': opt['type'],
                        'source_lengths': [opt['source_length']],
                        'lengths': [opt['output_length']],
                        'target_order_key': target_key,
                        'kp_id': plate_info.get('kp_id') if plate_info else None,
                        'customer': plate_info.get('customer') if plate_info else None,
                        'kp_date': plate_info.get('kp_date') if plate_info else None,
                        'plate_name': plate_info.get('plate_name') if plate_info else None
                    })
                    
                    _sec_load = (target_key[2] if isinstance(target_key, (list, tuple)) and len(target_key) > 2 else 800)
                    result['plate_assignments'].append({
                        'length': opt['output_length'],
                        'width': opt['output_width'],
                        'source': 'secondary',
                        'source_rest': opt['source_rest'],
                        'kp_id': plate_info.get('kp_id') if plate_info else None,
                        'customer': plate_info.get('customer') if plate_info else None,
                        'kp_date': plate_info.get('kp_date') if plate_info else None,
                        'plate_name': plate_info.get('plate_name') if plate_info else None,
                        'load_code': _sec_load,
                    })
    
    print(f"[OPT_2D] OK! Готово! Использовано {result['total_plates']} плит")
    print(f"[OPT_2D] Создано {len(result['plate_assignments'])} готовых плит")
    print(f"[OPT_2D] Остатков использовано вторично: {len(result['rests_used'])}")
    # #region agent log: result counts (H5) + plates by key (H_rescue_trace)
    try:
        open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8').write(__import__('json').dumps({"sessionId": "debug-session", "runId": "run1", "hypothesisId": "H5", "location": "optimization.py:result_built", "message": "plate_assignments count", "data": {"len_plate_assignments": len(result['plate_assignments']), "total_plates": result.get('total_plates', 0), "demand_sum": _total_demand}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # Лог по ключам (length, width, load_code) — для сравнения с дорожками и РЕСКЬЮ (любые размеры).
    try:
        _norm_lc = getattr(cfg, 'normalize_load_code', lambda x: int(x) if x is not None else 8)
        _by_key = {}
        for cut in result.get('primary_cuts', []):
            lc = _norm_lc(cut.get('load_code', 8))
            for L in cut.get('lengths', []):
                k = (round(float(L), 2), int(cut.get('width', 0)), lc)
                _by_key[k] = _by_key.get(k, 0) + 1
        for cut in result.get('secondary_cuts', []):
            lengths = cut.get('lengths', [])
            widths = cut.get('cuts', [])
            tk = cut.get('target_order_key')
            lc = _norm_lc(tk[2] if isinstance(tk, (tuple, list)) and len(tk) > 2 else 8)
            L = float(lengths[0]) if lengths else 0
            W = int(widths[0]) if widths else (int(tk[1]) if isinstance(tk, (tuple, list)) and len(tk) > 1 else 0)
            if L and W:
                k = (round(L, 2), W, lc)
                _by_key[k] = _by_key.get(k, 0) + 1
        open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", 'a', encoding='utf-8').write(__import__('json').dumps({"hypothesisId": "H_opt_plates_by_key", "location": "optimization.py:result_built", "message": "optimizer output plates by (length, width, load_code)", "data": {"plates_by_key": {str(list(k)): v for k, v in _by_key.items()}, "total": sum(_by_key.values())}, "timestamp": __import__('time').time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return result


def _optimize_1d_widths_only(orders: dict, plate_width: int = 1200, 
                              min_useful_width: int = 200) -> dict:
    """
    ПРИВАТНАЯ функция: Оптимизация только по ширинам (1D).
    Длины НЕ учитываются в оптимизации, присваиваются позже.
    
    Args:
        orders: {300: 4, 500: 3} — спрос по ширинам в мм (без учёта длин!)
        plate_width: ширина исходной плиты в мм (1200)
        min_useful_width: минимальная полезная ширина остатка
    
    Returns:
        {
            'primary_cuts': [{'width': 320, 'rest': 880, 'qty': 2}, ...],
            'secondary_cuts': [{'source': 880, 'cuts': [320, 560], 'qty': 1}, ...],
            'total_plates': 5,
            'total_cost': 75000,
            'waste_width': 120
        }
    """
    try:
        from pulp import LpProblem, LpMinimize, LpVariable, LpInteger, lpSum, value, PULP_CBC_CMD
    except ImportError:
        print('[OPT_CASCADING] PuLP не установлен, пропускаем.')
        return {}
    
    if not orders:
        return {}
    
    # Преобразуем заказы в список ширин с количеством
    target_widths = sorted(orders.keys())
    
    # Допустимый диапазон для каждой ширины (±20 мм)
    tolerance = 20
    
    # Генерируем все возможные варианты первичных резов (из плиты заданной ширины)
    # Для каждой целевой ширины создаём варианты: target_width + остаток
    primary_cut_options = []
    solid_widths = sorted(set([plate_width, 1080]))
    for target_w in target_widths:
        if target_w in solid_widths:
            primary_cut_options.append({
                'id': f'prim_{target_w}',
                'main': target_w,
                'rest': 0,
            })
            continue

        # Пропил косвенно учтён в таблице NARROWING_TABLE
        rest_w = plate_width - target_w
        if rest_w >= min_useful_width:  # Остаток достаточно большой
            primary_cut_options.append({
                'id': f'prim_{target_w}',
                'main': target_w,
                'rest': rest_w,
            })
    
    # Генерируем варианты вторичных резов (из остатков)
    # Для каждого возможного остатка смотрим, на какие ширины его можно разрезать
    secondary_cut_options = []
    possible_rests = set(opt['rest'] for opt in primary_cut_options if opt['rest'] > 0)
    
    for rest_w in possible_rests:
        # Пробуем разрезать остаток на 2 части
        for target_w1 in target_widths:
            target_w2 = rest_w - target_w1
            # Проверяем, подходит ли вторая часть для какой-то из целевых ширин
            for target_w2_candidate in target_widths:
                if abs(target_w2 - target_w2_candidate) <= tolerance:
                    secondary_cut_options.append({
                        'id': f'sec_{rest_w}_to_{target_w1}_{target_w2_candidate}',
                        'source_rest': rest_w,
                        'output1': target_w1,
                        'output2': target_w2_candidate,
                        'waste': abs(rest_w - target_w1 - target_w2_candidate),
                    })
                    break
            
            # Также проверяем вариант: остаток целиком режем на несколько одинаковых частей
            # Например, 880 мм → 2 части по 320 мм (остаток 240 мм)
            for target_w_candidate in target_widths:
                # Считаем, сколько частей нужной ширины влезет в остаток
                num_pieces = rest_w // target_w_candidate
                if num_pieces >= 2:  # Минимум 2 куска, иначе не выгодно
                    waste = rest_w - (target_w_candidate * num_pieces)
                    if waste < rest_w * 0.5:  # Отход < 50% остатка (разумное ограничение)
                        secondary_cut_options.append({
                            'id': f'sec_{rest_w}_to_{num_pieces}x{target_w_candidate}',
                            'source_rest': rest_w,
                            'output1': target_w_candidate,
                            'output2': 0,  # вторая "часть" не используется
                            'pieces': num_pieces,
                            'waste': waste,
                        })
                        # Не ставим break - проверяем все варианты
            
            # НОВАЯ ЛОГИКА: Сужение (narrowing) - из остатка делаем ОДНУ плиту с небольшим отходом
            # Например, 340 мм → 320 мм (отход 20 мм)
            for target_w_candidate in target_widths:
                if target_w_candidate < rest_w <= target_w_candidate + 100:  # Остаток чуть больше целевой ширины
                    waste = rest_w - target_w_candidate
                    # Разрешаем сужение до 100 мм (но лучше меньше)
                    if waste <= 100:
                        secondary_cut_options.append({
                            'id': f'sec_{rest_w}_narrow_to_{target_w_candidate}',
                            'source_rest': rest_w,
                            'output1': target_w_candidate,
                            'output2': 0,
                            'pieces': 1,  # Только ОДНА плита
                            'waste': waste,
                        })
    
    print(f"[DEBUG] Найдено вариантов вторичных резов: {len(secondary_cut_options)}")
    for opt in secondary_cut_options[:3]:  # Показываем первые 3
        print(f"  {opt['source_rest']}мм -> {opt.get('pieces', 1)}x{opt['output1']}мм (отход {opt['waste']}мм)")
    
    # Создаём задачу оптимизации
    prob = LpProblem('cascading_longitudinal_cuts', LpMinimize)
    
    # Переменные: количество первичных резов каждого типа
    x_prim = {}
    for i, opt in enumerate(primary_cut_options):
        x_prim[i] = LpVariable(f"prim_{i}_{opt['id']}", lowBound=0, cat=LpInteger)
    
    # Переменные: количество вторичных резов каждого типа
    x_sec = {}
    for i, opt in enumerate(secondary_cut_options):
        x_sec[i] = LpVariable(f"sec_{i}_{opt['id']}", lowBound=0, cat=LpInteger)
    
    # Ограничение 1: Покрыть спрос по каждой ширине
    for target_w, qty in orders.items():
        sources = []
        
        # Источник 1: Основные части первичных резов
        for i, opt in enumerate(primary_cut_options):
            if abs(opt['main'] - target_w) <= tolerance:
                sources.append(x_prim[i])
        
        # Источник 2: Выходы из вторичных резов (output1)
        for i, opt in enumerate(secondary_cut_options):
            if abs(opt['output1'] - target_w) <= tolerance:
                pieces = opt.get('pieces', 1)
                sources.append(x_sec[i] * pieces)
        
        # Источник 3: Выходы из вторичных резов (output2, если есть)
        for i, opt in enumerate(secondary_cut_options):
            if opt['output2'] > 0 and abs(opt['output2'] - target_w) <= tolerance:
                sources.append(x_sec[i])
        
        if sources:
            prob += lpSum(sources) >= qty, f"demand_{target_w}"
    
    # Ограничение 2: Баланс остатков (не можем резать больше, чем произвели)
    for rest_w in possible_rests:
        # Сколько создали остатков rest_w (первичными резами)
        produced = []
        for i, opt in enumerate(primary_cut_options):
            if opt['rest'] == rest_w:
                produced.append(x_prim[i])
        
        # Сколько используем для вторичных резов
        consumed = []
        for i, opt in enumerate(secondary_cut_options):
            if opt['source_rest'] == rest_w:
                consumed.append(x_sec[i])
        
        if produced and consumed:
            prob += lpSum(consumed) <= lpSum(produced), f"balance_rest_{rest_w}"
    
    # Целевая функция: УМНАЯ ОПТИМИЗАЦИЯ
    # Приоритет 1: Минимум плит (главное!)
    # Приоритет 2: Минимум неиспользованных остатков (важно!)
    # Приоритет 3: Минимум отходов (хорошо, но не критично)
    
    total_plates = lpSum(x_prim.values())
    
    # Считаем неиспользованные остатки для каждого типа
    # (остатки, которые создали, но НЕ пустили во вторичные резы)
    unused_rests_penalty = 0
    for rest_w in possible_rests:
        # Сколько создали остатков
        produced = []
        for i, opt in enumerate(primary_cut_options):
            if opt['rest'] == rest_w:
                produced.append(x_prim[i])
        
        # Сколько использовали для вторичных резов
        consumed = []
        for i, opt in enumerate(secondary_cut_options):
            if opt['source_rest'] == rest_w:
                consumed.append(x_sec[i])
        
        if produced and consumed:
            # Неиспользованные остатки = произведено - использовано
            unused = lpSum(produced) - lpSum(consumed)
            # Штраф за каждый неиспользованный остаток
            # Больше остаток = больше штраф (в пропорции к ширине)
            unused_rests_penalty += unused * (rest_w / 1000.0) * 0.05
    
    # Штраф за отходы от вторичных резов (суммарные мм)
    waste_penalty = 0
    for i, opt in enumerate(secondary_cut_options):
        waste_penalty += x_sec[i] * opt.get('waste', 0) * 0.0001
    
    # ИТОГОВАЯ ЦЕЛЕВАЯ ФУНКЦИЯ:
    # 1. total_plates - доминирует (вес = 1.0)
    # 2. unused_rests_penalty - влияет, но слабее плит (вес ~0.05)
    # 3. waste_penalty - минимальное влияние (вес ~0.0001)
    # Оптимизатор САМ выберет, что выгоднее: множественная резка или сужение!
    prob += total_plates + unused_rests_penalty + waste_penalty
    prob.solve(PULP_CBC_CMD(msg=0))
    
    # Извлекаем результаты
    result = {
        'primary_cuts': [],
        'secondary_cuts': [],
        'total_plates': 0,
        'total_cost': 0,
        'waste_width': 0,
    }
    
    for i, opt in enumerate(primary_cut_options):
        try:
            qty = int(round(value(x_prim[i])))
            if qty > 0:
                result['primary_cuts'].append({
                    'width': opt['main'],
                    'rest': opt['rest'],
                    'qty': qty,
                })
                result['total_plates'] += qty
        except:
            pass
    
    for i, opt in enumerate(secondary_cut_options):
        try:
            qty = int(round(value(x_sec[i])))
            if qty > 0:
                cuts = [opt['output1']]
                if opt['output2'] > 0:
                    cuts.append(opt['output2'])
                result['secondary_cuts'].append({
                    'source': opt['source_rest'],
                    'cuts': cuts,
                    'pieces': opt.get('pieces', 1),
                    'qty': qty,
                    'waste': opt.get('waste', 0),
                })
                result['waste_width'] += opt.get('waste', 0) * qty
        except:
            pass
    
    # ========== НОВАЯ ЛОГИКА: СОРТИРОВКА ДЛЯ ПРОИЗВОДСТВА ==========
    # Требования завода:
    # 1. Первая плита ДОЛЖНА быть целой (без реза, rest=0)
    # 2. Плиты с одинаковым резом должны идти подряд
    print("[OPT_1D] 🔧 Применяем правила завода для порядка плит...")
    
    # Разделяем primary_cuts на группы
    solid_plates = []    # Целые плиты (rest=0)
    cut_plates = []      # Плиты с резом (rest>0)
    
    for cut in result['primary_cuts']:
        if cut['rest'] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)
    
    # ВАЖНО: Сортируем ЦЕЛЫЕ плиты тоже! (чтобы они не перемешивались с резанными)
    # Сортируем по ширине (целые 1200мм все вместе), потом по first элементу
    if solid_plates:
        solid_plates.sort(key=lambda x: (-x['width']))
    
    # Сортируем плиты с резом по (rest, width) — одинаковые резы идут подряд
    # Сортируем по убыванию ширины остатка, чтобы крупные резы шли в начале
    cut_plates.sort(key=lambda x: (-x['rest'], -x['width']))
    
    # Новый порядок: СНАЧАЛА целые (ОТСОРТИРОВАННЫЕ!), ПОТОМ плиты с резом (сгруппированные)
    result['primary_cuts'] = solid_plates + cut_plates
    
    print(f"[OPT_1D] ✓ Целых плит в начале: {len(solid_plates)}")
    print(f"[OPT_1D] ✓ Плит с резом (сгруппировано): {len(cut_plates)}")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========
    
    print(f"[DEBUG] Оптимизатор выбрал вторичных резов: {len(result['secondary_cuts'])}")
    
    # Рассчитываем стоимость
    plate_price = 12000  # примерная цена плиты
    long_cut_cost = 460 * 6  # продольный рез (460 руб/м × 6 м средняя длина)
    
    result['total_cost'] = (
        result['total_plates'] * plate_price +
        result['total_plates'] * long_cut_cost +  # первичные резы
        len(result['secondary_cuts']) * long_cut_cost  # вторичные резы
    )
    
    return result


def optimize_with_cascading_longitudinal_cuts(orders: dict = None, 
                                               orders_2d: list = None,
                                               plate_width: int = 1200, 
                                               min_useful_width: int = 200,
                                               opt_config: OptimizationConfig = None) -> dict:
    """
    Универсальная оптимизация с каскадными резами (PUBLIC API).
    
    АВТОМАТИЧЕСКИ ВЫБИРАЕТ РЕЖИМ на основе входных данных:
    
    Режим 1D (старый, обратная совместимость):
        >>> result = optimize_with_cascading_longitudinal_cuts(
        ...     orders={320: 14, 860: 9}
        ... )
        # Оптимизирует по ШИРИНАМ, длины присваиваются позже
    
    Режим 2D (новый, полная оптимизация):
        >>> result = optimize_with_cascading_longitudinal_cuts(
        ...     orders_2d=[
        ...         {'length': 5.6, 'width': 320, 'qty': 11},
        ...         {'length': 6.63, 'width': 860, 'qty': 4}
        ...     ]
        ... )
        # Полная 2D оптимизация (длина + ширина) в ILP модели
    
    Args:
        orders: {width: qty} — спрос по ширинам (для режима 1D)
        orders_2d: [{'length', 'width', 'qty'}] — спрос 2D (для режима 2D)
        plate_width: ширина исходной плиты (1200 мм)
        min_useful_width: минимальная полезная ширина остатка
        opt_config: конфигурация параметров оптимизации (штрафы, бонусы)
    
    Returns:
        dict с результатами оптимизации:
            {
                'primary_cuts': [...],
                'secondary_cuts': [...],
                'total_plates': int,
                'plate_assignments': [...],  # только для режима 2D
                ...
            }
    """
    
    # АВТООПРЕДЕЛЕНИЕ РЕЖИМА
    if orders_2d is not None and len(orders_2d) > 0:
        # ===== РЕЖИМ 2D (НОВЫЙ) =====
        print("[OPT] Режим: ПОЛНАЯ 2D оптимизация (длина + ширина)")
        return _optimize_2d_with_lengths(orders_2d, plate_width, min_useful_width, opt_config)
    
    elif orders is not None and len(orders) > 0:
        # ===== РЕЖИМ 1D (СТАРЫЙ) =====
        print("[OPT] Режим: 1D оптимизация (только ширина, обратная совместимость)")
        return _optimize_1d_widths_only(orders, plate_width, min_useful_width)
    
    else:
        print("[OPT] ⚠️ Не указаны ни orders, ни orders_2d!")
        return {}


# ==================== FFD ОПТИМИЗАЦИЯ РАСКРОЯ ДОРОЖЕК ====================

from dataclasses import dataclass, field


@dataclass
class Piece:
    """Кусок плиты для укладки в дорожку"""
    length_m: float
    qty: int
    kind: str              # 'standard' | 'addon'
    load_class: float
    width_m: float = 1.196


@dataclass
class Track:
    """Дорожка (линия производства)"""
    width_m: float = 1.196
    total_m: float = 0.0
    pieces: list = field(default_factory=list)
    leftover_m: float = 0.0


def first_fit_decreasing(
    pieces: list[Piece],
    stock_len_m: float = 9.88
) -> list[Track]:
    """
    Алгоритм First Fit Decreasing для оптимизации раскроя
    Минимизирует количество дорожек (плит-заготовок)
    
    Args:
        pieces: Список Piece объектов (куски для размещения)
        stock_len_m: Длина заготовки (максимальная длина плиты)
        
    Returns:
        Список Track объектов (дорожек) с размещёнными кусками
    """
    pool = []
    
    # Сортируем по убыванию длины (FFD алгоритм)
    sorted_pieces = sorted(pieces, key=lambda x: x.length_m, reverse=True)
    
    # Развёртываем количество в отдельные элементы
    expanded = []
    for p in sorted_pieces:
        for _ in range(p.qty):
            expanded.append(Piece(p.length_m, 1, p.kind, p.load_class, p.width_m))
    
    # Размещаем каждый кусок
    for piece in expanded:
        placed = False
        
        # Пробуем поместить в существующие дорожки
        for track in pool:
            if track.total_m + piece.length_m <= stock_len_m:
                track.pieces.append(piece)
                track.total_m += piece.length_m
                placed = True
                break
        
        # Если не поместился, создаём новую дорожку
        if not placed:
            track = Track()
            track.pieces.append(piece)
            track.total_m = piece.length_m
            pool.append(track)
    
    # Вычисляем остатки
    for track in pool:
        track.leftover_m = stock_len_m - track.total_m
    
    return pool


def optimize_tracks(
    items: list,
    stock_len_m: float = 9.88
) -> dict:
    """
    Оптимизирует размещение плит в дорожки
    
    Args:
        items: Список позиций [{'length_m': float, 'qty': int, 'kind': str, 'load_class': float}, ...]
        stock_len_m: Длина заготовки (максимальная длина)
        
    Returns:
        Словарь с результатами оптимизации
    """
    pieces = []
    
    for item in items:
        pieces.append(Piece(
            length_m=item.get('length_m', 0),
            qty=item.get('qty', 1),
            kind=item.get('kind', 'standard'),
            load_class=item.get('load_class', 8.0),
            width_m=item.get('width_m', 1.196)
        ))
    
    tracks = first_fit_decreasing(pieces, stock_len_m)
    
    # Статистика
    total_tracks = len(tracks)
    total_used = sum(t.total_m for t in tracks)
    total_leftover = sum(t.leftover_m for t in tracks)
    efficiency = (total_used / (total_tracks * stock_len_m) * 100) if total_tracks > 0 else 0
    
    return {
        'tracks': tracks,
        'total_tracks': total_tracks,
        'total_used_m': round(total_used, 2),
        'total_leftover_m': round(total_leftover, 2),
        'efficiency_pct': round(efficiency, 1),
        'stock_length_m': stock_len_m
    }


