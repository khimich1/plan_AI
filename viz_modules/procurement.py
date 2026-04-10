#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль построения закупки и сметы:
- Формирование позиций закупки
- Построение строк сметы
- Детальная разбивка компонентов
"""
import re
from collections import Counter

import core.config_and_data as cfg
from core.optimization import OPT_PLAN
from core.price_db import get_price
from .price_utils import find_price_for_plate

WIDE_WIDTH_M = 1.2
WIDE_EPS = 1e-6


def _is_wide_width(width_m: float, *, threshold_m: float = WIDE_WIDTH_M, eps: float = WIDE_EPS) -> bool:
    """True для плит шире 12 дм (> 1.2 м)."""
    return float(width_m) > (threshold_m + eps)


def _block_ab_key(width_m: float) -> int:
    """Ключ двухблочной схемы: 0=Блок A (обычные), 1=Блок B (широкие)."""
    return 1 if _is_wide_width(width_m) else 0


def get_orders_from_opt_plan():
    """
    Возвращает исходный заказ (length/width/qty), сохранённый оптимизатором.
    
    ПРИОРИТЕТ:
    1. Собирает заказы из всех планов в OPT_CASCADING_PLAN_BY_LOAD (планы по нагрузкам)
    2. Если нет, берёт из OPT_CASCADING_PLAN (общий план)
    3. Если нет, возвращает None
    """
    try:
        from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    except ImportError:
        return None

    orders_copy = []
    
    # Приоритет 1: Собираем заказы из всех планов по нагрузкам
    if OPT_CASCADING_PLAN_BY_LOAD:
        print(f'[DEBUG] get_orders_from_opt_plan: Найдено {len(OPT_CASCADING_PLAN_BY_LOAD)} планов по нагрузкам')
        for load_key, plan in OPT_CASCADING_PLAN_BY_LOAD.items():
            if plan and plan.get('orders_requested'):
                print(f'[DEBUG] Извлекаем заказы из плана для нагрузки {load_key}п: {len(plan["orders_requested"])} позиций')
                for order in plan['orders_requested']:
                    try:
                        orders_copy.append({
                            'length': float(order.get('length', 0)),
                            'width': order.get('width', 0),
                            'qty': int(order.get('qty', 1)),
                            'load_code': order.get('load_code'),
                            'length_dm_raw': (order.get('length_dm_raw') or '').strip(),
                        })
                    except Exception as e:
                        print(f'[DEBUG] Ошибка парсинга заказа: {e}')
                        continue
        if orders_copy:
            print(f'[DEBUG] ✅ Собрано {len(orders_copy)} заказов из планов по нагрузкам')
            return orders_copy
    
    # Приоритет 2: Fallback на общий план (если BY_LOAD пуст)
    plan = OPT_CASCADING_PLAN
    if plan and plan.get('orders_requested'):
        print(f'[DEBUG] get_orders_from_opt_plan: Используем общий план OPT_CASCADING_PLAN')
        for order in plan['orders_requested']:
            try:
                orders_copy.append({
                    'length': float(order.get('length', 0)),
                    'width': order.get('width', 0),
                    'qty': int(order.get('qty', 1)),
                    'load_code': order.get('load_code'),
                    'length_dm_raw': (order.get('length_dm_raw') or '').strip(),
                })
            except Exception:
                continue
        if orders_copy:
            print(f'[DEBUG] ✅ Собрано {len(orders_copy)} заказов из общего плана')
            return orders_copy
    
    print('[DEBUG] ⚠️ get_orders_from_opt_plan: Нет данных ни в BY_LOAD, ни в общем плане')
    return None


def _length_dm_raw_from_m(length_m: float) -> str:
    """Вычисляет строку длины в дм (как в марке) из длины в метрах.

    Зеркало логики make_plate_name при отсутствии length_dm_raw:
    5.71 → '57,1', 5.7 → '57', 6.88 → '68,8'.
    """
    length_dm_val = length_m * 10
    if abs(length_dm_val - round(length_dm_val)) < 0.01:
        return str(int(round(length_dm_val)))
    return f'{length_dm_val:.1f}'.rstrip('0').rstrip('.').replace('.', ',')


def _is_same_length(lengths: list, target_len: float, tolerance: float = 0.05) -> bool:
    """Проверяет, подходит ли операция под длину плиты."""
    if not lengths:
        return True
    return any(abs(float(v) - target_len) < tolerance for v in lengths)


def _find_plan_for_plate(load_code: int, length: float, width_mm: int, name: str, debug_tag: str):
    """Ищет план оптимизации для конкретной плиты по нагрузке/длине/ширине."""
    from core.optimization import OPT_CASCADING_PLAN_BY_LOAD, LOAD_TO_REINFORCEMENT_MAP
    import math

    current_plan = None
    load_key = int(math.floor(load_code)) if isinstance(load_code, (int, float)) else 8

    if OPT_CASCADING_PLAN_BY_LOAD and LOAD_TO_REINFORCEMENT_MAP and load_key in LOAD_TO_REINFORCEMENT_MAP:
        for reinforcement_key in LOAD_TO_REINFORCEMENT_MAP[load_key]:
            plan = OPT_CASCADING_PLAN_BY_LOAD.get(reinforcement_key)
            if not plan:
                continue

            orders_req = plan.get('orders_requested') or []
            for ord_item in orders_req:
                try:
                    o_len = float(ord_item.get('length', 0))
                    o_width = int(ord_item.get('width', 0))
                    o_load = ord_item.get('load_code', load_key)
                except Exception:
                    continue

                if (
                    abs(o_len - length) < 0.05 and
                    o_width == width_mm and
                    int(math.floor(float(o_load))) == load_key
                ):
                    current_plan = plan
                    print(
                        f'[DEBUG] {debug_tag}: нашёл план для {name} — '
                        f'нагрузка {load_key}п, армирование {reinforcement_key}'
                    )
                    break
            if current_plan:
                break
    return current_plan, load_key


def _calc_trim_components(
    current_plan: dict | None,
    *,
    length: float,
    width_mm: int,
    qty: int,
    base_price_1_2m: float,
    base_price: float,
    load_code: int,
    price_table: dict,
):
    """
    Единый расчёт резов/остатков/отходов.
    Правило: непереиспользованный первичный обрезок идёт в `rest_cost`,
    вторичные отходы (`waste`/`length_waste`) — в `waste_cost`.
    """
    rest_cost = 0.0
    rest_width_mm = 1200 - width_mm if width_mm < 1150 else 0
    rest_used = False
    waste_cost = 0.0
    waste_terms = []
    trans_cuts = 0.0
    total_cuts_for_this_size = 0
    total_plates_from_cuts = 0

    if current_plan and current_plan.get('primary_cuts'):
        primary_rest_width_mm = 0

        for prim_cut in current_plan['primary_cuts']:
            if prim_cut.get('width') != width_mm:
                continue
            if not _is_same_length(prim_cut.get('lengths', []), length):
                continue

            prim_qty = int(prim_cut.get('qty', 0) or 0)
            total_cuts_for_this_size += prim_qty
            total_plates_from_cuts += prim_qty
            primary_rest_width_mm = int(prim_cut.get('rest', 0) or 0)

            if primary_rest_width_mm > 0:
                produced_rests = prim_qty
                used_rests = 0
                for sec_cut in (current_plan.get('secondary_cuts') or []):
                    if int(sec_cut.get('source', 0) or 0) != primary_rest_width_mm:
                        continue
                    if not _is_same_length(sec_cut.get('source_lengths', []), length):
                        continue
                    used_rests += int(sec_cut.get('qty', 0) or 0)

                unused_rests = max(0, produced_rests - used_rests)
                unused_rest_total_mm = unused_rests * primary_rest_width_mm
                if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                    rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                    rest_width_mm = unused_rest_total_mm
                elif produced_rests > 0:
                    rest_used = True
                    rest_width_mm = 0

            # secondary_cuts: основной матч по целевой ширине + fallback по source/rest
            secondary_matches = 0
            sec_skip_length = 0
            sec_skip_width = 0
            for sec_cut in (current_plan.get('secondary_cuts') or []):
                if not _is_same_length(sec_cut.get('lengths', []), length):
                    sec_skip_length += 1
                    continue

                sec_cuts = sec_cut.get('cuts', []) or []
                width_match = any(abs(width_mm - int(cut_width)) <= 20 for cut_width in sec_cuts)
                source_match = (
                    primary_rest_width_mm > 0 and
                    int(sec_cut.get('source', 0) or 0) == primary_rest_width_mm and
                    _is_same_length(sec_cut.get('source_lengths', []), length)
                )
                if not (width_match or source_match):
                    sec_skip_width += 1
                    continue

                secondary_matches += 1
                sec_qty = int(sec_cut.get('qty', 0) or 0)
                sec_pieces = int(sec_cut.get('pieces', 1) or 1)
                current_cuts = sec_qty * sec_pieces
                total_cuts_for_this_size += current_cuts
                total_plates_from_cuts += current_cuts

                waste_w_mm = float(sec_cut.get('waste', 0) or 0)
                if waste_w_mm > 0 and base_price_1_2m > 0 and qty > 0:
                    cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price_1_2m
                    waste_cost += (cost_of_waste_piece * sec_qty) / qty
                    waste_terms.append((waste_w_mm, sec_qty))

                src_lens = sec_cut.get('source_lengths', [])
                if src_lens:
                    src_len = float(src_lens[0])
                    if sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05:
                        if qty > 0:
                            trans_cuts += (1.0 * sec_qty) / qty

                        len_waste = src_len - length
                        if len_waste > 0.01:
                            src_price_full = get_price(src_len, load_code, cfg.PRICE_DB_PATH)
                            if src_price_full is None:
                                src_price_full = find_price_for_plate(price_table, src_len, load_code) or 0.0
                            src_price_width = src_price_full * (width_mm / 1200.0)
                            cost_len_waste = src_price_width - base_price
                            if cost_len_waste > 0 and qty > 0:
                                waste_cost += cost_len_waste * (sec_qty / qty)

            if secondary_matches == 0 and (sec_skip_length > 0 or sec_skip_width > 0):
                print(
                    f'[DEBUG] trim_match: {length}x{width_mm}мм '
                    f'secondary_cuts не сматчены (skip_length={sec_skip_length}, skip_width={sec_skip_width})'
                )
            break

    return {
        'rest_cost': rest_cost,
        'rest_width_mm': rest_width_mm,
        'rest_used': rest_used,
        'waste_cost': waste_cost,
        'waste_terms': waste_terms,
        'trans_cuts': trans_cuts,
        'total_cuts_for_this_size': total_cuts_for_this_size,
        'total_plates_from_cuts': total_plates_from_cuts,
    }


def build_procurement_items():
    """Формирует реальные позиции закупки из заказа пользователя."""
    items = []

    plan_orders = get_orders_from_opt_plan()
    if plan_orders:
        order_counter = Counter()
        for order in plan_orders:
            length = round(float(order['length']), 3)
            width_val = order['width']
            # В заказе ширина приходит в мм (например, 320). Преобразуем в метры.
            width_m = width_val / 1000.0 if width_val > 5 else float(width_val)
            ldr = (order.get('length_dm_raw') or '').strip()

            # ПРИОРИТЕТ 1: Если load_code уже пришёл из плана оптимизации - используем его напрямую
            if order.get('load_code') is not None:
                load_code = order['load_code']
                order_counter[(length, width_m, load_code, ldr, False)] += order['qty']
            else:
                # ПРИОРИТЕТ 2 (для обратной совместимости): Пытаемся найти в PLATE_LOAD_DETAILS
                found_details = []
                if cfg.PLATE_LOAD_DETAILS:
                    for key, q in cfg.PLATE_LOAD_DETAILS.items():
                        L, W, load = key[0], key[1], key[2]
                        key_ldr = key[3] if len(key) > 3 else ''
                        if abs(L - length) < 0.05 and abs(W - width_m) < 0.01:
                            found_details.append((load, q, key_ldr))

                total_found = sum(q for _, q, _ in found_details)

                if found_details and total_found == order['qty']:
                    for load_code, q, detail_ldr in found_details:
                        order_counter[(length, width_m, load_code, detail_ldr, False)] += q
                else:
                    load_code = cfg.get_load_code_for_plate(length, width_m, default=(6 if width_m < 1.0 else 8))
                    warning_flag = True if found_details else False
                    if found_details:
                        print(f"[WARNING] Несовпадение количества для плиты {length}x{width_m}: "
                              f"В плане {order['qty']}, в деталях {total_found}. Присвоена нагрузка {load_code}п (проверьте!)")
                    order_counter[(length, width_m, load_code, ldr, warning_flag)] += order['qty']

        for (length, width_m, load_code, ldr, warning_flag), qty in sorted(
            order_counter.items(),
            key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0], x[0][2]),
        ):
            if abs(width_m - 1.2) < 0.01:
                long_cuts = 0
            else:
                long_cuts = 1 if width_m < 1.15 else 0

            # Берём length_dm_raw из ключа счётчика (уже содержит номинал из заказа)
            full_key = (length, width_m, load_code, ldr)
            length_dm_raw_val = ldr or cfg.PLATE_LENGTH_DM_RAW.get(full_key, '') or _length_dm_raw_from_m(length)
            items.append({
                'length': length,
                'width': width_m,
                'qty': qty,
                'long_cuts': long_cuts,
                'trans_cuts': 0,
                'load_code': load_code,
                'warning': warning_flag,
                'length_dm_raw': length_dm_raw_val,
            })
        return items

    # Приоритет 2: Используем PLATE_LOAD_DETAILS (если есть) или реальный заказ из cfg.PLATES_*
    # PLATE_LOAD_DETAILS содержит (длина, ширина, нагрузка) → количество
    if cfg.PLATE_LOAD_DETAILS:
        # Используем PLATE_LOAD_DETAILS напрямую - там уже правильно разделены нагрузки!
        for key, qty in sorted(
            cfg.PLATE_LOAD_DETAILS.items(),
            key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0], x[0][2]),
        ):
            length, width_m, load_code = key[0], key[1], key[2]
            ldr = key[3] if len(key) > 3 else cfg.PLATE_LENGTH_DM_RAW.get(key, '')

            if abs(width_m - 1.2) < 0.01:
                long_cuts = 0
            else:
                long_cuts = 1 if width_m < 1.15 else 0

            item = {
                'length': length,
                'width': width_m,
                'qty': qty,
                'long_cuts': long_cuts,
                'trans_cuts': 0,
                'load_code': load_code,
                'length_dm_raw': ldr or _length_dm_raw_from_m(length),
            }
            # Подставляем canonical_name и nomenclature_id из кэша (4-кортежный ключ)
            cached = cfg.PLATE_NOMENCLATURE_CACHE.get(key)
            if cached:
                if cached.get('canonical_name') is not None:
                    item['canonical_name'] = cached['canonical_name']
                if cached.get('nomenclature_id') is not None:
                    item['nomenclature_id'] = cached['nomenclature_id']
            items.append(item)
        return items
    
    # Legacy режим: Используем cfg.PLATES_* (если PLATE_LOAD_DETAILS пуст)
    all_plates = []
    # ВАЖНО: Добавлен target_name для получения точных ширин из PLATE_EXACT_WIDTHS
    for width_mm, plates_list, target_name in [
        (320, cfg.PLATES_0_32, 'PLATES_0_32'), (460, cfg.PLATES_0_46, 'PLATES_0_46'), (700, cfg.PLATES_0_70, 'PLATES_0_70'),
        (720, cfg.PLATES_0_72, 'PLATES_0_72'), (860, cfg.PLATES_0_86, 'PLATES_0_86'), (880, cfg.PLATES_0_88, 'PLATES_0_88'),
        (740, cfg.PLATES_0_74, 'PLATES_0_74'), (480, cfg.PLATES_0_48, 'PLATES_0_48'), (500, cfg.PLATES_0_50, 'PLATES_0_50'),
        (340, cfg.PLATES_0_34, 'PLATES_0_34'), (1080, cfg.PLATES_1_08, 'PLATES_1_08'), (1200, cfg.PLATES_1_2, 'PLATES_1_2'),
        (1000, cfg.PLATES_1_0, 'PLATES_1_0')
    ]:
        if plates_list:
            length_counts = Counter(plates_list)
            for length, qty in length_counts.items():
                # Получаем ТОЧНУЮ ширину из PLATE_EXACT_WIDTHS
                exact_width_m = cfg.get_exact_width(length, target_name, width_mm / 1000.0)
                
                # Получаем нагрузку для этой плиты
                load_code = cfg.get_load_code_for_plate(length, exact_width_m, default=(6 if exact_width_m < 1.0 else 8))
                
                all_plates.append({
                    'length': length,
                    'width': exact_width_m,  # ТОЧНАЯ ширина в метрах!
                    'qty': qty,
                    'load_code': load_code  # Сохраняем нагрузку
                })
    
    if all_plates:
        # Группируем плиты по (длина, ширина, НАГРУЗКА) чтобы не объединять 8п и 10п
        plate_groups = Counter()
        for plate in all_plates:
            length = plate['length']
            width_m = plate['width']
            load_code = plate.get('load_code', 8)
            qty = plate['qty']
            plate_groups[(length, width_m, load_code)] += qty
        
        # Формируем items с учетом нагрузки
        for (length, width_m, load_code), qty in sorted(
            plate_groups.items(),
            key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0], x[0][2]),
        ):
            # Жёсткое правило: плиты 1.2 м считаем целыми, без продольных резов
            if abs(width_m - 1.2) < 0.01:
                long_cuts = 0
            else:
                # Продольные резы: если ширина < 1.2м, значит был рез
                long_cuts = 1 if width_m < 1.15 else 0

            # Поперечные резы: пока 0, они учтены в оптимизации
            trans_cuts = 0

            items.append({
                'length': length,
                'width': width_m,
                'qty': qty,
                'long_cuts': long_cuts,
                'trans_cuts': trans_cuts,
                'load_code': load_code,
                'length_dm_raw': _length_dm_raw_from_m(length),
            })

        return items
    
    # Приоритет 3: Используем старый OPT_PLAN (если нет заказа)
    if OPT_PLAN and OPT_PLAN.get('actions'):
        for act in OPT_PLAN['actions']:
            src_type, W1, W2, L, qty, lc, tc = act
            W1_m = W1 / 1000.0; W2_m = W2 / 1000.0 if W2 else 0
            if src_type == 'split':
                items.append({'length': round(L, 2), 'width': 1.2, 'qty': qty, 'long_cuts': lc, 'trans_cuts': tc, 'purpose': 'split_source'})
            elif src_type == 'narrow':
                items.append({'length': round(L, 2), 'width': W2_m, 'qty': qty, 'long_cuts': lc, 'trans_cuts': tc, 'purpose': 'narrow_source'})
            elif src_type == 'solid':
                items.append({'length': round(L, 2), 'width': W1_m, 'qty': qty, 'long_cuts': lc, 'trans_cuts': tc, 'purpose': 'solid'})
        agg = {}
        for it in items:
            key = (it['length'], it['width'], it['long_cuts'], it['trans_cuts'])
            agg[key] = agg.get(key, 0) + it['qty']
        result = []
        for (L, W, long_cuts, trans_cuts), qty in sorted(
            agg.items(),
            key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0]),
        ):
            result.append({'length': L, 'width': W, 'qty': qty, 'long_cuts': long_cuts, 'trans_cuts': trans_cuts})
        return result
    
    # Fallback: старая логика
    def mismatch_count(main_list, pair_demand):
        if not main_list or not pair_demand:
            return 0
        a = sorted(round(x, 2) for x in main_list)
        b = sorted(round(x, 2) for x in pair_demand)
        i = j = matches = 0
        while i < len(a) and j < len(b):
            if abs(a[i] - b[j]) <= 0.05:
                matches += 1; i += 1; j += 1
            elif a[i] < b[j]:
                i += 1
            else:
                j += 1
        return max(0, min(len(main_list), len(pair_demand)) - matches)

    pair_plan = {
        '0.32': mismatch_count(cfg.PLATES_0_32, cfg.PLATES_0_88),
        '0.46': mismatch_count(cfg.PLATES_0_46, cfg.PLATES_0_74),
        '0.72': mismatch_count(cfg.PLATES_0_72, cfg.PLATES_0_48),
        '0.70': mismatch_count(cfg.PLATES_0_70, cfg.PLATES_0_50),
        '0.86': mismatch_count(cfg.PLATES_0_86, cfg.PLATES_0_34),
    }
    
    # Округление до 3 знаков: сохраняем 5.71 (ПБ 57,1), не схлопываем в 5.7 (ПБ 57)
    _r = lambda x: round(x, 3)
    for L in cfg.PLATES_1_2:
        items.append({'length': _r(L), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'as_is'})
    for L in cfg.PLATES_1_5_TO_1_2:
        items.append({'length': _r(L), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'to_1_2_main'})
        items.append({'length': _r(L), 'width': 0.3, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_2_strip'})
    for L in cfg.PLATES_1_0:
        items.append({'length': _r(L), 'width': 1.0, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_main'})
        items.append({'length': _r(L), 'width': 0.2, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_strip'})
    for L in cfg.PLATES_1_08:
        items.append({'length': _r(L), 'width': 1.08, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_main'})
        items.append({'length': _r(L), 'width': 0.12, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_strip'})
    for L in cfg.PLATES_0_46:
        items.append({'length': _r(L), 'width': 0.46, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_main'})
        items.append({'length': _r(L), 'width': 0.74, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_strip'})
    
    mismatch = pair_plan['0.32']
    for idx, L in enumerate(cfg.PLATES_0_32):
        items.append({'length': _r(L), 'width': 0.32, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_32_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.88, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_32_strip'})
    
    mismatch = pair_plan['0.72']
    for idx, L in enumerate(cfg.PLATES_0_72):
        items.append({'length': _r(L), 'width': 0.72, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_72_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.48, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_72_strip'})
    
    mismatch = pair_plan['0.70']
    for idx, L in enumerate(cfg.PLATES_0_70):
        items.append({'length': _r(L), 'width': 0.70, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_70_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.50, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_70_strip'})
    
    mismatch = pair_plan['0.86']
    for idx, L in enumerate(cfg.PLATES_0_86):
        items.append({'length': _r(L), 'width': 0.86, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_86_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.34, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_86_strip'})
    
    agg = {}
    for it in items:
        key = (it['length'], it['width'], it['long_cuts'], it['trans_cuts'])
        agg[key] = agg.get(key, 0) + it['qty']
    result = []
    for (L, W, long_cuts, trans_cuts), qty in sorted(
        agg.items(),
        key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0]),
    ):
        result.append({'length': L, 'width': W, 'qty': qty, 'long_cuts': long_cuts, 'trans_cuts': trans_cuts})
    return result


def build_price_rows(price_table: dict, reinforcement_code: int = 8):
    """Формирует строки сметы."""
    items = build_procurement_items()
    rows = []
    total = 0.0
    idx = 1
    for it in items:
        L, W, qty = it['length'], it['width'], it['qty']
        long_cuts, trans_cuts = it['long_cuts'], it['trans_cuts']

        # Определяем код нагрузки:
        #  1) Если нагрузка уже была определена в build_procurement_items - используем её
        #  2) Если при парсинге заказа для (L, W) уже известна нагрузка — берём её;
        #  3) Иначе используем прежнюю логику (6 для узких, reinforcement_code для широких).
        if 'load_code' in it and it['load_code'] is not None:
            load_code = it['load_code']  # Используем нагрузку из items (приоритет!)
        else:
            try:
                load_code = cfg.get_load_code_for_plate(L, W, default=(6 if W < 1.0 else reinforcement_code))
            except Exception:
                load_code = 6 if W < 1.0 else reinforcement_code

        # Формируем имя плиты: берём canonical_name из кэша (если есть), иначе make_plate_name.
        # length_dm_raw прокинут из build_procurement_items, чтобы «57,1» не схлопывалось в «57».
        name = it.get('canonical_name') or cfg.make_plate_name(
            L, W, load_code=load_code, length_dm_raw=it.get('length_dm_raw') or None
        )
        # #region agent log (57/57,1: имя строки сметы)
        if 5.69 <= L <= 5.73:
            try:
                import os as _os
                _log_path = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), 'debug-8e9428.log')
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_price_row", "location": "procurement:build_price_rows", "message": "57/57,1: price_row name", "data": {"L": L, "name": name, "canonical_name": it.get('canonical_name')}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
        # #region agent log (a9176e: 57/57,1 — build_price_rows имя)
        if 5.69 <= L <= 5.73:
            try:
                import os as _os
                _log_path = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), 'debug-a9176e.log')
                import json as _json
                _pay = {"sessionId": "a9176e", "hypothesisId": "H2", "location": "procurement:build_price_rows", "message": "57/57,1 price_row name source", "data": {"L": L, "canonical_name": it.get("canonical_name"), "length_dm_raw": it.get("length_dm_raw"), "name": name}, "timestamp": __import__("time").time() * 1000}
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(_json.dumps(_pay, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        if it.get('warning'):
            name += " (нагрузка?)"
        db_price = get_price(L, load_code, cfg.PRICE_DB_PATH, round_up=True)
        use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
        find_price = find_price_for_plate(price_table, L, load_code) if use_fallback else None
        base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
        # #region agent log
        import json
        import os
        import time
        if base_price_1_2m == 0.0:
            _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-db7a51.log')
            try:
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(json.dumps({"sessionId": "db7a51", "hypothesisId": "build_price_rows", "location": "procurement.py:build_price_rows", "message": "price chain", "data": {"name": name, "L": L, "W": W, "load_code": load_code, "db_price": db_price, "find_price": find_price, "base_price_1_2m": base_price_1_2m}, "timestamp": int(time.time() * 1000)}, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        if base_price_1_2m > 0:
            width_factor = W / 1.2
            base_price = base_price_1_2m * width_factor
        else:
            base_price = 0.0
        
        # Используем единую логику расчета обрезков/отходов
        width_mm = int(round(W * 1000))
        
        # Инициализация переменных
        long_cut_cost = 0.0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        rest_cost = 0.0
        waste_cost = 0.0
        
        current_plan, load_key = _find_plan_for_plate(load_code, L, width_mm, name, 'build_price_rows')
        if not current_plan:
            print(
                f'[DEBUG] build_price_rows: не найден подходящий план для плиты '
                f'{name} при нагрузке {load_key}п. Остатки будут считаться неиспользованными.'
            )

        trim = _calc_trim_components(
            current_plan,
            length=L,
            width_mm=width_mm,
            qty=qty,
            base_price_1_2m=base_price_1_2m,
            base_price=base_price,
            load_code=load_code,
            price_table=price_table,
        )
        rest_cost = trim['rest_cost']
        waste_cost = trim['waste_cost']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE

        if trim['total_cuts_for_this_size'] > 0 and qty > 0:
            long_cut_cost = (trim['total_cuts_for_this_size'] * cfg.LONG_CUT_PRICE_PER_M * L) / qty
        else:
            long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * L)

        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(W - 1.2) < 0.01:
            long_cut_cost = 0.0

        # 🔹 ДОПОЛНИТЕЛЬНО: для диапазона 1020–1080 мм считаем обрезок как отход
        # Таблица завода говорит, что при ширине 1020–1080 мм остаток от реза идёт в утилизацию.
        # Здесь мы всегда добавляем стоимость этого обрезка в цену плиты.
        if 1020 <= width_mm <= 1080 and base_price_1_2m > 0:
            extra_waste_mm = 1200 - width_mm  # например, 120мм для 1080
            if extra_waste_mm > 0:
                waste_cost += (extra_waste_mm / 1200.0) * base_price_1_2m

        unit_price = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost
        weight = cfg.approximate_weight_kg(L, W)
        row_sum = unit_price * qty
        total += row_sum

        metadata = []
        if hasattr(cfg, 'consume_plate_metadata'):
            try:
                metadata = cfg.consume_plate_metadata(L, int(round(W * 1000)), qty)
            except Exception:
                metadata = []
        weeks = [m.get('forming_week') for m in metadata if m.get('forming_week') not in (None, '')]
        week_str = ", ".join(str(w) for w in sorted(set(weeks))) if weeks else ''
        contractors = [m.get('contractor') for m in metadata if m.get('contractor')]
        contractor_str = ", ".join(sorted(set(contractors))) if contractors else ''

        rows.append([
            idx,
            name,
            qty,
            'шт',
            week_str or '—',
            contractor_str or '—',
            f'{weight:.0f}',
            f'{unit_price:,.2f}'.replace(',', ' ').replace('.', ','),
            f'{row_sum:,.2f}'.replace(',', ' ').replace('.', ',')
        ])
        idx += 1
    return rows, total


def build_price_rows_production(price_table: dict, reinforcement_code: int = 8):
    """
    Формирует строки сметы для планирования производства.
    ОТЛИЧИЯ от build_price_rows:
    - Базовая цена берется из таблицы raw_material_costs (стоимость сырья + производство)
    - Добавлен компонент "Переармирование" (перерасход прутьев)
    """
    from core.raw_material_db import get_raw_material_cost
    from core.reinforcement_db import get_reinforcement
    
    items = build_procurement_items()
    rows = []
    total = 0.0
    idx = 1
    
    # ШАГ 1: Определяем максимальное армирование
    # НОВОЕ: Используем PLATE_MAX_REINFORCEMENT_MAP если он заполнен (макс. армирование по дорожке)
    use_track_based_reinforcement = bool(cfg.PLATE_MAX_REINFORCEMENT_MAP)
    
    if use_track_based_reinforcement:
        print(f'[PRODUCTION PRICING] ✅ Используем максимальное армирование по ДОРОЖКАМ')
    else:
        # Fallback: находим максимальное армирование во всем заказе
        global_max_reinforcement = 0.0
        for it in items:
            L, W, qty = it['length'], it['width'], it['qty']
            load_code = it.get('load_code')
            if load_code is None:
                load_code = cfg.get_load_code_for_plate(L, W, default=(6 if W < 1.0 else reinforcement_code))
            
            reinforcement = get_reinforcement(L, load_code, db_path=cfg.PRICE_DB_PATH)
            if reinforcement and reinforcement > global_max_reinforcement:
                global_max_reinforcement = reinforcement
        
        print(f'[PRODUCTION PRICING] ⚠️ PLATE_MAX_REINFORCEMENT_MAP пуст, используем глобальный максимум: {global_max_reinforcement} прутьев')
    
    # ШАГ 2: Рассчитываем стоимость каждой плиты
    for it in items:
        L, W, qty = it['length'], it['width'], it['qty']
        long_cuts, trans_cuts = it['long_cuts'], it['trans_cuts']
        
        # Определяем код нагрузки
        if 'load_code' in it and it['load_code'] is not None:
            load_code = it['load_code']
        else:
            try:
                load_code = cfg.get_load_code_for_plate(L, W, default=(6 if W < 1.0 else reinforcement_code))
            except Exception:
                load_code = 6 if W < 1.0 else reinforcement_code
        
        # Формируем имя плиты
        name = cfg.make_plate_name(L, W, load_code=load_code)
        if it.get('warning'):
            name += " (нагрузка?)"
        
        # ✅ ИЗМЕНЕНИЕ 1: Базовая цена берется из raw_material_costs
        # В БД все плиты имеют формат "ПБ XX-12-НАГРУЗКА" (ширина 1.2м)
        # Поэтому ищем по базовому имени с той же нагрузкой и пересчитываем на фактическую ширину
        base_name_1_2m = cfg.make_plate_name(L, 1.2, load_code=load_code)
        # Убираем "Плиты " и букву "п", т.к. в БД хранится формат "ПБ 23-12-8"
        base_name_1_2m_short = base_name_1_2m.replace('Плиты ', '').replace('п', '')
        base_price_1_2m = get_raw_material_cost(base_name_1_2m_short, db_path=cfg.PRICE_DB_PATH)
        
        if base_price_1_2m is not None:
            # Пересчитываем на фактическую ширину
            width_factor = W / 1.2
            base_price = base_price_1_2m * width_factor
            print(f'[PRODUCTION PRICING] {name}: базовая цена из БД ({base_name_1_2m_short}: {base_price_1_2m:.2f}) × {width_factor:.3f} = {base_price:.2f} руб')
        else:
            # Fallback: если нет в БД, используем старый метод
            db_price = get_price(L, load_code, cfg.PRICE_DB_PATH, round_up=True)
            use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
            find_price = find_price_for_plate(price_table, L, load_code) if use_fallback else None
            base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
            if base_price_1_2m > 0:
                width_factor = W / 1.2
                base_price = base_price_1_2m * width_factor
            else:
                base_price = 0.0
            print(f'[WARNING] Нет данных в raw_material_costs для {base_name_1_2m_short}, использую старый метод: {base_price:.2f}')
        
        # Используем ту же логику расчета резов, остатков и отходов
        from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
        width_mm = int(round(W * 1000))
        
        # Инициализация переменных
        long_cut_cost = 0.0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        rest_cost = 0.0
        waste_cost = 0.0
        rest_used = False  # Флаг: используются ли остатки от резов
        
        # Проверяем план оптимизации (копируем логику из build_price_rows)
        current_plan = None
        if OPT_CASCADING_PLAN_BY_LOAD:
            import math
            from core.optimization import LOAD_TO_REINFORCEMENT_MAP

            load_key = int(math.floor(load_code)) if isinstance(load_code, (int, float)) else 8

            if LOAD_TO_REINFORCEMENT_MAP and load_key in LOAD_TO_REINFORCEMENT_MAP:
                reinforcement_keys = LOAD_TO_REINFORCEMENT_MAP[load_key]

                # Перебираем ВСЕ группы армирования для этой нагрузки
                for reinforcement_key in reinforcement_keys:
                    plan = OPT_CASCADING_PLAN_BY_LOAD.get(reinforcement_key)
                    if not plan:
                        continue

                    orders_req = plan.get('orders_requested') or []
                    # Ищем нашу плиту по длине/ширине/нагрузке
                    for ord_item in orders_req:
                        try:
                            o_len = float(ord_item.get('length', 0))
                            o_width = int(ord_item.get('width', 0))
                            o_load = ord_item.get('load_code', load_key)
                        except Exception:
                            continue

                        if (
                            abs(o_len - L) < 0.05 and
                            o_width == width_mm and
                            int(math.floor(float(o_load))) == load_key
                        ):
                            current_plan = plan
                            print(
                                f'[DEBUG] build_price_rows_production: нашёл план для {name} — '
                                f'нагрузка {load_key}п, армирование {reinforcement_key}'
                            )
                            break

                    if current_plan:
                        break

                if not current_plan:
                    print(
                        f'[DEBUG] build_price_rows_production: не найден подходящий план для плиты '
                        f'{name} при нагрузке {load_key}п. Остатки будут считаться неиспользованными.'
                    )
            else:
                print(
                    f'[DEBUG] build_price_rows_production: нагрузка {load_key}п не найдена в LOAD_TO_REINFORCEMENT_MAP. '
                    f'Остатки будут считаться неиспользованными.'
                )
        
        if current_plan and current_plan.get('primary_cuts'):
            total_cuts_for_this_size = 0
            
            # Первичные резы
            for prim_cut in current_plan['primary_cuts']:
                if prim_cut['width'] == width_mm:
                    prim_lengths = prim_cut.get('lengths', [])
                    if not prim_lengths or any(abs(l - L) < 0.05 for l in prim_lengths):
                        prim_qty = prim_cut.get('qty', 0)
                        total_cuts_for_this_size += prim_qty
                        primary_rest_width_mm = prim_cut['rest']

                        # Остатки
                        unused_rest_total_mm = 0.0
                        if primary_rest_width_mm > 0:
                            produced_rests = prim_qty
                            used_rests = 0

                            if current_plan.get('secondary_cuts'):
                                for sec_cut in current_plan['secondary_cuts']:
                                    if sec_cut.get('source') != primary_rest_width_mm:
                                        continue
                                    src_lengths = sec_cut.get('source_lengths', [])
                                    if src_lengths and not any(abs(sl - L) < 0.05 for sl in src_lengths):
                                        continue
                                    used_rests += sec_cut.get('qty', 0)

                            unused_rests = max(0, produced_rests - used_rests)
                            unused_rest_total_mm = unused_rests * primary_rest_width_mm

                            if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                # ✅ ИСПРАВЛЕНО: Используем base_price_1_2m (цена плиты 1.2м), а не base_price (пересчитанная цена)
                                rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                            elif used_rests > 0:
                                # Остатки были использованы
                                rest_used = True

                        # Вторичные резы и отходы
                        if current_plan.get('secondary_cuts'):
                            for sec_cut in current_plan['secondary_cuts']:
                                sec_lengths = sec_cut.get('lengths', [])
                                if not sec_lengths or any(abs(l - L) < 0.05 for l in sec_lengths):
                                    sec_cuts = sec_cut.get('cuts', [])
                                    if any(abs(width_mm - cut_width) <= 20 for cut_width in sec_cuts):
                                        sec_qty = sec_cut.get('qty', 0)
                                        sec_pieces = sec_cut.get('pieces', 1)

                                        current_cuts = sec_qty * sec_pieces
                                        total_cuts_for_this_size += current_cuts

                                        # Отходы по ширине
                                        waste_w_mm = sec_cut.get('waste', 0)
                                        if waste_w_mm > 0 and base_price > 0 and qty > 0:
                                            cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price
                                            waste_cost += (cost_of_waste_piece * sec_qty) / qty

                                        # Поперечные резы и отходы по длине
                                        src_lens = sec_cut.get('source_lengths', [])
                                        if src_lens:
                                            src_len = src_lens[0]

                                            if sec_cut.get('type') == 'transverse' or abs(src_len - L) > 0.05:
                                                if qty > 0:
                                                    trans_cuts += (1.0 * sec_qty) / qty

                                                len_waste = src_len - L
                                                if len_waste > 0.01:
                                                    src_plate_name = cfg.make_plate_name(src_len, W, load_code=load_code).replace('Плиты ', '').replace('п', '')
                                                    src_price_full = get_raw_material_cost(
                                                        src_plate_name,
                                                        db_path=cfg.PRICE_DB_PATH
                                                    )
                                                    if src_price_full is None:
                                                        # Fallback
                                                        src_price_full_db = get_price(src_len, load_code, cfg.PRICE_DB_PATH)
                                                        if src_price_full_db:
                                                            src_price_full = src_price_full_db * (width_mm / 1200.0)
                                                        else:
                                                            src_price_full = 0.0

                                                    cost_len_waste = src_price_full - base_price

                                                    if cost_len_waste > 0:
                                                        waste_cost += cost_len_waste * (sec_qty / qty)
                        break
            
            # Пересчитываем стоимость продольных резов
            if total_cuts_for_this_size > 0:
                long_cut_cost = (total_cuts_for_this_size * cfg.LONG_CUT_PRICE_PER_M * L) / qty if qty > 0 else 0
            else:
                long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * L)
            
            trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        else:
            # Если нет данных из плана
            long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * L)

        # Жёсткое правило: плиты 1.2м без продольных резов
        if abs(W - 1.2) < 0.01:
            long_cut_cost = 0.0

        # Для диапазона 1020–1080 мм считаем обрезок как отход
        if 1020 <= width_mm <= 1080 and base_price > 0:
            extra_waste_mm = 1200 - width_mm
            if extra_waste_mm > 0:
                waste_cost += (extra_waste_mm / 1200.0) * base_price

        # ✅ ИЗМЕНЕНИЕ 2: Добавляем компонент "Переармирование"
        # НОВОЕ: Используем максимальное армирование ДОРОЖКИ для этой плиты
        rearm_cost = 0.0
        reinforcement = get_reinforcement(L, load_code, db_path=cfg.PRICE_DB_PATH)
        
        # Получаем max_reinforcement для этой конкретной плиты
        if use_track_based_reinforcement:
            plate_key = (round(L, 3), width_mm)
            max_reinforcement = cfg.PLATE_MAX_REINFORCEMENT_MAP.get(plate_key, 0)
            if max_reinforcement == 0:
                for (l, w), mr in cfg.PLATE_MAX_REINFORCEMENT_MAP.items():
                    if abs(l - L) < 0.05 and w == width_mm:
                        max_reinforcement = mr
                        break
        else:
            max_reinforcement = global_max_reinforcement
        
        if reinforcement and max_reinforcement > reinforcement:
            rearm_diff = max_reinforcement - reinforcement
            
            # Если плита целая (без резов) ИЛИ остаток не используется
            if long_cuts == 0 or not rest_used:
                # Базовая формула для ЦЕЛОЙ плиты 1.2м
                rearm_cost = rearm_diff * L * 0.170 * 80
                print(f'[PRODUCTION PRICING] {name}: переармирование (целая/остаток не использован) = ({max_reinforcement:.1f} - {reinforcement:.1f}) × {L} × 0.170 × 80 = {rearm_cost:.2f} руб')
            else:
                # Если остаток используется - формула с учётом фактической ширины плиты
                rearm_cost = rearm_diff * L * 0.170 * 80 * (width_mm / 1200.0)
                print(f'[PRODUCTION PRICING] {name}: переармирование (остаток использован) = ({max_reinforcement:.1f} - {reinforcement:.1f}) × {L} × 0.170 × 80 × ({width_mm} / 1200) = {rearm_cost:.2f} руб')

        unit_price = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost + rearm_cost
        weight = cfg.approximate_weight_kg(L, W)
        row_sum = unit_price * qty
        total += row_sum

        metadata = []
        if hasattr(cfg, 'consume_plate_metadata'):
            try:
                metadata = cfg.consume_plate_metadata(L, int(round(W * 1000)), qty)
            except Exception:
                metadata = []
        weeks = [m.get('forming_week') for m in metadata if m.get('forming_week') not in (None, '')]
        week_str = ", ".join(str(w) for w in sorted(set(weeks))) if weeks else ''
        contractors = [m.get('contractor') for m in metadata if m.get('contractor')]
        contractor_str = ", ".join(sorted(set(contractors))) if contractors else ''

        rows.append([
            idx,
            name,
            qty,
            'шт',
            week_str or '—',
            contractor_str or '—',
            f'{weight:.0f}',
            f'{unit_price:,.2f}'.replace(',', ' ').replace('.', ','),
            f'{row_sum:,.2f}'.replace(',', ' ').replace('.', ',')
        ])
        idx += 1
    return rows, total


def build_component_breakdown(price_table: dict, price_rows: list = None, reinforcement_code: int = 8):
    """Формирует детальную разбивку компонентов для каждого наименования."""
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    
    # Получаем заказы
    plan_orders = get_orders_from_opt_plan()
    if not plan_orders:
        # Fallback на старый способ
        all_orders = []
        for width_mm, plates_list in [
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34), (1080, cfg.PLATES_1_08)
        ]:
            if plates_list:
                length_counts = Counter(plates_list)
                for length, qty in length_counts.items():
                    all_orders.append({
                        'length': length,
                        'width': width_mm,
                        'qty': qty
                    })
        if not all_orders:
            # Если нет заказов из cfg, используем данные из price_rows
            if price_rows:
                print('[DEBUG] build_component_breakdown: используем данные из price_rows')
                for row in price_rows:
                    # row: [idx, name, qty, 'шт', week, contractor, weight, price, sum]
                    name = row[1] if len(row) > 1 else ''
                    try:
                        qty = int(str(row[2]).replace(' ', '').replace(',', '')) if len(row) > 2 else 1
                    except (ValueError, TypeError):
                        qty = 1
                    # Парсим имя для получения длины и ширины
                    # ФОРМАТ ИМЕНИ: ПБ {длина_дм}-{ширина_дм}-{нагрузка}п
                    # Примеры:
                    #   ПБ 28-7,2-8п → длина 28дм=2.8м, ширина 7,2дм=0.72м
                    #   ПБ 59,8-12-8п → длина 59,8дм=5.98м, ширина 12дм=1.2м
                    #   ПБ 73-12-8п → длина 73дм=7.3м, ширина 12дм=1.2м
                    match = re.search(r'ПБ\s+([\d,]+)-([\d,]+)-', name)
                    if match:
                        length = cfg.length_dm_to_m(match.group(1))  # Целое = номинал в дм (длина = номинал/10 м); с запятой/точкой = точное значение в дм
                        
                        width_str = match.group(2).replace(',', '.').replace(' ', '')
                        width_dm = float(width_str)  # ← ШИРИНА ТОЖЕ В ДЕЦИМЕТРАХ!
                        
                        # Преобразуем дециметры в миллиметры и метры
                        width_mm = int(round(width_dm * 100))  # дм × 100 = мм (7.2 → 720мм)
                        width_m = width_mm / 1000.0  # мм → м (720 → 0.72м)
                        
                        all_orders.append({
                            'length': length,
                            'width': width_mm,
                            'qty': qty
                        })
                        print(f'[DEBUG] Распарсили {name}: длина={length}м, ширина={width_m}м ({width_mm}мм)')
            if not all_orders:
                print('[DEBUG] build_component_breakdown: нет заказов, возвращаем пустой список')
                return []
        plan_orders = all_orders
    
    print(f'[DEBUG] build_component_breakdown: найдено заказов: {len(plan_orders)}')
    
    # Группируем заказы по (длина, ширина, НАГРУЗКА, номинал_длины)
    order_counter = Counter()
    for order in plan_orders:
        length = round(float(order.get('length', 0)), 3)
        width_val = order.get('width', 0)
        width_mm = width_val if width_val > 5 else round(width_val * 1000)  # round для корректного округления
        width_m = width_mm / 1000.0
        order_ldr = (order.get('length_dm_raw') or '').strip()

        # Получаем нагрузку для этой плиты
        found_details = []
        if cfg.PLATE_LOAD_DETAILS:
            for key, q in cfg.PLATE_LOAD_DETAILS.items():
                L, W, load = key[0], key[1], key[2]
                key_ldr = key[3] if len(key) > 3 else ''
                if abs(L - length) < 0.05 and abs(W - width_m) < 0.01:
                    found_details.append((load, q, key_ldr))

        order_qty = order.get('qty', 1)

        if found_details and sum(q for _, q, _ in found_details) == order_qty:
            for load_code, q, detail_ldr in found_details:
                order_counter[(length, width_mm, load_code, detail_ldr, False)] += q
        else:
            load_code = cfg.get_load_code_for_plate(length, width_m, default=(6 if width_m < 1.0 else reinforcement_code))
            warning_flag = True if found_details else False
            order_counter[(length, width_mm, load_code, order_ldr, warning_flag)] += order_qty

    breakdown_tables = []

    for (length, width_mm, load_code, ldr, warning_flag), qty in sorted(order_counter.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        width_m = width_mm / 1000.0

        # Имя плиты в детальной разбивке тоже должно отражать фактическую нагрузку
        name = cfg.make_plate_name(length, width_m, load_code=load_code, length_dm_raw=ldr or None)
        if warning_flag:
            name += " (нагрузка?)"
        db_price = get_price(length, load_code, cfg.PRICE_DB_PATH, round_up=True)
        use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
        find_price = find_price_for_plate(price_table, length, load_code) if use_fallback else None
        base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
        
        # Базовая цена с учетом ширины
        if base_price_1_2m > 0:
            width_factor = width_m / 1.2
            base_price = base_price_1_2m * width_factor
        else:
            base_price = 0.0
            base_price_1_2m = 0.0

        # Продольный рез: для плит 1.2 м считаем, что продольных резов нет
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
        else:
            # Для плит уже меньше 1.2 м предполагаем один продольный рез
            long_cuts = 1 if width_m < 1.15 else 0
        long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * length)
        
        # Поперечный рез (пока 0, нужно будет добавить логику)
        trans_cuts = 0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        
        # Остаток
        rest_width_mm = 1200 - width_mm if width_m < 1.15 else 0
        rest_cost = 0.0
        rest_used = False
        
        # Отходы
        waste_cost = 0.0
        waste_width_mm = 0
        waste_terms = []  # (ширина_отхода_мм, количество_операций) для наглядной формулы
        
        current_plan, load_key = _find_plan_for_plate(load_code, length, width_mm, name, 'build_component_breakdown')
        if not current_plan:
            print(
                f'[DEBUG] build_component_breakdown: не найден подходящий план по нагрузке {load_key}п '
                f'для плиты {name}. Остатки будут считаться неиспользованными.'
            )

        total_cuts_count = 0  # Общее количество резов для отображения в таблице
        trim = _calc_trim_components(
            current_plan,
            length=length,
            width_mm=width_mm,
            qty=qty,
            base_price_1_2m=base_price_1_2m,
            base_price=base_price,
            load_code=load_code,
            price_table=price_table,
        )
        rest_cost = trim['rest_cost']
        rest_width_mm = trim['rest_width_mm']
        rest_used = trim['rest_used']
        waste_cost = trim['waste_cost']
        waste_terms = trim['waste_terms']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE

        if trim['total_plates_from_cuts'] > 0 and qty > 0:
            total_cuts_count = trim['total_cuts_for_this_size']
            long_cuts = trim['total_cuts_for_this_size']
            long_cut_cost = (long_cuts * cfg.LONG_CUT_PRICE_PER_M * length) / qty
        elif not (current_plan and current_plan.get('primary_cuts')):
            # ✅ НОВАЯ ЛОГИКА: Если плана оптимизации нет, рассчитываем остатки вручную
            print(f'[DEBUG] Плана оптимизации нет для {name}, используем ручной расчёт остатков')
            
            # Если ширина < 1.15м, значит была продольная резка
            if width_m < 1.15:
                # Рассчитываем остаток
                rest_width_mm = 1200 - width_mm
                
                # ВАЖНО: Остаток считаем НЕИСПОЛЬЗОВАННЫМ (т.к. нет информации об оптимизации)
                if rest_width_mm > 0 and base_price_1_2m > 0:
                    # Стоимость остатка = (ширина_остатка / 1200) × базовая_цена_1.2м
                    rest_cost = (rest_width_mm / 1200.0) * base_price_1_2m
                    print(f'[DEBUG] Остаток {rest_width_mm}мм не использован, добавляем к цене: {rest_cost:.2f} руб')
                else:
                    rest_cost = 0.0

        # 🔹 ДОПОЛНИТЕЛЬНО: диапазон 1020–1080 мм — обрезок по таблице завода
        # Таблица допустимых резов говорит, что при ширине 1020–1080 мм
        # остаток идёт в утилизацию. Добавим его стоимость как отход.
        if 1020 <= width_mm <= 1080 and base_price_1_2m > 0:
            extra_waste_mm = 1200 - width_mm
            if extra_waste_mm > 0:
                # Для формулы: показываем, что было extra_waste_mm × qty мм отхода
                waste_terms.append((extra_waste_mm, qty))
                # Стоимость отхода на одну плиту
                waste_cost += (extra_waste_mm / 1200.0) * base_price_1_2m
                    
        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
            long_cut_cost = 0.0

        # ИТОГО за 1 плиту
        total_per_unit = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost
        
        # ✅ ИСПРАВЛЕНИЕ: Сначала умножаем, ПОТОМ округляем (как в build_price_rows)
        total_for_qty = total_per_unit * qty  # БЕЗ промежуточного округления!
        
        # Округление только для отображения в таблице
        total_rounded = round(total_per_unit, 2)
        
        # Формируем таблицу
        table_rows = []
        
        # Базовая цена
        base_price_1_2m_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
        width_m_str = f"{width_m:.2f}".replace('.', ',')
        calculation_str = f"{base_price_1_2m_str} × ({width_m_str} / 1.2)"
        table_rows.append([
            f"Базовая цена ({width_m_str}м)",
            calculation_str,
            f"{base_price:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # Продольный рез
        if long_cuts > 0:
            # Показываем общее количество резов в расчете
            if total_cuts_count > 0:
                # Если есть информация о количестве резов из плана оптимизации
                if qty > 1:
                    # Показываем: "460 × 3,6 × 4 / 4" (общее количество резов / количество плит)
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {total_cuts_count:.0f} / {qty}".replace('.', ',')
                else:
                    # Для одной плиты просто показываем количество резов
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {total_cuts_count:.0f}".replace('.', ',')
            else:
                # Если нет информации из плана, используем старое отображение
                if abs(long_cuts - 1.0) > 0.001:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {long_cuts:.2f}".replace('.', ',')
                else:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f}".replace('.', ',')
            table_rows.append([
                "Продольный рез",
                long_calc,
                f"{long_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Поперечный рез
        if trans_cuts > 0:
            trans_calc = f"{cfg.TRANSVERSE_CUT_PRICE:.0f} × {trans_cuts}"
            table_rows.append([
                "Поперечный рез",
                trans_calc,
                f"{trans_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Остаток
        if rest_width_mm > 0:
            if rest_cost > 0:
                # Показываем стоимость остатка
                base_price_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
                rest_calc = f"({rest_width_mm} / 1200) × {base_price_str} / {qty}"
                rest_status_str = f"{rest_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            else:
                rest_calc = "0 (использован)" if rest_used else "0 (не использован)"
                rest_status_str = "0,00 руб"
            table_rows.append([
                f"Остаток ({rest_width_mm}мм)",
                rest_calc,
                rest_status_str
            ])
        
        # Отходы
        if waste_cost > 0 or waste_terms:
            if base_price_1_2m > 0:
                base_price_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
                if waste_terms:
                    # Формируем наглядное выражение: "240×2" или "240×2 + 20"
                    parts = []
                    for w_mm, n in waste_terms:
                        w_str = f"{int(w_mm)}"
                        parts.append(f"{w_str}×{n}" if n > 1 else w_str)
                    waste_expr = " + ".join(parts)
                else:
                    waste_expr = "0"
                waste_calc = f"({waste_expr} / 1200) × {base_price_str} / {qty}"
            else:
                waste_expr = "0"
                waste_calc = "0"
            table_rows.append([
                f"Отходы ({waste_expr}мм)",
                waste_calc,
                f"{waste_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # ИТОГО за 1 плиту
        table_rows.append([
            "ИТОГО за 1 плиту",
            "",
            f"{total_per_unit:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # Округлено
        table_rows.append([
            "Округлено",
            "",
            f"{total_rounded:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # За N плит
        table_rows.append([
            f"За {qty} плит",
            f"{total_rounded:,.2f} × {qty}".replace(',', ' ').replace('.', ','),
            f"{total_for_qty:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        breakdown_tables.append({
            'name': name,
            'rows': table_rows
        })
    
    print(f'[DEBUG] build_component_breakdown: создано таблиц разбивки: {len(breakdown_tables)}')
    return breakdown_tables


def build_component_breakdown_production(price_table: dict, price_rows: list = None, reinforcement_code: int = 8, tracks_for_day: list = None):
    """
    Формирует детальную разбивку компонентов для планирования производства.
    ОТЛИЧИЯ от build_component_breakdown:
    - Базовая цена берется из таблицы raw_material_costs
    - Добавлен компонент "Переармирование"
    
    Args:
        tracks_for_day: Список дорожек текущего дня. Если указан, будут включены только плиты из этих дорожек.
    """
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    from core.raw_material_db import get_raw_material_cost
    from core.reinforcement_db import get_reinforcement
    
    # ✅ НОВОЕ: Если указаны дорожки текущего дня, собираем плиты только из них
    if tracks_for_day:
        print(f'[PRODUCTION BREAKDOWN] ✅ Фильтруем плиты по дорожкам текущего дня ({len(tracks_for_day)} дорожек)')
        plan_orders = []
        plates_from_tracks = Counter()
        
        # Собираем все плиты из дорожек текущего дня
        for track in tracks_for_day:
            for item in track.get('items', []):
                length = item.get('length', 0)
                if not length:
                    continue
                
                # Определяем ширину плиты
                if item.get('mode') == 'solid':
                    width_mm = 1200
                elif item.get('mode') == 'split':
                    width_mm = int(round(item.get('main_w', 1.2) * 1000))
                elif item.get('mode') == 'transverse':
                    width_mm = int(round(item.get('width', 1.2) * 1000))
                else:
                    width_mm = 1200
                
                # Считаем количество плит этого типа
                plates_from_tracks[(round(length, 3), width_mm)] += 1
        
        # Формируем список заказов из плит дорожек
        for (length, width_mm), qty in plates_from_tracks.items():
            plan_orders.append({
                'length': length,
                'width': width_mm,
                'qty': qty
            })
        
        print(f'[PRODUCTION BREAKDOWN] Собрано {len(plan_orders)} типов плит из дорожек текущего дня')
    else:
        # Получаем заказы из плана (старая логика - все плиты)
        plan_orders = get_orders_from_opt_plan()
        if not plan_orders:
            # Fallback на старый способ
            all_orders = []
            for width_mm, plates_list in [
                (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
                (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
                (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
                (340, cfg.PLATES_0_34), (1080, cfg.PLATES_1_08)
            ]:
                if plates_list:
                    length_counts = Counter(plates_list)
                    for length, qty in length_counts.items():
                        all_orders.append({
                            'length': length,
                            'width': width_mm,
                            'qty': qty
                        })
            if not all_orders:
                # Если нет заказов из cfg, используем данные из price_rows
                if price_rows:
                    print('[DEBUG] build_component_breakdown_production: используем данные из price_rows')
                    for row in price_rows:
                        name = row[1] if len(row) > 1 else ''
                        try:
                            qty = int(str(row[2]).replace(' ', '').replace(',', '')) if len(row) > 2 else 1
                        except (ValueError, TypeError):
                            qty = 1
                        # Парсим имя (поддержка дробных дециметров: "59,8" или "60")
                        match = re.search(r'ПБ\s+([\d,]+)-([\d,]+)-', name)
                        if match:
                            length = cfg.length_dm_to_m(match.group(1))
                            width_str = match.group(2).replace(',', '.').replace(' ', '')
                            width_dm = float(width_str)
                            width_mm = int(round(width_dm * 100))
                            width_m = width_mm / 1000.0
                            
                            all_orders.append({
                                'length': length,
                                'width': width_mm,
                                'qty': qty
                            })
                            print(f'[DEBUG] Распарсили {name}: длина={length}м, ширина={width_m}м ({width_mm}мм)')
                if not all_orders:
                    print('[DEBUG] build_component_breakdown_production: нет заказов, возвращаем пустой список')
                    return []
            plan_orders = all_orders
    
    print(f'[DEBUG] build_component_breakdown_production: найдено заказов: {len(plan_orders)}')
    
    # Группируем заказы по (длина, ширина, НАГРУЗКА, номинал_длины)
    order_counter = Counter()
    for order in plan_orders:
        length = round(float(order.get('length', 0)), 3)
        width_val = order.get('width', 0)
        width_mm = width_val if width_val > 5 else round(width_val * 1000)  # round для корректного округления
        width_m = width_mm / 1000.0
        order_ldr = (order.get('length_dm_raw') or '').strip()

        # Получаем нагрузку для этой плиты
        found_details = []
        if cfg.PLATE_LOAD_DETAILS:
            for key, q in cfg.PLATE_LOAD_DETAILS.items():
                L, W, load = key[0], key[1], key[2]
                key_ldr = key[3] if len(key) > 3 else ''
                if abs(L - length) < 0.05 and abs(W - width_m) < 0.01:
                    found_details.append((load, q, key_ldr))

        order_qty = order.get('qty', 1)

        if found_details and sum(q for _, q, _ in found_details) == order_qty:
            for load_code, q, detail_ldr in found_details:
                order_counter[(length, width_mm, load_code, detail_ldr, False)] += q
        else:
            load_code = cfg.get_load_code_for_plate(length, width_m, default=(6 if width_m < 1.0 else reinforcement_code))
            warning_flag = True if found_details else False
            order_counter[(length, width_mm, load_code, order_ldr, warning_flag)] += order_qty

    # ШАГ 1: Определяем максимальное армирование
    # НОВОЕ: Используем PLATE_MAX_REINFORCEMENT_MAP если он заполнен (макс. армирование по дорожке)
    # Иначе - fallback на максимум по всему заказу
    use_track_based_reinforcement = bool(cfg.PLATE_MAX_REINFORCEMENT_MAP)

    if use_track_based_reinforcement:
        print(f'[PRODUCTION BREAKDOWN] ✅ Используем максимальное армирование по ДОРОЖКАМ (из PLATE_MAX_REINFORCEMENT_MAP)')
    else:
        # Fallback: находим максимальное армирование во всем заказе
        global_max_reinforcement = 0.0
        for (length, width_mm, load_code, ldr, warning_flag), qty in order_counter.items():
            reinforcement = get_reinforcement(length, load_code, db_path=cfg.PRICE_DB_PATH)
            if reinforcement and reinforcement > global_max_reinforcement:
                global_max_reinforcement = reinforcement
        print(f'[PRODUCTION BREAKDOWN] ⚠️ PLATE_MAX_REINFORCEMENT_MAP пуст, используем глобальный максимум: {global_max_reinforcement} прутьев')

    breakdown_tables = []

    for (length, width_mm, load_code, ldr, warning_flag), qty in sorted(order_counter.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        width_m = width_mm / 1000.0

        # Имя плиты
        name = cfg.make_plate_name(length, width_m, load_code=load_code, length_dm_raw=ldr or None)
        if warning_flag:
            name += " (нагрузка?)"
        
        # ✅ ИЗМЕНЕНИЕ 1: Базовая цена из raw_material_costs
        # В БД все плиты имеют формат "ПБ XX-12-НАГРУЗКА" (ширина 1.2м)
        # Поэтому ищем по базовому имени с той же нагрузкой и пересчитываем на фактическую ширину
        base_name_1_2m = cfg.make_plate_name(length, 1.2, load_code=load_code)
        # Убираем "Плиты " и букву "п", т.к. в БД хранится формат "ПБ 23-12-8"
        base_name_1_2m_short = base_name_1_2m.replace('Плиты ', '').replace('п', '')
        base_price_1_2m = get_raw_material_cost(base_name_1_2m_short, db_path=cfg.PRICE_DB_PATH)
        
        if base_price_1_2m is not None:
            # Пересчитываем на фактическую ширину
            width_factor = width_m / 1.2
            base_price = base_price_1_2m * width_factor
            print(f'[PRODUCTION BREAKDOWN] {name}: базовая цена из БД ({base_name_1_2m_short}: {base_price_1_2m:.2f}) × {width_factor:.3f} = {base_price:.2f} руб')
        else:
            # Fallback
            db_price = get_price(length, load_code, cfg.PRICE_DB_PATH, round_up=True)
            use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
            find_price = find_price_for_plate(price_table, length, load_code) if use_fallback else None
            base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
            if base_price_1_2m > 0:
                width_factor = width_m / 1.2
                base_price = base_price_1_2m * width_factor
            else:
                base_price = 0.0
                base_price_1_2m = 0.0
            print(f'[WARNING] Нет данных в raw_material_costs для {base_name_1_2m_short}, использую старый метод: {base_price:.2f}')

        # Продольный рез
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
        else:
            long_cuts = 1 if width_m < 1.15 else 0
        long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * length)
        
        # Поперечный рез
        trans_cuts = 0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        
        # Остаток
        rest_width_mm = 1200 - width_mm if width_m < 1.15 else 0
        rest_cost = 0.0
        rest_used = False
        
        # Отходы
        waste_cost = 0.0
        waste_width_mm = 0
        waste_terms = []
        
        # Проверяем план оптимизации (копируем логику из build_component_breakdown)
        current_plan = None
        if OPT_CASCADING_PLAN_BY_LOAD:
            import math
            from core.optimization import LOAD_TO_REINFORCEMENT_MAP

            load_key = int(math.floor(load_code)) if isinstance(load_code, (int, float)) else 8

            print(f'[DEBUG] build_component_breakdown_production: ищем план для нагрузки {load_key}п (плита {name})')

            if LOAD_TO_REINFORCEMENT_MAP and load_key in LOAD_TO_REINFORCEMENT_MAP:
                reinforcement_keys = LOAD_TO_REINFORCEMENT_MAP[load_key]

                for reinforcement_key in reinforcement_keys:
                    plan = OPT_CASCADING_PLAN_BY_LOAD.get(reinforcement_key)
                    if not plan:
                        continue

                    orders_req = plan.get('orders_requested') or []
                    for ord_item in orders_req:
                        try:
                            o_len = float(ord_item.get('length', 0))
                            o_width = int(ord_item.get('width', 0))
                            o_load = ord_item.get('load_code', load_key)
                        except Exception:
                            continue

                        if (
                            abs(o_len - length) < 0.05 and
                            o_width == width_mm and
                            int(math.floor(float(o_load))) == load_key
                        ):
                            current_plan = plan
                            print(
                                f'[DEBUG] ✅ Используем план для нагрузки {load_key}п '
                                f'(армирование {reinforcement_key}) для плиты {name}'
                            )
                            break

                    if current_plan:
                        break

                if not current_plan:
                    print(
                        f'[DEBUG] ⚠️ Не найден подходящий план по нагрузке {load_key}п '
                        f'для плиты {name}. Остатки будут считаться неиспользованными.'
                    )
            else:
                print(
                    f'[DEBUG] ⚠️ Нагрузка {load_key}п не найдена в LOAD_TO_REINFORCEMENT_MAP. '
                    f'Остатки для плиты {name} будут считаться неиспользованными.'
                )
        else:
            print(f'[DEBUG] ⚠️ OPT_CASCADING_PLAN_BY_LOAD пустой! Нет планов оптимизации по нагрузкам.')
        
        total_cuts_count = 0
        if current_plan and current_plan.get('primary_cuts'):
            total_cuts_for_this_size = 0
            total_plates_from_cuts = 0
            primary_rest_width_mm = 0
            
            # Первичные резы
            for prim_cut in current_plan['primary_cuts']:
                if prim_cut['width'] == width_mm:
                    prim_lengths = prim_cut.get('lengths', [])
                    if not prim_lengths or any(abs(l - length) < 0.05 for l in prim_lengths):
                        prim_qty = prim_cut.get('qty', 0)
                        total_cuts_for_this_size += prim_qty
                        total_plates_from_cuts += prim_qty
                        primary_rest_width_mm = prim_cut['rest']

                        # Остатки
                        unused_rest_total_mm = 0.0
                        if primary_rest_width_mm > 0:
                            produced_rests = prim_qty
                            used_rests = 0

                            if current_plan.get('secondary_cuts'):
                                for sec_cut in current_plan['secondary_cuts']:
                                    if sec_cut.get('source') != primary_rest_width_mm:
                                        continue
                                    src_lengths = sec_cut.get('source_lengths', [])
                                    if src_lengths and not any(abs(sl - length) < 0.05 for sl in src_lengths):
                                        continue
                                    used_rests += sec_cut.get('qty', 0)

                            unused_rests = max(0, produced_rests - used_rests)
                            unused_rest_total_mm = unused_rests * primary_rest_width_mm

                            if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                # ✅ ИСПРАВЛЕНО: Используем base_price_1_2m (цена плиты 1.2м), а не base_price (пересчитанная цена)
                                # Остаток считается от полной ширины 1200мм, поэтому нужна цена плиты 1.2м
                                rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                                rest_width_mm = unused_rest_total_mm
                            elif produced_rests > 0:
                                rest_used = True
                                rest_width_mm = 0

                        # Вторичные резы и отходы
                        if current_plan.get('secondary_cuts'):
                            for sec_cut in current_plan['secondary_cuts']:
                                sec_lengths = sec_cut.get('lengths', [])
                                if not sec_lengths or any(abs(l - length) < 0.05 for l in sec_lengths):
                                    sec_cuts = sec_cut.get('cuts', [])
                                    
                                    if any(abs(width_mm - cut_width) <= 20 for cut_width in sec_cuts):
                                        sec_qty = sec_cut.get('qty', 0)
                                        sec_pieces = sec_cut.get('pieces', 1)
                                        
                                        current_cuts = sec_qty * sec_pieces
                                        total_cuts_for_this_size += current_cuts
                                        total_plates_from_cuts += current_cuts
                                        
                                        # Отходы по ширине
                                        waste_w_mm = sec_cut.get('waste', 0)
                                        if waste_w_mm > 0 and base_price > 0 and qty > 0:
                                            cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price
                                            waste_cost += (cost_of_waste_piece * sec_qty) / qty
                                            waste_terms.append((waste_w_mm, sec_qty))

                                        # Поперечные резы
                                        src_lens = sec_cut.get('source_lengths', [])
                                        if src_lens:
                                            src_len = src_lens[0]

                                            if sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05:
                                                if qty > 0:
                                                    trans_cuts += (1.0 * sec_qty) / qty

                                                len_waste = src_len - length
                                                if len_waste > 0.01:
                                                    src_plate_name = cfg.make_plate_name(src_len, width_m, load_code=load_code).replace('Плиты ', '').replace('п', '')
                                                    src_price_full = get_raw_material_cost(
                                                        src_plate_name,
                                                        db_path=cfg.PRICE_DB_PATH
                                                    )
                                                    if src_price_full is None:
                                                        src_price_full_db = get_price(src_len, load_code, cfg.PRICE_DB_PATH)
                                                        if src_price_full_db:
                                                            src_price_full = src_price_full_db * (width_mm / 1200.0)
                                                        else:
                                                            src_price_full = 0.0

                                                    cost_len_waste = src_price_full - base_price

                                                    if cost_len_waste > 0:
                                                        waste_cost += cost_len_waste * (sec_qty / qty)

                            trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
                        break
            
            # Если нашли резы в плане
            if total_plates_from_cuts > 0:
                total_cuts_count = total_cuts_for_this_size
                long_cuts = total_cuts_for_this_size
                long_cut_cost = (long_cuts * cfg.LONG_CUT_PRICE_PER_M * length) / qty if qty > 0 else 0
        else:
            # Если плана нет, рассчитываем остатки вручную
            print(f'[DEBUG] Плана оптимизации нет для {name}, используем ручной расчёт остатков')
            
            if width_m < 1.15:
                rest_width_mm = 1200 - width_mm
                
                if rest_width_mm > 0 and base_price_1_2m > 0:
                    # ✅ ИСПРАВЛЕНО: Используем base_price_1_2m (цена плиты 1.2м), а не base_price (пересчитанная цена)
                    # Остаток считается от полной ширины 1200мм, поэтому нужна цена плиты 1.2м
                    rest_cost = (rest_width_mm / 1200.0) * base_price_1_2m
                    print(f'[DEBUG] Остаток {rest_width_mm}мм не использован, добавляем к цене: {rest_cost:.2f} руб')
                else:
                    rest_cost = 0.0

        # Для диапазона 1020–1080 мм
        if 1020 <= width_mm <= 1080 and base_price > 0:
            extra_waste_mm = 1200 - width_mm
            if extra_waste_mm > 0:
                waste_terms.append((extra_waste_mm, qty))
                waste_cost += (extra_waste_mm / 1200.0) * base_price
                    
        # Жёсткое правило: плиты 1.2м без продольных резов
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
            long_cut_cost = 0.0

        # ✅ ИЗМЕНЕНИЕ 2: Переармирование
        # НОВОЕ: Используем максимальное армирование ДОРОЖКИ для этой плиты
        rearm_cost = 0.0
        reinforcement = get_reinforcement(length, load_code, db_path=cfg.PRICE_DB_PATH)
        
        # Получаем max_reinforcement для этой конкретной плиты
        if use_track_based_reinforcement:
            # Ищем в карте по (length, width_mm)
            plate_key = (round(length, 3), width_mm)
            max_reinforcement = cfg.PLATE_MAX_REINFORCEMENT_MAP.get(plate_key, 0)
            if max_reinforcement == 0:
                # Попробуем с другими округлениями
                for (l, w), mr in cfg.PLATE_MAX_REINFORCEMENT_MAP.items():
                    if abs(l - length) < 0.05 and w == width_mm:
                        max_reinforcement = mr
                        break
            if max_reinforcement > 0:
                print(f'[PRODUCTION BREAKDOWN] {name}: макс. армирование дорожки = {max_reinforcement:.1f}')
        else:
            max_reinforcement = global_max_reinforcement
        
        if reinforcement and max_reinforcement > reinforcement:
            rearm_diff = max_reinforcement - reinforcement
            
            # Если плита целая (без резов) ИЛИ остаток не используется
            if long_cuts == 0 or not rest_used:
                # Базовая формула для ЦЕЛОЙ плиты 1.2м
                rearm_cost = rearm_diff * length * 0.170 * 80
                print(f'[PRODUCTION BREAKDOWN] {name}: переармирование (целая/остаток не использован) = ({max_reinforcement:.1f} - {reinforcement:.1f}) × {length} × 0.170 × 80 = {rearm_cost:.2f} руб')
            else:
                # Если остаток используется - формула с учётом фактической ширины плиты
                rearm_cost = rearm_diff * length * 0.170 * 80 * (width_mm / 1200.0)
                print(f'[PRODUCTION BREAKDOWN] {name}: переармирование (остаток использован) = ({max_reinforcement:.1f} - {reinforcement:.1f}) × {length} × 0.170 × 80 × ({width_mm} / 1200) = {rearm_cost:.2f} руб')

        # ИТОГО
        total_per_unit = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost + rearm_cost
        total_for_qty = total_per_unit * qty
        total_rounded = round(total_per_unit, 2)
        
        # Формируем таблицу
        table_rows = []
        
        # Базовая цена
        # Базовая цена
        if base_price_1_2m > 0 and abs(width_m - 1.2) > 0.01:
            # Показываем пересчет для узких плит
            base_price_1_2m_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
            width_m_str = f"{width_m:.2f}".replace('.', ',')
            base_calc = f"{base_price_1_2m_str} × ({width_m_str} / 1,2)"
        else:
            # Целая плита 1.2м
            base_calc = "из БД raw_material_costs"
        
        table_rows.append([
            f"Базовая цена (сырьё + произв.)",
            base_calc,
            f"{base_price:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # Продольный рез
        if long_cuts > 0:
            if total_cuts_count > 0:
                if qty > 1:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {total_cuts_count:.0f} / {qty}".replace('.', ',')
                else:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {total_cuts_count:.0f}".replace('.', ',')
            else:
                if abs(long_cuts - 1.0) > 0.001:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {long_cuts:.2f}".replace('.', ',')
                else:
                    long_calc = f"{cfg.LONG_CUT_PRICE_PER_M:.0f} × {length:.1f}".replace('.', ',')
            table_rows.append([
                "Продольный рез",
                long_calc,
                f"{long_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Поперечный рез
        if trans_cuts > 0:
            trans_calc = f"{cfg.TRANSVERSE_CUT_PRICE:.0f} × {trans_cuts}"
            table_rows.append([
                "Поперечный рез",
                trans_calc,
                f"{trans_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Остаток
        if rest_width_mm > 0:
            if rest_cost > 0:
                # ✅ ИСПРАВЛЕНО: Показываем base_price_1_2m (цена плиты 1.2м) в формуле
                base_price_1_2m_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
                rest_calc = f"({rest_width_mm} / 1200) × {base_price_1_2m_str} / {qty}"
                rest_status_str = f"{rest_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            else:
                rest_calc = "0 (использован)" if rest_used else "0 (не использован)"
                rest_status_str = "0,00 руб"
            table_rows.append([
                f"Остаток ({rest_width_mm}мм)",
                rest_calc,
                rest_status_str
            ])
        
        # Отходы
        if waste_cost > 0 or waste_terms:
            if base_price > 0:
                base_price_str = f"{base_price:,.2f}".replace(',', ' ').replace('.', ',')
                if waste_terms:
                    parts = []
                    for w_mm, n in waste_terms:
                        w_str = f"{int(w_mm)}"
                        parts.append(f"{w_str}×{n}" if n > 1 else w_str)
                    waste_expr = " + ".join(parts)
                else:
                    waste_expr = "0"
                waste_calc = f"({waste_expr} / 1200) × {base_price_str} / {qty}"
            else:
                waste_expr = "0"
                waste_calc = "0"
            table_rows.append([
                f"Отходы ({waste_expr}мм)",
                waste_calc,
                f"{waste_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # ✅ НОВАЯ СТРОКА: Переармирование
        if rearm_cost > 0:
            # Формула зависит от того, используется ли остаток
            if long_cuts == 0 or not rest_used:
                # Целая плита или остаток не используется
                rearm_calc = f"({max_reinforcement:.1f} - {reinforcement:.1f}) × {length:.1f} × 0,170 × 80".replace('.', ',')
            else:
                # Остаток используется - с учётом ширины
                width_m_display = width_mm / 1000.0
                rearm_calc = f"({max_reinforcement:.1f} - {reinforcement:.1f}) × {length:.1f} × 0,170 × 80 × ({width_m_display:.2f} / 1,2)".replace('.', ',')
            
            table_rows.append([
                "Переармирование",
                rearm_calc,
                f"{rearm_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # ИТОГО за 1 плиту
        table_rows.append([
            "ИТОГО за 1 плиту",
            "",
            f"{total_per_unit:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # Округлено
        table_rows.append([
            "Округлено",
            "",
            f"{total_rounded:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        # За N плит
        table_rows.append([
            f"За {qty} плит",
            f"{total_rounded:,.2f} × {qty}".replace(',', ' ').replace('.', ','),
            f"{total_for_qty:,.2f} руб".replace(',', ' ').replace('.', ',')
        ])
        
        breakdown_tables.append({
            'name': name,
            'rows': table_rows
        })
    
    print(f'[DEBUG] build_component_breakdown_production: создано таблиц разбивки: {len(breakdown_tables)}')
    return breakdown_tables

