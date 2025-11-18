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


def get_orders_from_opt_plan():
    """
    Возвращает исходный заказ (length/width/qty), сохранённый оптимизатором.
    Если заказа нет (например, визуализацию запустили без бота), возвращает None.
    """
    try:
        from core.optimization import OPT_CASCADING_PLAN
    except ImportError:
        return None

    plan = OPT_CASCADING_PLAN
    if plan and plan.get('orders_requested'):
        # Заказ хранится в формате [{'length': float, 'width': мм, 'qty': int}, ...]
        orders_copy = []
        for order in plan['orders_requested']:
            try:
                orders_copy.append({
                    'length': float(order.get('length', 0)),
                    'width': order.get('width', 0),
                    'qty': int(order.get('qty', 1))
                })
            except Exception:
                continue
        return orders_copy
    return None


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
            order_counter[(length, width_m)] += order['qty']
        for (length, width_m), qty in sorted(order_counter.items(), key=lambda x: (x[0][1], x[0][0])):
            # Жёсткое правило: плиты 1.2 м считаем целыми, без продольных резов
            if abs(width_m - 1.2) < 0.01:
                long_cuts = 0
            else:
                long_cuts = 1 if width_m < 1.15 else 0

            items.append({
                'length': length,
                'width': width_m,
                'qty': qty,
                'long_cuts': long_cuts,
                'trans_cuts': 0
            })
        return items

    # Приоритет 2: Используем реальный заказ из cfg.PLATES_* (legacy режим)
    # Это то, что пользователь заказал вручную, если бот не запускался.
    all_plates = []
    for width_mm, plates_list in [
        (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
        (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
        (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
        (340, cfg.PLATES_0_34), (1080, cfg.PLATES_1_08), (1200, cfg.PLATES_1_2),
        (1000, cfg.PLATES_1_0)
    ]:
        if plates_list:
            length_counts = Counter(plates_list)
            for length, qty in length_counts.items():
                all_plates.append({
                    'length': length,
                    'width': width_mm / 1000.0,  # в метрах
                    'qty': qty
                })
    
    if all_plates:
        # Есть реальный заказ - используем его
        for plate in all_plates:
            # Определяем количество резов (примерная оценка)
            width_m = plate['width']

            # Жёсткое правило: плиты 1.2 м считаем целыми, без продольных резов
            if abs(width_m - 1.2) < 0.01:
                long_cuts = 0
            else:
                # Продольные резы: если ширина < 1.2м, значит был рез
                long_cuts = 1 if width_m < 1.15 else 0

            # Поперечные резы: пока 0, они учтены в оптимизации
            trans_cuts = 0

            items.append({
                'length': plate['length'],
                'width': width_m,
                'qty': plate['qty'],
                'long_cuts': long_cuts,
                'trans_cuts': trans_cuts
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
        for (L, W, long_cuts, trans_cuts), qty in sorted(agg.items(), key=lambda x: (x[0][1], x[0][0])):
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
    
    for L in cfg.PLATES_1_2:
        items.append({'length': round(L, 1), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'as_is'})
    for L in cfg.PLATES_1_5_TO_1_2:
        items.append({'length': round(L, 1), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'to_1_2_main'})
        items.append({'length': round(L, 1), 'width': 0.3, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_2_strip'})
    for L in cfg.PLATES_1_0:
        items.append({'length': round(L, 1), 'width': 1.0, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_main'})
        items.append({'length': round(L, 1), 'width': 0.2, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_strip'})
    for L in cfg.PLATES_1_08:
        items.append({'length': round(L, 1), 'width': 1.08, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_main'})
        items.append({'length': round(L, 1), 'width': 0.12, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_strip'})
    for L in cfg.PLATES_0_46:
        items.append({'length': round(L, 1), 'width': 0.46, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_main'})
        items.append({'length': round(L, 1), 'width': 0.74, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_strip'})
    
    mismatch = pair_plan['0.32']
    for idx, L in enumerate(cfg.PLATES_0_32):
        items.append({'length': round(L, 1), 'width': 0.32, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_32_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': round(L, 1), 'width': 0.88, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_32_strip'})
    
    mismatch = pair_plan['0.72']
    for idx, L in enumerate(cfg.PLATES_0_72):
        items.append({'length': round(L, 1), 'width': 0.72, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_72_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': round(L, 1), 'width': 0.48, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_72_strip'})
    
    mismatch = pair_plan['0.70']
    for idx, L in enumerate(cfg.PLATES_0_70):
        items.append({'length': round(L, 1), 'width': 0.70, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_70_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': round(L, 1), 'width': 0.50, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_70_strip'})
    
    mismatch = pair_plan['0.86']
    for idx, L in enumerate(cfg.PLATES_0_86):
        items.append({'length': round(L, 1), 'width': 0.86, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_86_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': round(L, 1), 'width': 0.34, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_86_strip'})
    
    agg = {}
    for it in items:
        key = (it['length'], it['width'], it['long_cuts'], it['trans_cuts'])
        agg[key] = agg.get(key, 0) + it['qty']
    result = []
    for (L, W, long_cuts, trans_cuts), qty in sorted(agg.items(), key=lambda x: (x[0][1], x[0][0])):
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
        name = cfg.make_plate_name(L, W)
        
        load_code = 6 if W < 1.0 else reinforcement_code
        db_price = get_price(L, load_code, cfg.PRICE_DB_PATH)
        base_price_1_2m = db_price if db_price is not None else (find_price_for_plate(price_table, L, load_code) or 0.0)
        
        if base_price_1_2m > 0:
            width_factor = W / 1.2
            base_price = base_price_1_2m * width_factor
        else:
            base_price = 0.0
        
        # Используем ту же логику расчета, что и в build_component_breakdown
        from core.optimization import OPT_CASCADING_PLAN
        width_mm = int(round(W * 1000))
        
        # Инициализация переменных
        long_cut_cost = 0.0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        rest_cost = 0.0
        waste_cost = 0.0
        
        # Проверяем OPT_CASCADING_PLAN для правильного подсчета резов, остатков и отходов
        if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):
            total_cuts_for_this_size = 0
            
            # Первичные резы
            for prim_cut in OPT_CASCADING_PLAN['primary_cuts']:
                if prim_cut['width'] == width_mm:
                    prim_lengths = prim_cut.get('lengths', [])
                    if not prim_lengths or any(abs(l - L) < 0.05 for l in prim_lengths):
                        prim_qty = prim_cut.get('qty', 0)
                        total_cuts_for_this_size += prim_qty  # Каждый первичный рез = 1 рез
                        primary_rest_width_mm = prim_cut['rest']

                        # --- Остатки: считаем только те, что образуются при резе ЭТОГО типа плит ---
                        unused_rest_total_mm = 0.0
                        if primary_rest_width_mm > 0:
                            produced_rests = prim_qty  # столько остатков этой ширины образовалось
                            used_rests = 0

                            # Считаем, сколько этих остатков реально использовано во вторичных резах
                            if OPT_CASCADING_PLAN.get('secondary_cuts'):
                                for sec_cut in OPT_CASCADING_PLAN['secondary_cuts']:
                                    if sec_cut.get('source') != primary_rest_width_mm:
                                        continue
                                    # Учитываем только те вторичные резы, которые берут остатки от ЭТОЙ длины
                                    src_lengths = sec_cut.get('source_lengths', [])
                                    if src_lengths and not any(abs(sl - L) < 0.05 for sl in src_lengths):
                                        continue
                                    used_rests += sec_cut.get('qty', 0)

                            unused_rests = max(0, produced_rests - used_rests)
                            unused_rest_total_mm = unused_rests * primary_rest_width_mm

                            if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                # Стоимость всех неиспользованных остатков, распределённая между плитами этого типа
                                rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty

                        # --- Отходы и поперечные резы: только те вторичные операции, которые дают ЭТИ плиты ---
                        if OPT_CASCADING_PLAN.get('secondary_cuts'):
                            for sec_cut in OPT_CASCADING_PLAN['secondary_cuts']:
                                sec_lengths = sec_cut.get('lengths', [])
                                if not sec_lengths or any(abs(l - L) < 0.05 for l in sec_lengths):
                                    sec_cuts = sec_cut.get('cuts', [])
                                    if width_mm in sec_cuts:
                                        sec_qty = sec_cut.get('qty', 0)
                                        sec_pieces = sec_cut.get('pieces', 1)

                                        # 1. Учет количества продольных резов
                                        current_cuts = sec_qty * sec_pieces
                                        total_cuts_for_this_size += current_cuts

                                        # 2. Отходы по ширине (waste по ширине распределяем между плитами этого типа)
                                        waste_w_mm = sec_cut.get('waste', 0)
                                        if waste_w_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                            cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price_1_2m
                                            waste_cost += (cost_of_waste_piece * sec_qty) / qty

                                        # 3. Поперечные резы и отходы по длине
                                        src_lens = sec_cut.get('source_lengths', [])
                                        if src_lens:
                                            src_len = src_lens[0]

                                            if sec_cut.get('type') == 'transverse' or abs(src_len - L) > 0.05:
                                                # Поперечные резы распределяем по плитам этого типа
                                                if qty > 0:
                                                    trans_cuts += (1.0 * sec_qty) / qty

                                                # Отходы по длине
                                                len_waste = src_len - L
                                                if len_waste > 0.01:
                                                    src_price_full = get_price(src_len, load_code, cfg.PRICE_DB_PATH)
                                                    if src_price_full is None:
                                                        src_price_full = find_price_for_plate(price_table, src_len, load_code) or 0.0

                                                    src_price_width = src_price_full * (width_mm / 1200.0)
                                                    cost_len_waste = src_price_width - base_price

                                                    if cost_len_waste > 0:
                                                        waste_cost += cost_len_waste * (sec_qty / qty)
                        break
            
            # Пересчитываем стоимость продольных резов с учетом реального количества
            if total_cuts_for_this_size > 0:
                long_cut_cost = (total_cuts_for_this_size * cfg.LONG_CUT_PRICE_PER_M * L) / qty if qty > 0 else 0
            else:
                # Если нет данных из плана, используем старую логику
                long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * L)
        else:
            # Если нет данных из плана, используем старую логику
            long_cut_cost = long_cuts * (cfg.LONG_CUT_PRICE_PER_M * L)

        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(W - 1.2) < 0.01:
            long_cut_cost = 0.0

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


def build_component_breakdown(price_table: dict, price_rows: list = None, reinforcement_code: int = 8):
    """Формирует детальную разбивку компонентов для каждого наименования."""
    from core.optimization import OPT_CASCADING_PLAN
    
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
                    # Формат: ПБ 36-0,3-8п (длина 3.6м, ширина 0.3м = 300мм)
                    # Или: ПБ 36-3,2-8п (длина 3.6м, ширина 3.2м = 3200мм - это неправильно)
                    # Правильный формат: ПБ 36-0,32-8п (длина 3.6м, ширина 0.32м = 320мм)
                    match = re.search(r'ПБ\s+(\d+)-([\d,]+)-', name)
                    if match:
                        length_dm = int(match.group(1))
                        length = length_dm / 10.0
                        width_str = match.group(2).replace(',', '.').replace(' ', '')
                        width_m = float(width_str)
                        # Если ширина больше 2, значит это в миллиметрах, иначе в метрах
                        if width_m > 2:
                            width_mm = int(width_m)
                            width_m = width_mm / 1000.0
                        else:
                            width_mm = int(width_m * 1000)
                        all_orders.append({
                            'length': length,
                            'width': width_mm,
                            'qty': qty
                        })
            if not all_orders:
                print('[DEBUG] build_component_breakdown: нет заказов, возвращаем пустой список')
                return []
        plan_orders = all_orders
    
    print(f'[DEBUG] build_component_breakdown: найдено заказов: {len(plan_orders)}')
    
    # Группируем заказы
    order_counter = Counter()
    for order in plan_orders:
        length = round(float(order.get('length', 0)), 3)
        width_val = order.get('width', 0)
        width_mm = width_val if width_val > 5 else int(width_val * 1000)
        order_counter[(length, width_mm)] += order.get('qty', 1)
    
    breakdown_tables = []
    
    for (length, width_mm), qty in sorted(order_counter.items(), key=lambda x: (x[0][0], x[0][1])):
        width_m = width_mm / 1000.0
        name = cfg.make_plate_name(length, width_m)
        
        # Получаем базовую цену за 1.2м
        load_code = 6 if width_m < 1.0 else reinforcement_code
        db_price = get_price(length, load_code, cfg.PRICE_DB_PATH)
        base_price_1_2m = db_price if db_price is not None else (find_price_for_plate(price_table, length, load_code) or 0.0)
        
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
        
        # Проверяем OPT_CASCADING_PLAN для определения остатков, отходов и количества резов
        total_cuts_count = 0  # Общее количество резов для отображения в таблице
        if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):
            # Считаем общее количество резов (первичных + вторичных) для всех плит этой ширины и длины
            total_cuts_for_this_size = 0
            total_plates_from_cuts = 0
            primary_rest_width_mm = 0  # Ширина остатка от первичных резов
            
            # Первичные резы
            for prim_cut in OPT_CASCADING_PLAN['primary_cuts']:
                if prim_cut['width'] == width_mm:
                    prim_lengths = prim_cut.get('lengths', [])
                    if not prim_lengths or any(abs(l - length) < 0.05 for l in prim_lengths):
                        prim_qty = prim_cut.get('qty', 0)
                        total_cuts_for_this_size += prim_qty  # Каждый первичный рез = 1 рез
                        total_plates_from_cuts += prim_qty   # Каждый первичный рез даёт 1 плиту
                        primary_rest_width_mm = prim_cut['rest']

                        # --- Остатки: считаем только те, что образуются при резе ЭТОГО типа плит ---
                        unused_rest_total_mm = 0.0
                        if primary_rest_width_mm > 0:
                            produced_rests = prim_qty  # столько остатков этой ширины образовалось
                            used_rests = 0

                            # Считаем, сколько этих остатков реально использовано во вторичных резах
                            if OPT_CASCADING_PLAN.get('secondary_cuts'):
                                for sec_cut in OPT_CASCADING_PLAN['secondary_cuts']:
                                    if sec_cut.get('source') != primary_rest_width_mm:
                                        continue
                                    # учитываем только вторичные резы от ЭТОЙ длины
                                    src_lengths = sec_cut.get('source_lengths', [])
                                    if src_lengths and not any(abs(sl - length) < 0.05 for sl in src_lengths):
                                        continue
                                    used_rests += sec_cut.get('qty', 0)

                            unused_rests = max(0, produced_rests - used_rests)
                            unused_rest_total_mm = unused_rests * primary_rest_width_mm

                            if unused_rest_total_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                # Стоимость всех неиспользованных остатков, распределённая между плитами этого типа
                                rest_cost = (unused_rest_total_mm / 1200.0) * base_price_1_2m / qty
                                rest_width_mm = unused_rest_total_mm  # для отображения
                            elif produced_rests > 0:
                                # Остатки были, но полностью использованы
                                rest_used = True
                                rest_width_mm = 0

                        # --- Вторичные резы и отходы: только операции, которые дают ЭТИ плиты ---
                        if OPT_CASCADING_PLAN.get('secondary_cuts'):
                            for sec_cut in OPT_CASCADING_PLAN['secondary_cuts']:
                                sec_lengths = sec_cut.get('lengths', [])
                                if not sec_lengths or any(abs(l - length) < 0.05 for l in sec_lengths):
                                    sec_cuts = sec_cut.get('cuts', [])
                                    
                                    # Проверяем, относится ли рез к нашей ширине
                                    if width_mm in sec_cuts:
                                        sec_qty = sec_cut.get('qty', 0)
                                        sec_pieces = sec_cut.get('pieces', 1)
                                        
                                        # 1. Учет количества продольных резов и плит
                                        current_cuts = sec_qty * sec_pieces
                                        total_cuts_for_this_size += current_cuts
                                        total_plates_from_cuts += current_cuts
                                        
                                        # 2. Отходы по ширине (серые зоны «отход» на визуализации),
                                        # распределяем стоимость между ВСЕМИ плитами данного типа
                                        waste_w_mm = sec_cut.get('waste', 0)
                                        if waste_w_mm > 0 and base_price_1_2m > 0 and qty > 0:
                                            cost_of_waste_piece = (waste_w_mm / 1200.0) * base_price_1_2m
                                            waste_cost += (cost_of_waste_piece * sec_qty) / qty
                                            # Для отображения запоминаем ширину и количество операций
                                            waste_terms.append((waste_w_mm, sec_qty))

                                        # 3. Поперечные резы (transverse) и отходы по длине
                                        src_lens = sec_cut.get('source_lengths', [])
                                        if src_lens:
                                            src_len = src_lens[0]

                                            # Если была операция поперечного реза или изменилась длина
                                            if sec_cut.get('type') == 'transverse' or abs(src_len - length) > 0.05:
                                                # Количество поперечных резов распределяем по плитам этого типа
                                                if qty > 0:
                                                    trans_cuts += (1.0 * sec_qty) / qty

                                                # Отходы по длине (серые обрезки по длине)
                                                len_waste = src_len - length
                                                if len_waste > 0.01:
                                                    src_price_full = get_price(src_len, load_code, cfg.PRICE_DB_PATH)
                                                    if src_price_full is None:
                                                        src_price_full = find_price_for_plate(price_table, src_len, load_code) or 0.0

                                                    src_price_width = src_price_full * (width_mm / 1200.0)
                                                    cost_len_waste = src_price_width - base_price

                                                    if cost_len_waste > 0:
                                                        waste_cost += cost_len_waste * (sec_qty / qty)

                            # После подсчёта trans_cuts обновляем стоимость поперечных резов
                            trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
                        break
            
            # Если нашли резы в плане, используем их количество
            if total_plates_from_cuts > 0:
                # ИСПРАВЛЕНИЕ: Используем общее количество резов, а не делим на количество плит
                # Стоимость реза на одну плиту = (общее количество резов × стоимость одного реза) / количество плит в заказе
                total_cuts_count = total_cuts_for_this_size  # Сохраняем для отображения в таблице
                long_cuts = total_cuts_for_this_size  # Общее количество резов
                # Пересчитываем стоимость продольных резов: общая стоимость всех резов / количество плит
                long_cut_cost = (long_cuts * cfg.LONG_CUT_PRICE_PER_M * length) / qty if qty > 0 else 0

        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
            long_cut_cost = 0.0

        # ИТОГО за 1 плиту
        total_per_unit = base_price + long_cut_cost + trans_cut_cost + rest_cost + waste_cost
        
        # Округление (до 2 знаков после запятой)
        total_rounded = round(total_per_unit, 2)
        
        # За N плит
        total_for_qty = total_rounded * qty
        
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

