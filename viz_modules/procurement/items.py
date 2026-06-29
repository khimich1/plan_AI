from __future__ import annotations

from collections import Counter
from collections.abc import Mapping
from typing import Any

from core.config_and_data import get_exact_width
from core.plate_runtime_state import get_plate_mutable_runtime
from core.optimization import OPT_PLAN

from .load_context import LoadCodeFn, resolve_procurement_load_context
from .orders import get_orders_from_opt_plan
from .plan_lookup import _block_ab_key, _length_dm_raw_from_m


def build_procurement_items(
    *,
    plate_load_details: Mapping[tuple[float, float, Any, str], int] | None = None,
    get_load_code: LoadCodeFn | None = None,
):
    """Формирует реальные позиции закупки из заказа пользователя."""
    rt = get_plate_mutable_runtime()
    items = []
    details, resolve_load = resolve_procurement_load_context(
        plate_load_details=plate_load_details,
        get_load_code=get_load_code,
    )

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
                if details:
                    for key, q in details.items():
                        L, W, load = key[0], key[1], key[2]
                        key_ldr = key[3] if len(key) > 3 else ''
                        if abs(L - length) < 0.05 and abs(W - width_m) < 0.01:
                            found_details.append((load, q, key_ldr))

                total_found = sum(q for _, q, _ in found_details)

                if found_details and total_found == order['qty']:
                    for load_code, q, detail_ldr in found_details:
                        order_counter[(length, width_m, load_code, detail_ldr, False)] += q
                else:
                    load_code = resolve_load(length, width_m, default=(6 if width_m < 1.0 else 8))
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
            length_dm_raw_val = ldr or rt.plate_length_dm_raw.get(full_key, '') or _length_dm_raw_from_m(length)
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

      # Приоритет 2: plate_load_details или списки rt.plates_* (legacy fallback)
    # plate_load_details содержит (длина, ширина, нагрузка) → количество
    if details:
        for key, qty in sorted(
            details.items(),
            key=lambda x: (_block_ab_key(x[0][1]), x[0][1], x[0][0], x[0][2]),
        ):
            length, width_m, load_code = key[0], key[1], key[2]
            ldr = key[3] if len(key) > 3 else rt.plate_length_dm_raw.get(key, '')

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
            cached = rt.plate_nomenclature_cache.get(key)
            if cached:
                if cached.get('canonical_name') is not None:
                    item['canonical_name'] = cached['canonical_name']
                if cached.get('nomenclature_id') is not None:
                    item['nomenclature_id'] = cached['nomenclature_id']
            items.append(item)
        return items
    
    # Legacy fallback: rt.plates_* когда plate_load_details пуст
    all_plates = []
    # ВАЖНО: Добавлен target_name для получения точных ширин из PLATE_EXACT_WIDTHS
    for width_mm, plates_list, target_name in [
        (320, rt.plates_0_32, 'PLATES_0_32'), (460, rt.plates_0_46, 'PLATES_0_46'), (700, rt.plates_0_70, 'PLATES_0_70'),
        (720, rt.plates_0_72, 'PLATES_0_72'), (860, rt.plates_0_86, 'PLATES_0_86'), (880, rt.plates_0_88, 'PLATES_0_88'),
        (740, rt.plates_0_74, 'PLATES_0_74'), (480, rt.plates_0_48, 'PLATES_0_48'), (500, rt.plates_0_50, 'PLATES_0_50'),
        (340, rt.plates_0_34, 'PLATES_0_34'), (1080, rt.plates_1_08, 'PLATES_1_08'), (1200, rt.plates_1_2, 'PLATES_1_2'),
        (1000, rt.plates_1_0, 'PLATES_1_0')
    ]:
        if plates_list:
            length_counts = Counter(plates_list)
            for length, qty in length_counts.items():
                # Получаем ТОЧНУЮ ширину из PLATE_EXACT_WIDTHS
                exact_width_m = get_exact_width(length, target_name, width_mm / 1000.0)
                
                # Получаем нагрузку для этой плиты
                load_code = resolve_load(length, exact_width_m, default=(6 if exact_width_m < 1.0 else 8))
                
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
        '0.32': mismatch_count(rt.plates_0_32, rt.plates_0_88),
        '0.46': mismatch_count(rt.plates_0_46, rt.plates_0_74),
        '0.72': mismatch_count(rt.plates_0_72, rt.plates_0_48),
        '0.70': mismatch_count(rt.plates_0_70, rt.plates_0_50),
        '0.86': mismatch_count(rt.plates_0_86, rt.plates_0_34),
    }
    
    # Округление до 3 знаков: сохраняем 5.71 (ПБ 57,1), не схлопываем в 5.7 (ПБ 57)
    _r = lambda x: round(x, 3)
    for L in rt.plates_1_2:
        items.append({'length': _r(L), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'as_is'})
    for L in rt.plates_1_5_to_1_2:
        items.append({'length': _r(L), 'width': 1.2, 'qty': 1, 'long_cuts': 0, 'trans_cuts': 0, 'purpose': 'to_1_2_main'})
        items.append({'length': _r(L), 'width': 0.3, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_2_strip'})
    for L in rt.plates_1_0:
        items.append({'length': _r(L), 'width': 1.0, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_main'})
        items.append({'length': _r(L), 'width': 0.2, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_0_strip'})
    for L in rt.plates_1_08:
        items.append({'length': _r(L), 'width': 1.08, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_main'})
        items.append({'length': _r(L), 'width': 0.12, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_1_08_strip'})
    for L in rt.plates_0_46:
        items.append({'length': _r(L), 'width': 0.46, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_main'})
        items.append({'length': _r(L), 'width': 0.74, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_46_strip'})
    
    mismatch = pair_plan['0.32']
    for idx, L in enumerate(rt.plates_0_32):
        items.append({'length': _r(L), 'width': 0.32, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_32_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.88, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_32_strip'})
    
    mismatch = pair_plan['0.72']
    for idx, L in enumerate(rt.plates_0_72):
        items.append({'length': _r(L), 'width': 0.72, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_72_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.48, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_72_strip'})
    
    mismatch = pair_plan['0.70']
    for idx, L in enumerate(rt.plates_0_70):
        items.append({'length': _r(L), 'width': 0.70, 'qty': 1, 'long_cuts': 1, 'trans_cuts': 0, 'purpose': 'to_0_70_main'})
        trans = 1 if idx < mismatch else 0
        items.append({'length': _r(L), 'width': 0.50, 'qty': 1, 'long_cuts': 1, 'trans_cuts': trans, 'purpose': 'to_0_70_strip'})
    
    mismatch = pair_plan['0.86']
    for idx, L in enumerate(rt.plates_0_86):
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
