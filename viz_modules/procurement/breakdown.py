from __future__ import annotations

import math
import re
from collections import Counter
from collections.abc import Mapping
from typing import Any

from core.config.constants import (
    LONG_CUT_PRICE_PER_M,
    MIN_BILLABLE_TRIM_MM,
    TRANSVERSE_CUT_PRICE,
    length_dm_to_m,
)
from core.config_and_data import make_plate_name
from core.plate_runtime_state import get_plate_mutable_runtime

from ..price_utils import _find_price_for_plate_production_fallback, find_price_for_plate
from .load_context import LoadCodeFn, resolve_procurement_load_context
from .orders import get_orders_from_opt_plan
from .plan_lookup import _find_plan_for_plate
from .ports import ProcurementDeps, resolve_procurement_deps
from .trim import (
    _calc_trim_components,
    apply_factory_strip_waste,
    format_long_cut_calculation,
    format_transverse_remainder_calculation,
    resolve_long_cut_pricing,
)


def _accumulate_order_counter(
    order_counter: Counter,
    plan_orders: list,
    *,
    reinforcement_code: int = 8,
    plate_load_details: Mapping[tuple[float, float, Any, str], int] | None = None,
    get_load_code: LoadCodeFn | None = None,
) -> None:
    """Группирует строки заказа по (длина, ширина, нагрузка, номинал_длины).

    Приоритет нагрузки: ``load_code`` на строке заказа (из плана оптимизатора),
    иначе ``plate_load_details``, иначе ``get_load_code_for_plate``.
    """
    details, resolve_load = resolve_procurement_load_context(
        plate_load_details=plate_load_details,
        get_load_code=get_load_code,
    )
    for order in plan_orders:
        length = round(float(order.get('length', 0)), 3)
        width_val = order.get('width', 0)
        width_mm = width_val if width_val > 5 else round(width_val * 1000)
        width_m = width_mm / 1000.0
        order_ldr = (order.get('length_dm_raw') or '').strip()
        order_qty = order.get('qty', 1)

        order_load = order.get('load_code')
        if order_load is not None and order_load != '':
            load_code = int(math.floor(float(order_load)))
            order_counter[(length, width_mm, load_code, order_ldr, False)] += order_qty
            continue

        found_details = []
        if details:
            for key, q in details.items():
                L, W, load = key[0], key[1], key[2]
                key_ldr = key[3] if len(key) > 3 else ''
                if abs(L - length) < 0.05 and abs(W - width_m) < 0.01:
                    found_details.append((load, q, key_ldr))

        if found_details and sum(q for _, q, _ in found_details) == order_qty:
            for load_code, q, detail_ldr in found_details:
                order_counter[(length, width_mm, load_code, detail_ldr, False)] += q
        else:
            load_code = resolve_load(
                length, width_m, default=(6 if width_m < 1.0 else reinforcement_code)
            )
            warning_flag = bool(found_details)
            order_counter[(length, width_mm, load_code, order_ldr, warning_flag)] += order_qty


def build_component_breakdown(
    price_table: dict,
    price_rows: list = None,
    reinforcement_code: int = 8,
    deps: ProcurementDeps | None = None,
    *,
    plate_load_details: Mapping[tuple[float, float, Any, str], int] | None = None,
    get_load_code: LoadCodeFn | None = None,
):
    """Формирует детальную разбивку компонентов для каждого наименования."""
    d = resolve_procurement_deps(deps)
    rt = get_plate_mutable_runtime()
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    
    # Получаем заказы
    plan_orders = get_orders_from_opt_plan()
    if not plan_orders:
        # Fallback на старый способ
        all_orders = []
        for width_mm, plates_list in [
            (320, rt.plates_0_32), (460, rt.plates_0_46), (700, rt.plates_0_70),
            (720, rt.plates_0_72), (860, rt.plates_0_86), (880, rt.plates_0_88),
            (740, rt.plates_0_74), (480, rt.plates_0_48), (500, rt.plates_0_50),
            (340, rt.plates_0_34), (1080, rt.plates_1_08)
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
                        length = length_dm_to_m(match.group(1))  # Целое = номинал в дм (длина = номинал/10 м); с запятой/точкой = точное значение в дм
                        
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

    order_counter: Counter = Counter()
    _accumulate_order_counter(
        order_counter,
        plan_orders,
        reinforcement_code=reinforcement_code,
        plate_load_details=plate_load_details,
        get_load_code=get_load_code,
    )

    breakdown_tables = []

    for (length, width_mm, load_code, ldr, warning_flag), qty in sorted(order_counter.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        width_m = width_mm / 1000.0

        # Имя плиты в детальной разбивке тоже должно отражать фактическую нагрузку
        name = make_plate_name(length, width_m, load_code=load_code, length_dm_raw=ldr or None)
        if warning_flag:
            name += " (нагрузка?)"
        db_price = d.get_price(length, load_code, d.db_path)
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

        fallback_long_cuts = 0 if abs(width_m - 1.2) < 0.01 else (1 if width_m < 1.15 else 0)
        long_cuts = 0
        long_cut_cost = 0.0
        trans_cuts = 0
        trans_cut_cost = 0.0
        rest_width_mm = 0
        rest_cost = 0.0
        rest_used = False
        waste_cost = 0.0
        waste_terms = []
        
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
            deps=d,
        )
        rest_cost = trim['rest_cost']
        rest_width_mm = trim['rest_width_mm']
        rest_used = trim['rest_used']
        waste_cost = trim['waste_cost']
        waste_terms = trim['waste_terms']
        transverse_remainder_cost = trim['transverse_remainder_cost']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * TRANSVERSE_CUT_PRICE

        long_cut_cost, long_cuts, total_cuts_count = resolve_long_cut_pricing(
            trim,
            qty=qty,
            length=length,
            width_m=width_m,
            current_plan=current_plan,
            fallback_long_cuts=fallback_long_cuts,
            plate_name=name,
        )

        if not (current_plan and current_plan.get('primary_cuts')):
            if trim.get('long_cut_meterage', 0) <= 0 and width_m < 1.15:
                # 1020–1080 мм: отход factory strip — через apply_factory_strip_waste (R5).
                if not (1020 <= width_mm <= 1080):
                    print(f'[DEBUG] Плана оптимизации нет для {name}, используем ручной расчёт остатков')
                    rest_width_mm = 1200 - width_mm
                    if rest_width_mm > MIN_BILLABLE_TRIM_MM and base_price_1_2m > 0:
                        rest_cost = (rest_width_mm / 1200.0) * base_price_1_2m
                        print(f'[DEBUG] Остаток {rest_width_mm}мм не использован, добавляем к цене: {rest_cost:.2f} руб')
                    else:
                        rest_cost = 0.0

        waste_cost, waste_terms = apply_factory_strip_waste(
            width_mm=width_mm,
            base_price_1_2m=base_price_1_2m,
            rest_cost=rest_cost,
            rest_used=rest_used,
            waste_cost=waste_cost,
            waste_terms=waste_terms,
            qty=qty,
        )

        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
            long_cut_cost = 0.0

        # ИТОГО за 1 плиту
        total_per_unit = (
            base_price
            + long_cut_cost
            + trans_cut_cost
            + rest_cost
            + waste_cost
            + transverse_remainder_cost
        )
        total_rounded = round(total_per_unit, 2)
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
        
        if long_cut_cost > 0:
            long_calc = format_long_cut_calculation(trim, qty)
            if not long_calc:
                if abs(long_cuts - 1.0) > 0.001:
                    long_calc = f"{LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {long_cuts:.2f}".replace('.', ',')
                else:
                    long_calc = f"{LONG_CUT_PRICE_PER_M:.0f} × {length:.1f}".replace('.', ',')
            table_rows.append([
                "Продольный рез",
                long_calc,
                f"{long_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Поперечный рез
        if trans_cuts > 0:
            trans_calc = f"{TRANSVERSE_CUT_PRICE:.0f} × {trans_cuts}"
            table_rows.append([
                "Поперечный рез",
                trans_calc,
                f"{trans_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])

        if transverse_remainder_cost > 0:
            rem_terms = trim.get('transverse_remainder_terms') or []
            rem_m = rem_terms[0][0] if rem_terms else 0.0
            rem_label = f"{rem_m:.2f}".replace('.', ',')
            trans_rem_calc = format_transverse_remainder_calculation(
                trim,
                qty,
                base_price_1_2m=base_price_1_2m,
                width_m=width_m,
                length_m=length,
            ) or ""
            table_rows.append([
                f"Остаток после поперечного реза ({rem_label}м)",
                trans_rem_calc,
                f"{transverse_remainder_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        if rest_cost > 0:
            base_price_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
            rest_calc = f"({rest_width_mm} / 1200) × {base_price_str} / {qty}"
            table_rows.append([
                f"Остаток ({rest_width_mm}мм)",
                rest_calc,
                f"{rest_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        elif rest_used:
            table_rows.append([
                "Остаток (использован в каскаде)",
                "0 (использован)",
                "0,00 руб"
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


def build_component_breakdown_production(
    price_table: dict,
    price_rows: list = None,
    reinforcement_code: int = 8,
    tracks_for_day: list = None,
    deps: ProcurementDeps | None = None,
    *,
    plate_load_details: Mapping[tuple[float, float, Any, str], int] | None = None,
    get_load_code: LoadCodeFn | None = None,
):
    """
    Формирует детальную разбивку компонентов для планирования производства.
    ОТЛИЧИЯ от build_component_breakdown:
    - Базовая цена берется из таблицы raw_material_costs
    - Добавлен компонент "Переармирование"
    
    Args:
        tracks_for_day: Список дорожек текущего дня. Если указан, будут включены только плиты из этих дорожек.
    """
    d = resolve_procurement_deps(deps)
    rt = get_plate_mutable_runtime()
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    
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
                (320, rt.plates_0_32), (460, rt.plates_0_46), (700, rt.plates_0_70),
                (720, rt.plates_0_72), (860, rt.plates_0_86), (880, rt.plates_0_88),
                (740, rt.plates_0_74), (480, rt.plates_0_48), (500, rt.plates_0_50),
                (340, rt.plates_0_34), (1080, rt.plates_1_08)
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
                            length = length_dm_to_m(match.group(1))
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

    order_counter: Counter = Counter()
    _accumulate_order_counter(
        order_counter,
        plan_orders,
        reinforcement_code=reinforcement_code,
        plate_load_details=plate_load_details,
        get_load_code=get_load_code,
    )

    # ШАГ 1: Определяем максимальное армирование
    # НОВОЕ: Используем PLATE_MAX_REINFORCEMENT_MAP если он заполнен (макс. армирование по дорожке)
    # Иначе - fallback на максимум по всему заказу
    use_track_based_reinforcement = bool(rt.plate_max_reinforcement_map)

    if use_track_based_reinforcement:
        print(f'[PRODUCTION BREAKDOWN] ✅ Используем максимальное армирование по ДОРОЖКАМ (из PLATE_MAX_REINFORCEMENT_MAP)')
    else:
        # Fallback: находим максимальное армирование во всем заказе
        global_max_reinforcement = 0.0
        for (length, width_mm, load_code, ldr, warning_flag), qty in order_counter.items():
            reinforcement = d.get_reinforcement(length, load_code, db_path=d.db_path)
            if reinforcement and reinforcement > global_max_reinforcement:
                global_max_reinforcement = reinforcement
        print(f'[PRODUCTION BREAKDOWN] ⚠️ PLATE_MAX_REINFORCEMENT_MAP пуст, используем глобальный максимум: {global_max_reinforcement} прутьев')

    breakdown_tables = []

    for (length, width_mm, load_code, ldr, warning_flag), qty in sorted(order_counter.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        width_m = width_mm / 1000.0

        # Имя плиты
        name = make_plate_name(length, width_m, load_code=load_code, length_dm_raw=ldr or None)
        if warning_flag:
            name += " (нагрузка?)"
        
        # ✅ ИЗМЕНЕНИЕ 1: Базовая цена из raw_material_costs
        # В БД все плиты имеют формат "ПБ XX-12-НАГРУЗКА" (ширина 1.2м)
        # Поэтому ищем по базовому имени с той же нагрузкой и пересчитываем на фактическую ширину
        base_name_1_2m = make_plate_name(length, 1.2, load_code=load_code)
        # Убираем "Плиты " и букву "п", т.к. в БД хранится формат "ПБ 23-12-8"
        base_name_1_2m_short = base_name_1_2m.replace('Плиты ', '').replace('п', '')
        base_price_1_2m = d.get_raw_material_cost(base_name_1_2m_short, db_path=d.db_path)
        
        if base_price_1_2m is not None:
            # Пересчитываем на фактическую ширину
            width_factor = width_m / 1.2
            base_price = base_price_1_2m * width_factor
            print(f'[PRODUCTION BREAKDOWN] {name}: базовая цена из БД ({base_name_1_2m_short}: {base_price_1_2m:.2f}) × {width_factor:.3f} = {base_price:.2f} руб')
        else:
            # Fallback
            db_price = d.get_price(length, load_code, d.db_path)
            use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
            find_price = (
                _find_price_for_plate_production_fallback(price_table, length, load_code)
                if use_fallback
                else None
            )
            base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
            if base_price_1_2m > 0:
                width_factor = width_m / 1.2
                base_price = base_price_1_2m * width_factor
            else:
                base_price = 0.0
                base_price_1_2m = 0.0
            print(f'[WARNING] Нет данных в raw_material_costs для {base_name_1_2m_short}, использую старый метод: {base_price:.2f}')

        fallback_long_cuts = 0 if abs(width_m - 1.2) < 0.01 else (1 if width_m < 1.15 else 0)
        long_cuts = 0
        long_cut_cost = 0.0
        trans_cuts = 0
        trans_cut_cost = 0.0
        rest_width_mm = 0
        rest_cost = 0.0
        rest_used = False
        waste_cost = 0.0
        waste_terms = []

        current_plan, load_key = _find_plan_for_plate(
            load_code, length, width_mm, name, 'build_component_breakdown_production'
        )
        if not current_plan:
            print(
                f'[DEBUG] build_component_breakdown_production: не найден подходящий план '
                f'для плиты {name} при нагрузке {load_key}п.'
            )

        total_cuts_count = 0
        trim = _calc_trim_components(
            current_plan,
            length=length,
            width_mm=width_mm,
            qty=qty,
            base_price_1_2m=base_price_1_2m,
            base_price=base_price,
            load_code=load_code,
            price_table=price_table,
            deps=d,
        )
        rest_cost = trim['rest_cost']
        rest_width_mm = trim['rest_width_mm']
        rest_used = trim['rest_used']
        waste_cost = trim['waste_cost']
        waste_terms = trim['waste_terms']
        transverse_remainder_cost = trim['transverse_remainder_cost']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * TRANSVERSE_CUT_PRICE

        long_cut_cost, long_cuts, total_cuts_count = resolve_long_cut_pricing(
            trim,
            qty=qty,
            length=length,
            width_m=width_m,
            current_plan=current_plan,
            fallback_long_cuts=fallback_long_cuts,
            plate_name=name,
        )

        if not (current_plan and current_plan.get('primary_cuts')):
            if trim.get('long_cut_meterage', 0) <= 0 and width_m < 1.15:
                # 1020–1080 мм: отход factory strip — через apply_factory_strip_waste (R5).
                if not (1020 <= width_mm <= 1080):
                    print(f'[DEBUG] Плана оптимизации нет для {name}, используем ручной расчёт остатков')
                    rest_width_mm = 1200 - width_mm
                    if rest_width_mm > MIN_BILLABLE_TRIM_MM and base_price_1_2m > 0:
                        rest_cost = (rest_width_mm / 1200.0) * base_price_1_2m
                        print(f'[DEBUG] Остаток {rest_width_mm}мм не использован, добавляем к цене: {rest_cost:.2f} руб')
                    else:
                        rest_cost = 0.0

        waste_cost, waste_terms = apply_factory_strip_waste(
            width_mm=width_mm,
            base_price_1_2m=base_price_1_2m,
            rest_cost=rest_cost,
            rest_used=rest_used,
            waste_cost=waste_cost,
            waste_terms=waste_terms,
            qty=qty,
        )

        # Жёсткое правило: плиты 1.2м без продольных резов
        if abs(width_m - 1.2) < 0.01:
            long_cuts = 0
            long_cut_cost = 0.0

        # ✅ ИЗМЕНЕНИЕ 2: Переармирование
        # НОВОЕ: Используем максимальное армирование ДОРОЖКИ для этой плиты
        rearm_cost = 0.0
        reinforcement = d.get_reinforcement(length, load_code, db_path=d.db_path)
        
        # Получаем max_reinforcement для этой конкретной плиты
        if use_track_based_reinforcement:
            # Ищем в карте по (length, width_mm)
            plate_key = (round(length, 3), width_mm)
            max_reinforcement = rt.plate_max_reinforcement_map.get(plate_key, 0)
            if max_reinforcement == 0:
                # Попробуем с другими округлениями
                for (l, w), mr in rt.plate_max_reinforcement_map.items():
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
        total_per_unit = (
            base_price
            + long_cut_cost
            + trans_cut_cost
            + rest_cost
            + waste_cost
            + transverse_remainder_cost
            + rearm_cost
        )
        total_rounded = round(total_per_unit, 2)
        total_for_qty = total_rounded * qty
        
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
        
        if long_cut_cost > 0:
            long_calc = format_long_cut_calculation(trim, qty)
            if not long_calc:
                if abs(long_cuts - 1.0) > 0.001:
                    long_calc = f"{LONG_CUT_PRICE_PER_M:.0f} × {length:.1f} × {long_cuts:.2f}".replace('.', ',')
                else:
                    long_calc = f"{LONG_CUT_PRICE_PER_M:.0f} × {length:.1f}".replace('.', ',')
            table_rows.append([
                "Продольный рез",
                long_calc,
                f"{long_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        # Поперечный рез
        if trans_cuts > 0:
            trans_calc = f"{TRANSVERSE_CUT_PRICE:.0f} × {trans_cuts}"
            table_rows.append([
                "Поперечный рез",
                trans_calc,
                f"{trans_cut_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])

        if transverse_remainder_cost > 0:
            rem_terms = trim.get('transverse_remainder_terms') or []
            rem_m = rem_terms[0][0] if rem_terms else 0.0
            rem_label = f"{rem_m:.2f}".replace('.', ',')
            trans_rem_calc = format_transverse_remainder_calculation(
                trim,
                qty,
                base_price_1_2m=base_price_1_2m,
                width_m=width_m,
                length_m=length,
            ) or ""
            table_rows.append([
                f"Остаток после поперечного реза ({rem_label}м)",
                trans_rem_calc,
                f"{transverse_remainder_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        
        if rest_cost > 0:
            base_price_1_2m_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
            rest_calc = f"({rest_width_mm} / 1200) × {base_price_1_2m_str} / {qty}"
            table_rows.append([
                f"Остаток ({rest_width_mm}мм)",
                rest_calc,
                f"{rest_cost:,.2f} руб".replace(',', ' ').replace('.', ',')
            ])
        elif rest_used:
            table_rows.append([
                "Остаток (использован в каскаде)",
                "0 (использован)",
                "0,00 руб"
            ])
        
        # Отходы
        if waste_cost > 0 or waste_terms:
            if base_price_1_2m > 0:
                base_price_str = f"{base_price_1_2m:,.2f}".replace(',', ' ').replace('.', ',')
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
