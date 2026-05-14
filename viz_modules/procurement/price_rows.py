from __future__ import annotations

import json
import time

import core.config_and_data as cfg

from ..price_utils import _find_price_for_plate_production_fallback, find_price_for_plate
from .debug_logs import _DEBUG_LOG_8E9428, _DEBUG_LOG_A9176E, _DEBUG_LOG_DB7A51
from .items import build_procurement_items
from .plan_lookup import _find_plan_for_plate
from .ports import ProcurementDeps, resolve_procurement_deps
from .trim import _calc_trim_components


def build_price_rows(price_table: dict, reinforcement_code: int = 8, deps: ProcurementDeps | None = None):
    """Формирует строки сметы."""
    d = resolve_procurement_deps(deps)
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
                _log_path = _DEBUG_LOG_8E9428
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_price_row", "location": "procurement:build_price_rows", "message": "57/57,1: price_row name", "data": {"L": L, "name": name, "canonical_name": it.get('canonical_name')}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
        # #region agent log (a9176e: 57/57,1 — build_price_rows имя)
        if 5.69 <= L <= 5.73:
            try:
                import os as _os
                _log_path = _DEBUG_LOG_A9176E
                import json as _json
                _pay = {"sessionId": "a9176e", "hypothesisId": "H2", "location": "procurement:build_price_rows", "message": "57/57,1 price_row name source", "data": {"L": L, "canonical_name": it.get("canonical_name"), "length_dm_raw": it.get("length_dm_raw"), "name": name}, "timestamp": __import__("time").time() * 1000}
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(_json.dumps(_pay, ensure_ascii=False) + "\n")
            except Exception:
                pass
        # #endregion
        if it.get('warning'):
            name += " (нагрузка?)"
        db_price = d.get_price(L, load_code, d.db_path)
        use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
        find_price = find_price_for_plate(price_table, L, load_code) if use_fallback else None
        base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
        # #region agent log
        import json
        import os
        import time
        if base_price_1_2m == 0.0:
            _log_path = _DEBUG_LOG_DB7A51
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
            deps=d,
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


def build_price_rows_production(price_table: dict, reinforcement_code: int = 8, deps: ProcurementDeps | None = None):
    """
    Формирует строки сметы для планирования производства.
    ОТЛИЧИЯ от build_price_rows:
    - Базовая цена берется из таблицы raw_material_costs (стоимость сырья + производство)
    - Добавлен компонент "Переармирование" (перерасход прутьев)
    """
    d = resolve_procurement_deps(deps)
    
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
            
            reinforcement = d.get_reinforcement(L, load_code, db_path=d.db_path)
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
        base_price_1_2m = d.get_raw_material_cost(base_name_1_2m_short, db_path=d.db_path)
        
        if base_price_1_2m is not None:
            # Пересчитываем на фактическую ширину
            width_factor = W / 1.2
            base_price = base_price_1_2m * width_factor
            print(f'[PRODUCTION PRICING] {name}: базовая цена из БД ({base_name_1_2m_short}: {base_price_1_2m:.2f}) × {width_factor:.3f} = {base_price:.2f} руб')
        else:
            # Fallback: если нет в БД, используем старый метод
            db_price = d.get_price(L, load_code, d.db_path)
            use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
            find_price = (
                _find_price_for_plate_production_fallback(price_table, L, load_code)
                if use_fallback
                else None
            )
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
                                                    src_price_full = d.get_raw_material_cost(
                                                        src_plate_name,
                                                        db_path=d.db_path
                                                    )
                                                    if src_price_full is None:
                                                        # Fallback
                                                        src_price_full_db = d.get_price(src_len, load_code, d.db_path)
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
        reinforcement = d.get_reinforcement(L, load_code, db_path=d.db_path)
        
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
