from __future__ import annotations

import time

import core.config_and_data as cfg

from ..price_utils import _find_price_for_plate_production_fallback, find_price_for_plate
from core.debug_paths import append_agent_debug_log
from .debug_logs import _DEBUG_LOG_8E9428, _DEBUG_LOG_A9176E, _DEBUG_LOG_DB7A51
from .items import build_procurement_items
from .plan_lookup import _find_plan_for_plate
from .ports import ProcurementDeps, resolve_procurement_deps
from .trim import _calc_trim_components, apply_factory_strip_waste, resolve_long_cut_pricing


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
            append_agent_debug_log(
                _DEBUG_LOG_8E9428,
                {
                    "sessionId": "8e9428",
                    "hypothesisId": "H_price_row",
                    "location": "procurement:build_price_rows",
                    "message": "57/57,1: price_row name",
                    "data": {"L": L, "name": name, "canonical_name": it.get("canonical_name")},
                    "timestamp": __import__("time").time() * 1000,
                },
            )
        if 5.69 <= L <= 5.73:
            append_agent_debug_log(
                _DEBUG_LOG_A9176E,
                {
                    "sessionId": "a9176e",
                    "hypothesisId": "H2",
                    "location": "procurement:build_price_rows",
                    "message": "57/57,1 price_row name source",
                    "data": {
                        "L": L,
                        "canonical_name": it.get("canonical_name"),
                        "length_dm_raw": it.get("length_dm_raw"),
                        "name": name,
                    },
                    "timestamp": __import__("time").time() * 1000,
                },
            )
        if it.get('warning'):
            name += " (нагрузка?)"
        db_price = d.get_price(L, load_code, d.db_path)
        use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
        find_price = find_price_for_plate(price_table, L, load_code) if use_fallback else None
        base_price_1_2m = (db_price if (db_price is not None and isinstance(db_price, (int, float)) and db_price > 0) else None) or find_price or 0.0
        if base_price_1_2m == 0.0:
            append_agent_debug_log(
                _DEBUG_LOG_DB7A51,
                {
                    "sessionId": "db7a51",
                    "hypothesisId": "build_price_rows",
                    "location": "procurement.py:build_price_rows",
                    "message": "price chain",
                    "data": {
                        "name": name,
                        "L": L,
                        "W": W,
                        "load_code": load_code,
                        "db_price": db_price,
                        "find_price": find_price,
                        "base_price_1_2m": base_price_1_2m,
                    },
                    "timestamp": int(time.time() * 1000),
                },
            )
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
        rest_used = trim['rest_used']
        waste_cost = trim['waste_cost']
        transverse_remainder_cost = trim['transverse_remainder_cost']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE

        fallback_long = long_cuts if long_cuts else (0 if abs(W - 1.2) < 0.01 else (1 if W < 1.15 else 0))
        long_cut_cost, long_cuts, _ = resolve_long_cut_pricing(
            trim,
            qty=qty,
            length=L,
            width_m=W,
            current_plan=current_plan,
            fallback_long_cuts=fallback_long,
            plate_name=name,
        )

        # Жёсткое правило: плиты шириной 1.2 м считаем без продольных резов
        if abs(W - 1.2) < 0.01:
            long_cut_cost = 0.0

        waste_cost, _ = apply_factory_strip_waste(
            width_mm=width_mm,
            base_price_1_2m=base_price_1_2m,
            rest_cost=rest_cost,
            rest_used=rest_used,
            waste_cost=waste_cost,
            waste_terms=[],
            qty=qty,
        )

        unit_price = (
            base_price
            + long_cut_cost
            + trans_cut_cost
            + rest_cost
            + waste_cost
            + transverse_remainder_cost
        )
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
        
        width_mm = int(round(W * 1000))
        long_cut_cost = 0.0
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE
        rest_cost = 0.0
        waste_cost = 0.0
        rest_used = False

        current_plan, load_key = _find_plan_for_plate(
            load_code, L, width_mm, name, 'build_price_rows_production'
        )
        if not current_plan:
            print(
                f'[DEBUG] build_price_rows_production: не найден подходящий план для плиты '
                f'{name} при нагрузке {load_key}п.'
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
        rest_used = trim['rest_used']
        transverse_remainder_cost = trim['transverse_remainder_cost']
        trans_cuts += trim['trans_cuts']
        trans_cut_cost = trans_cuts * cfg.TRANSVERSE_CUT_PRICE

        fallback_long = long_cuts if long_cuts else (0 if abs(W - 1.2) < 0.01 else (1 if W < 1.15 else 0))
        long_cut_cost, long_cuts, _ = resolve_long_cut_pricing(
            trim,
            qty=qty,
            length=L,
            width_m=W,
            current_plan=current_plan,
            fallback_long_cuts=fallback_long,
            plate_name=name,
        )

        # Жёсткое правило: плиты 1.2м без продольных резов
        if abs(W - 1.2) < 0.01:
            long_cut_cost = 0.0

        waste_cost, _ = apply_factory_strip_waste(
            width_mm=width_mm,
            base_price_1_2m=base_price_1_2m,
            rest_cost=rest_cost,
            rest_used=rest_used,
            waste_cost=waste_cost,
            waste_terms=[],
            qty=qty,
        )

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

        unit_price = (
            base_price
            + long_cut_cost
            + trans_cut_cost
            + rest_cost
            + waste_cost
            + transverse_remainder_cost
            + rearm_cost
        )
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
