#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль построения последовательности раскладки плит:
- Формирование последовательности сегментов вдоль дорожки
"""
import core.config_and_data as cfg
from core.optimization import OPT_PLAN, OPT_WIDTH_PRIORITY
from core.debug_paths import get_debug_log_path
from collections import defaultdict

_DEBUG_LOG = get_debug_log_path("debug.log")
_DEBUG_LOG_95694E = get_debug_log_path("debug-95694e.log")
_DEBUG_LOG_2D5C43 = get_debug_log_path("debug-2d5c43.log")


def _choose_best_separator(solid_list, next_group, reinforcement_map):
    """
    Выбирает оптимальную плиту-разделитель по армированию.
    
    НОВАЯ Стратегия:
    Выбирает целую плиту с МИНИМАЛЬНЫМ армированием из оставшихся.
    
    Args:
        solid_list: список целых плит [{width, lengths, qty}, ...]
        next_group: следующая группа резов (не используется в новой логике)
        reinforcement_map: {(length, width_mm, load_code): reinforcement}
    
    Returns:
        index: индекс лучшей плиты в solid_list (или 0, если не найдено)
    """
    if not solid_list:
        return None
    
    # Собираем информацию об армировании для каждой целой плиты
    candidates = []
    for idx, plate in enumerate(solid_list):
        length = plate['lengths'][0] if plate.get('lengths') else 6.0
        width_mm = plate['width']
        reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm) or 999.0
        
        candidates.append({
            'index': idx,
            'length': length,
            'width_mm': width_mm,
            'reinforcement': reinforcement
        })
    
    # НОВАЯ ЛОГИКА: Выбираем плиту с МИНИМАЛЬНЫМ армированием из оставшихся
    best = min(candidates, key=lambda x: x['reinforcement'])
    print(f"[VISUAL] ✅ Выбран разделитель с мин. армированием: {best['length']:.2f}м x {best['width_mm']}мм, "
          f"армирование {best['reinforcement']:.1f} кг/м")
    return best['index']


def _canonical_load_code(load_code):
    """Единый ключ нагрузки для карты: 12.5 и 13 дают 13 (как в БД pb_reinforcement_series)."""
    if load_code is None:
        return None
    try:
        return int(float(load_code) + 0.5)
    except Exception:
        return None


def _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code=None):
    """
    Получает армирование из карты по (length, width_mm, load_code).
    
    Карта имеет ключи (length, width_mm, canonical_load_code).
    Если load_code указан - ищет по каноническому коду (12.5 и 13 → один ключ).
    Если load_code=None - ищет первое совпадение по (length, width_mm).
    
    Args:
        reinforcement_map: {(length, width_mm, load_code): reinforcement}
        length: длина плиты в метрах
        width_mm: ширина плиты в мм
        load_code: код нагрузки (опционально)
    
    Returns:
        reinforcement или None
    """
    if load_code is not None:
        canonical = _canonical_load_code(load_code)
        if canonical is not None:
            # Точный поиск по каноническому коду (12.5 и 13 оба дают 13)
            result = reinforcement_map.get((length, width_mm, canonical))
            # В БД 12,5п = нагрузка 13; при "12п" в заказе приходит 12 — если по 12 не нашли, берём по 13
            if result is None and canonical == 12:
                result = reinforcement_map.get((length, width_mm, 13))
            # #region agent log
            if abs(length - 7.1) < 0.05 and canonical in (12, 13):
                try:
                    keys_71 = [(k[0], k[1], k[2]) for k in list(reinforcement_map.keys())[:50] if abs(k[0] - 7.1) < 0.05]
                    open(_DEBUG_LOG, "a", encoding="utf-8").write(__import__("json").dumps({"hypothesisId": "H71lookup", "location": "layout_sequence:_get_reinforcement_from_map", "message": "71-12п поиск в карте", "data": {"length": length, "width_mm": width_mm, "load_code": load_code, "canonical": canonical, "key": (round(length, 3), width_mm, canonical), "result": result, "keys_71_in_map": keys_71[:10]}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            return result
    # Поиск по (length, width_mm) - возвращаем первое найденное
    for (l, w, lc), reinforcement in reinforcement_map.items():
        if abs(l - length) < 0.01 and w == width_mm:
            return reinforcement
    return None


def _ensure_sequence_layout_uid(sequence, prefix="seq"):
    """Гарантирует наличие identity у каждого root item sequence."""
    for idx, item in enumerate(sequence or []):
        if not isinstance(item, dict):
            continue
        if item.get("layout_uid"):
            continue
        unit_id = item.get("unit_id")
        item["layout_uid"] = str(unit_id) if unit_id else f"{prefix}:{idx}"


def _split_group_into_subgroups(cut_group, max_length=90.0):
    """
    Разбивает группу резов на подгруппы по max_length метров.
    
    ВАЖНО: Группа уже содержит плиты с одинаковым резом И армированием
    (благодаря предварительной группировке через groupby).
    Эта функция просто режет по длине, НЕ меняя порядок плит.
    
    Args:
        cut_group: список записей [{width, rest, qty, lengths, reinforcement}, ...]
        max_length: максимальная длина подгруппы в метрах (по умолчанию 90м)
    
    Returns:
        list[list]: Список подгрупп [[cut1, cut2], [cut3, cut4], ...]
        
    Example:
        >>> group = [
        ...     {'width': 320, 'rest': 880, 'qty': 10, 'lengths': [8.0]*10},
        ...     {'width': 320, 'rest': 880, 'qty': 5, 'lengths': [7.5]*5}
        ... ]
        >>> subgroups = _split_group_into_subgroups(group, max_length=90.0)
        >>> # Результат: [[10 плит по 8м], [5 плит по 7.5м и ещё...]]
    """
    subgroups = []
    current_subgroup = []
    current_length = 0.0
    
    print(f"[VISUAL] Разбиваю группу на подгруппы (макс {max_length}м)...")
    
    for cut in cut_group:
        qty = cut['qty']
        lengths = cut.get('lengths', [])
        
        # Обрабатываем каждую плиту в записи
        for i in range(qty):
            length = lengths[i] if i < len(lengths) else (lengths[0] if lengths else 6.0)
            
            # Проверяем: влезет ли плита в текущую подгруппу?
            if current_length + length > max_length and current_subgroup:
                # Закрываем текущую подгруппу
                subgroups.append(current_subgroup)
                print(f"[VISUAL]   Подгруппа #{len(subgroups)} закрыта: {current_length:.1f}м ({len(current_subgroup)} записей)")
                current_subgroup = []
                current_length = 0.0
            
            # Создаём запись для 1 плиты (разворачиваем qty=10 → 10 записей с qty=1)
            single_cut = cut.copy()
            single_cut['qty'] = 1
            single_cut['lengths'] = [length]
            
            current_subgroup.append(single_cut)
            current_length += length
    
    # Добавляем последнюю подгруппу
    if current_subgroup:
        subgroups.append(current_subgroup)
        print(f"[VISUAL]   Подгруппа #{len(subgroups)} закрыта: {current_length:.1f}м ({len(current_subgroup)} записей)")
    
    print(f"[VISUAL] ✓ Группа разбита на {len(subgroups)} подгрупп")
    return subgroups


def build_layout_sequence():
    """Формирует последовательность сегментов вдоль дорожки, РАЗДЕЛЁННУЮ ПО НАГРУЗКАМ."""
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    from core.reinforcement_db import get_reinforcement
    from pathlib import Path
    
    # Создаём глобальную карту армирования из PLATE_LOAD_DETAILS
    # Ключ включает load_code для точного определения армирования
    reinforcement_map = {}  # {(length, width_mm, load_code): reinforcement}
    db_path = Path(__file__).parent.parent / "pb.db"
    if cfg.PLATE_LOAD_DETAILS:
        print(f"[VISUAL] Начинаем создание карты армирования из {len(cfg.PLATE_LOAD_DETAILS)} записей")
        for key, qty in cfg.PLATE_LOAD_DETAILS.items():
            length, width_m, load_code = key[0], key[1], key[2]
            width_mm = int(round(width_m * 1000))
            reinforcement = get_reinforcement(
                length_m=length,
                load_code=load_code,
                source="series",
                db_path=db_path,
                allow_fallback=True,
            )
            # Канонический ключ: 12.5 и 13 → один и тот же ключ 13 (как в БД)
            lc_canonical = _canonical_load_code(load_code) or 8
            # В БД 12,5п = нагрузка 13; при "12п" приходит 12 — подставляем значение из 13
            if reinforcement is None and lc_canonical == 12:
                reinforcement = get_reinforcement(
                    length_m=length,
                    load_code=13,
                    source="series",
                    db_path=db_path,
                    allow_fallback=True,
                )
            key = (length, width_mm, lc_canonical)
            if reinforcement and reinforcement < 999:
                reinforcement_map[key] = reinforcement
                print(f"[VISUAL]   Добавлено: ({length}м, {width_mm}мм, нагрузка {lc_canonical}) → армирование {reinforcement:.1f}")
    
    # Дополняем карту по плитам из плана (primary_cuts), чтобы 71-12,5п и др. всегда имели армирование из БД
    def _supplement_reinforcement_map_from_plan(reinforcement_map, plan, db_path):
        added = 0
        for cut in plan.get("primary_cuts", []):
            length = cut["lengths"][0] if cut.get("lengths") else 6.0
            width_mm = cut["width"]
            load_code = cfg.normalize_load_code(cut.get("load_code", 8))
            lc_canonical = _canonical_load_code(load_code) or 8
            key = (length, width_mm, lc_canonical)
            if key not in reinforcement_map:
                reinforcement = get_reinforcement(
                    length_m=length,
                    load_code=load_code,
                    source="series",
                    db_path=db_path,
                    allow_fallback=True,
                )
                # В БД армирование для 12,5п хранится под нагрузкой 13; для "12п" в заказе приходит load_code=12 — подставляем значение из 13
                if reinforcement is None and lc_canonical == 12:
                    reinforcement = get_reinforcement(
                        length_m=length,
                        load_code=13,
                        source="series",
                        db_path=db_path,
                        allow_fallback=True,
                    )
                if reinforcement is not None and reinforcement < 999:
                    reinforcement_map[key] = reinforcement
                    added += 1
                    print(f"[VISUAL]   Дополнено из плана: ({length}м, {width_mm}мм, нагрузка {lc_canonical}) → армирование {reinforcement:.1f}")
                    # #region agent log
                    if abs(length - 7.1) < 0.05 and lc_canonical in (12, 13):
                        try:
                            open(_DEBUG_LOG, "a", encoding="utf-8").write(__import__("json").dumps({"hypothesisId": "H71supp", "location": "layout_sequence:_supplement", "message": "71-12п дополнение карты", "data": {"length": length, "width_mm": width_mm, "lc_canonical": lc_canonical, "reinforcement": reinforcement, "raw_load_code": load_code}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                    # #endregion
        return added
    if OPT_CASCADING_PLAN_BY_LOAD:
        for plan in OPT_CASCADING_PLAN_BY_LOAD.values():
            _supplement_reinforcement_map_from_plan(reinforcement_map, plan, db_path)
    elif OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get("primary_cuts"):
        _supplement_reinforcement_map_from_plan(reinforcement_map, OPT_CASCADING_PLAN, db_path)
    
    print(f"[VISUAL] Создана карта армирования: {len(reinforcement_map)} записей")
    sequence = []

    def plate_label(L: float, W: float, load_code: int = None) -> str:
        """
        Создает метку плиты с правильной нагрузкой.
        Если передан load_code — используется он; иначе ищется в PLATE_LOAD_DETAILS или get_load_code_for_plate.
        """
        resolved_load = load_code
        if resolved_load is None and cfg.PLATE_LOAD_DETAILS:
            for key, qty in cfg.PLATE_LOAD_DETAILS.items():
                plate_L, plate_W, plate_load = key[0], key[1], key[2]
                if abs(plate_L - L) < 0.05 and abs(plate_W - W) < 0.01:
                    resolved_load = plate_load
                    break
        if resolved_load is None:
            resolved_load = cfg.get_load_code_for_plate(L, W, default=(6 if W < 1.0 else 8))
        return cfg.make_plate_name(L, W, load_code=resolved_load)
    
    # ✅ НОВЫЙ ПРИОРИТЕТ 0: OPT_CASCADING_PLAN_BY_LOAD (группировка по нагрузкам)
    print(f"[VISUAL] Проверяем OPT_CASCADING_PLAN_BY_LOAD: {bool(OPT_CASCADING_PLAN_BY_LOAD)}")
    if OPT_CASCADING_PLAN_BY_LOAD:
        print(f"[VISUAL] ✅ Используем группировку по нагрузкам! Групп: {len(OPT_CASCADING_PLAN_BY_LOAD)}")
        # #region agent log (95694e) сколько 5.98/665 в планах по нагрузкам (вход в layout)
        try:
            _n_598665 = 0
            for _lg in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
                _p = OPT_CASCADING_PLAN_BY_LOAD[_lg]
                for _c in _p.get('primary_cuts', []):
                    _L = round(float((_c.get('lengths') or [6.0])[0]), 2)
                    _w = _c.get('width') or 1200
                    if abs(_L - 5.98) < 0.02 and _w == 665:
                        _n_598665 += _c.get('qty', 1)
            _log_p = _DEBUG_LOG_95694E
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_plan_598665", "location": "layout_sequence:build_layout_sequence:plans_in", "message": "count 5.98/665 in primary_cuts across all load plans", "data": {"count_598_665": _n_598665}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) сколько 5.08/320 и 5.98/530 в планах по нагрузкам (вход в layout)
        try:
            _n_508320 = _n_598530 = 0
            for _lg in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
                _p = OPT_CASCADING_PLAN_BY_LOAD[_lg]
                for _c in _p.get('primary_cuts', []):
                    _L = round(float((_c.get('lengths') or [6.0])[0]), 2)
                    _w = _c.get('width') or 1200
                    if abs(_L - 5.08) < 0.02 and _w == 320:
                        _n_508320 += _c.get('qty', 1)
                    if abs(_L - 5.98) < 0.02 and _w == 530:
                        _n_598530 += _c.get('qty', 1)
            _log_p = _DEBUG_LOG_95694E
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_plan_rescue", "location": "layout_sequence:build_layout_sequence:plans_in", "message": "count 5.08/320 and 5.98/530 in primary_cuts across all load plans", "data": {"count_508_320": _n_508320, "count_598_530": _n_598530}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        all_sequences = []
        
        for load_group in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
            plan = OPT_CASCADING_PLAN_BY_LOAD[load_group]
            
            # Получаем оригинальные нагрузки из плана (12, 12.5, и т.д.)
            original_loads = plan.get('original_loads', [load_group])
            
            # Форматируем метку с ОРИГИНАЛЬНЫМИ нагрузками
            load_display_list = [cfg.format_reinforcement_from_load_code(lc) for lc in original_loads]
            if len(load_display_list) > 1:
                load_display = ", ".join(load_display_list)
                label = f'Нагрузка {load_display}'
            else:
                label = f'Нагрузка {load_display_list[0]}'
            
            print(f"[VISUAL] Обрабатываем группу {load_group} ({label})...")
            
            # Строим последовательность для ЭТОЙ нагрузки (используем существующую логику)
            group_sequence = _build_sequence_from_plan(plan, plate_label, reinforcement_map)
            
            all_sequences.append({
                'load_code': load_group,  # Группа для совместимости
                'original_loads': original_loads,  # Оригинальные нагрузки
                'sequence': group_sequence,
                'label': label
            })
            
            print(f"[VISUAL]   → {len(group_sequence)} плит в группе")
        
        # #region agent log (95694e) сколько 5.98/665 в последовательностях после _build_sequence_from_plan
        try:
            _n_out = 0
            for _gr in all_sequences:
                for _it in _gr.get('sequence', []) or []:
                    _L = round(float(_it.get('length', 0) or _it.get('target_length', 0)), 2)
                    _w = _it.get('width') or _it.get('main_w') or 1.2
                    _w_mm = round(float(_w) * 1000) if float(_w) < 20 else round(float(_w))
                    if abs(_L - 5.98) < 0.02 and _w_mm == 665:
                        _n_out += 1
            _log_p = _DEBUG_LOG_95694E
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_layout_out_598665", "location": "layout_sequence:build_layout_sequence:plans_out", "message": "count 5.98/665 in sequences after _build_sequence_from_plan", "data": {"count_598_665": _n_out}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (95694e) сколько 5.08/320 и 5.98/530 в последовательностях после _build_sequence_from_plan
        try:
            _n508, _n598 = 0, 0
            for _gr in all_sequences:
                for _it in _gr.get('sequence', []) or []:
                    _L = round(float(_it.get('length', 0) or _it.get('target_length', 0)), 2)
                    _w = _it.get('width') or _it.get('main_w') or 1.2
                    _w_mm = round(float(_w) * 1000) if float(_w) < 20 else round(float(_w))
                    if abs(_L - 5.08) < 0.02 and _w_mm == 320:
                        _n508 += 1
                    if abs(_L - 5.98) < 0.02 and _w_mm == 530:
                        _n598 += 1
            _log_p = _DEBUG_LOG_95694E
            with open(_log_p, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "95694e", "hypothesisId": "H_95694e_layout_out_rescue", "location": "layout_sequence:build_layout_sequence:plans_out", "message": "count 5.08/320 and 5.98/530 in sequences after _build_sequence_from_plan", "data": {"count_508_320": _n508, "count_598_530": _n598}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log (2d5c43) H3 grouped path: sequence totals and by key before return
        try:
            _log_2d5c43 = _DEBUG_LOG_2D5C43
            _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
            _seq_by_key = {tuple(tk): 0 for tk in _target_keys}
            _total_in_sequence = 0
            for _gr in all_sequences:
                for s in _gr.get('sequence', []) or []:
                    _total_in_sequence += 1
                    L = round(float(s.get('length', 0) or s.get('target_length', 0)), 2)
                    w = s.get('width') or s.get('main_w') or 1.2
                    w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                    lc = s.get('load_code', 8)
                    try:
                        lc = int(lc) if lc is not None else 8
                    except (TypeError, ValueError):
                        lc = 8
                    for tk in _target_keys:
                        if abs(L - tk[0]) <= 0.02 and w_mm == tk[1] and lc == tk[2]:
                            _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                            break
                    for sec in s.get('secondary_cuts', []):
                        sw = sec.get('width', 0)
                        sw_mm = round(float(sw) * 1000) if float(sw) < 20 else round(float(sw))
                        sl = round(float(sec.get('target_length') or L), 2)
                        for tk in _target_keys:
                            if abs(sl - tk[0]) <= 0.02 and sw_mm == tk[1]:
                                _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                                break
            _prim_total = sum(len(p.get('primary_cuts', [])) for p in OPT_CASCADING_PLAN_BY_LOAD.values())
            _seq_by_key_ser = [list(k) + [v] for k, v in _seq_by_key.items()]
            with open(_log_2d5c43, 'a', encoding='utf-8') as _f:
                _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H3", "location": "layout_sequence:build_layout_sequence:grouped_return", "message": "grouped path: sequence total and by key vs primary_cuts total", "data": {"total_from_primary": _prim_total, "total_in_sequence": _total_in_sequence, "sequence_by_key": _seq_by_key_ser}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # НОВЫЙ ФОРМАТ ВОЗВРАТА: список групп по нагрузкам
        return all_sequences
    
    # Приоритет 1: OPT_CASCADING_PLAN (старый формат, без группировки)
    print(f"[VISUAL] Проверяем OPT_CASCADING_PLAN: {OPT_CASCADING_PLAN is not None}")
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):
        sequence = _build_sequence_from_plan(OPT_CASCADING_PLAN, plate_label, reinforcement_map)
        _ensure_sequence_layout_uid(sequence, prefix="single")
        return sequence
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('primary_cuts'):
        print("[VISUAL] OK: Используем каскадную оптимизацию для визуализации")
        print(f"[VISUAL] Первичных резов: {len(OPT_CASCADING_PLAN.get('primary_cuts', []))}")
        print(f"[VISUAL] Вторичных резов: {len(OPT_CASCADING_PLAN.get('secondary_cuts', []))}")
        
        # Проверяем, есть ли 2D данные (plate_assignments)
        use_2d_data = 'plate_assignments' in OPT_CASCADING_PLAN and OPT_CASCADING_PLAN['plate_assignments']
        
        if use_2d_data:
            print("[VISUAL] ТОЧНО: Используем 2D данные с точными длинами")
        else:
            print("[VISUAL] ВНИМАНИЕ: 2D данных нет, используем приближение")
            
            # Собираем все плиты с их длинами из cfg
            all_plates_with_lengths = []
            for plates, width_mm in [
                (cfg.PLATES_1_2, 1200), (cfg.PLATES_1_08, 1080),
                (cfg.PLATES_0_32, 320), (cfg.PLATES_0_46, 460), (cfg.PLATES_0_70, 700),
                (cfg.PLATES_0_72, 720), (cfg.PLATES_0_86, 860), (cfg.PLATES_0_88, 880),
                (cfg.PLATES_0_74, 740), (cfg.PLATES_0_48, 480), (cfg.PLATES_0_50, 500),
                (cfg.PLATES_0_34, 340)
            ]:
                for length in plates:
                    all_plates_with_lengths.append({'length': length, 'width': width_mm})
            
            # Сортируем по ширине для соответствия с оптимизацией
            all_plates_with_lengths.sort(key=lambda x: (-x['width'], -x['length']))
        
        # Создаём карту поперечных резов: {(length, width): {target_length, remainder}}
        transverse_cut_map = {}
        if OPT_CASCADING_PLAN.get('transverse_cuts'):
            for tcut in OPT_CASCADING_PLAN['transverse_cuts']:
                key = (tcut['source_length'], tcut['source_width'])
                transverse_cut_map[key] = {
                    'target_length': tcut['target_length'],
                    'remainder': tcut['remainder']
                }
            print(f"[VISUAL] Найдено {len(transverse_cut_map)} типов поперечных резов: {list(transverse_cut_map.keys())}")
        
        # СТАРАЯ ЛОГИКА (если нет plate_assignments_with_transverse)
        # Создаём карту вторичных резов: {(source_length, остаток_мм): [ {pattern, qty, used}, ... ]}
        secondary_cuts_info = {}
        if OPT_CASCADING_PLAN.get('secondary_cuts'):
            for sec_cut in OPT_CASCADING_PLAN['secondary_cuts']:
                source_mm = sec_cut['source']
                pieces = sec_cut.get('pieces', 1)
                cuts_list = sec_cut.get('cuts', [])
                qty = sec_cut['qty']  # Сколько остатков режется вторично
                
                # ВАЖНО: Получаем ИСХОДНЫЕ длины остатков (ДО поперечного реза!)
                source_lengths_list = sec_cut.get('source_lengths', [])
                # Результирующие длины (ПОСЛЕ поперечного реза)
                target_lengths_list = sec_cut.get('lengths', [])
                # Целевой load_code заказа (для корректного учёта в дорожках и РЕСКЬЮ)
                target_order_key = sec_cut.get('target_order_key')
                target_load_code = cfg.normalize_load_code(target_order_key[2]) if (target_order_key and len(target_order_key) > 2) else None
                
                # Создаём шаблон вторичных резов для ОДНОГО остатка
                pattern = []
                if cuts_list:
                    target_width_mm = cuts_list[0]
                    # Для множественной резки (pieces >= 2) создаём несколько сегментов
                    # Для сужения (pieces == 1) создаём один сегмент
                    for _ in range(pieces):
                        pattern.append({
                            'width': target_width_mm / 1000.0,
                            'width_mm': target_width_mm,  # Ширина РЕЗУЛЬТАТА вторичного реза
                            'source_width_mm': source_mm,  # Ширина ОСТАТКА (для правильной метки)
                            'label': None,  # Метка будет создана позже с реальной длиной плиты
                            'target_length': target_lengths_list[0] if target_lengths_list else None,  # Для поперечных резов
                            'target_load_code': target_load_code
                        })
                
                # ИСПРАВЛЕНИЕ: Создаём запись для КАЖДОЙ ИСХОДНОЙ длины отдельно
                for i in range(qty):
                    source_length = source_lengths_list[i] if i < len(source_lengths_list) else 6.0
                    key = _secondary_geom_cut_key(source_length, source_mm)

                    if key not in secondary_cuts_info:
                        secondary_cuts_info[key] = []

                    secondary_cuts_info[key].append({
                        'pattern': [segment.copy() for segment in pattern],
                        'qty': 1,
                        'used': 0,
                        # BUG-4 FIX: сохраняем canonical ключ заказа для точного match в rescue
                        'target_order_key': target_order_key,
                    })
        
        print(f"[VISUAL] Создано {len(secondary_cuts_info)} вариантов вторичных резов:")
        for (src_len, src_w), variants in secondary_cuts_info.items():
            for idx, info in enumerate(variants, start=1):
                pattern_desc = ", ".join([f"{c['width_mm']}мм" for c in info['pattern']])
                print(f"  Остаток {src_len}м x {src_w}мм: вариант #{idx} -> [{pattern_desc}]")
        
        # ========== НОВАЯ ЛОГИКА: РАЗДЕЛИТЕЛИ МЕЖДУ ГРУППАМИ РЕЗОВ ==========
        # Требования завода (ОБНОВЛЁННЫЕ):
        # 1. Первая плита ДОЛЖНА быть целой с МИНИМАЛЬНЫМ армированием
        # 2. Плиты с одинаковым резом И одинаковым армированием идут подряд
        # 3. Группы сортируются по армированию (от меньшего к большему)
        # 4. Между группами с РАЗНЫМ резом/армированием должна быть целая плита-разделитель с мин. армированием
        all_primary_cuts = OPT_CASCADING_PLAN.get('primary_cuts', [])
        # ИСПРАВЛЕНО: Целая плита = любая плита без реза (rest=0)
        # Раньше фильтровали только 1200мм, из-за чего плиты 1080мм и другие терялись!
        solid_cuts = [cut for cut in all_primary_cuts 
                      if cut['rest'] == 0]
        
        # НОВОЕ: Вычисляем армирование для каждой целой плиты и сортируем по армированию (мин. первый)
        for cut in solid_cuts:
            length = cut['lengths'][0] if cut.get('lengths') else 6.0
            width_mm = cut['width']
            lc = cfg.normalize_load_code(cut.get('load_code', 8))
            cut['reinforcement'] = _get_reinforcement_from_map(reinforcement_map, length, width_mm, lc) or 999.0
        solid_cuts.sort(key=lambda x: (x.get('reinforcement', 999.0), -x['lengths'][0] if x.get('lengths') else 0))
        
        # Плиты с резом - вычисляем армирование для группировки
        cut_with_rest_raw = [cut for cut in all_primary_cuts if cut['rest'] > 0]
        
        # НОВОЕ: Добавляем армирование к каждой плите с резом для группировки
        for cut in cut_with_rest_raw:
            length = cut['lengths'][0] if cut.get('lengths') else 6.0
            width_mm = cut['width']
            lc = cfg.normalize_load_code(cut.get('load_code', 8))
            cut['reinforcement'] = _get_reinforcement_from_map(reinforcement_map, length, width_mm, lc) or 999.0
        
        # Сортируем плиты с резом по армированию (мин. первый), потом по типу реза
        cut_with_rest = sorted(
            cut_with_rest_raw,
            key=lambda x: (x.get('reinforcement', 999.0), x['width'], x['rest'])
        )
        
        print(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
        if solid_cuts:
            print(f"[VISUAL] Целые плиты (сортировка по армированию): {[(c['width'], c['qty'], c.get('reinforcement', '?')) for c in solid_cuts[:5]]}")
        
        # ✅ НОВОЕ: Логируем все плиты из primary_cuts
        import logging
        logger = logging.getLogger(__name__)
        logger.info(f"[TRACE] ===== ШАГ 3: ПЛИТЫ ИЗ ОПТИМИЗАЦИИ (primary_cuts) =====")
        total_from_primary = sum(c['qty'] for c in all_primary_cuts)
        logger.info(f"[TRACE] Всего записей primary_cuts: {len(all_primary_cuts)}")
        logger.info(f"[TRACE] Всего плит: {total_from_primary}")
        
        for i, cut in enumerate(all_primary_cuts):
            lengths_str = ', '.join([f"{l:.2f}" for l in cut.get('lengths', [])[:3]])
            if len(cut.get('lengths', [])) > 3:
                lengths_str += f", ... ({len(cut['lengths'])} шт)"
            logger.info(f"[TRACE]   #{i+1}: width={cut['width']}мм, rest={cut['rest']}мм, qty={cut['qty']}, lengths=[{lengths_str}], kp_id={cut.get('kp_id', '?')}")
        
        # НОВОЕ: Группируем плиты с резом по (width, rest, reinforcement)
        # Это гарантирует что плиты с одинаковым резом И армированием идут вместе
        from itertools import groupby
        cut_groups = [list(group) for key, group in groupby(cut_with_rest, key=lambda x: (x['width'], x['rest'], x.get('reinforcement', 999.0)))]
        
        if cut_groups:
            print(f"[VISUAL] Найдено {len(cut_groups)} групп резов (сгруппировано по рез+армирование):")
            for i, group in enumerate(cut_groups, 1):
                print(f"[VISUAL]   Группа {i}: width={group[0]['width']}мм, rest={group[0]['rest']}мм, армирование={group[0].get('reinforcement', '?'):.1f}, плит={sum(c['qty'] for c in group)}")
        
        # Формируем последовательность с разделителями
        ordered_cuts = []
        
        # ВАЖНО: Разворачиваем записи целых плит в отдельные плиты
        # Из записи {qty: 5, lengths: [6.15, 6.15, ...]} создаём 5 записей {qty: 1, lengths: [6.15]}
        solid_cuts_list = []
        for cut in solid_cuts:
            lengths = cut.get('lengths', [])
            for i in range(cut['qty']):
                single_cut = cut.copy()
                single_cut['qty'] = 1  # Каждая запись = 1 плита
                single_cut['lengths'] = [lengths[i]] if i < len(lengths) else [lengths[0] if lengths else 6.0]
                solid_cuts_list.append(single_cut)
        
        print(f"[VISUAL] Развёрнуто {len(solid_cuts_list)} отдельных целых плит для разделителей")
        
        # Правило 1: Первая плита ОБЯЗАТЕЛЬНО целая
        if solid_cuts_list:
            first_plate = solid_cuts_list.pop(0)
            ordered_cuts.append(first_plate)
            first_width = first_plate.get('width', 1200)
            print(f"[VISUAL] ✓ Первая плита: целая {first_width}мм")
        
        # Правило 2 и 3: Чередуем группы резов и целые плиты-разделители
        for i, cut_group in enumerate(cut_groups):
            # Рассчитываем суммарную длину группы
            total_group_length = sum(
                cut['lengths'][0] * cut['qty'] 
                for cut in cut_group
                if cut.get('lengths')
            )
            
            print(f"[VISUAL] Группа резов #{i+1}: width={cut_group[0]['width']}мм, rest={cut_group[0]['rest']}мм, "
                  f"типов={len(cut_group)}, длина={total_group_length:.1f}м")
            
            # Если группа слишком большая (>90м), разбиваем на подгруппы
            if total_group_length > 90.0:
                print(f"[VISUAL] ⚠️ Группа #{i+1} слишком большая ({total_group_length:.1f}м > 90м), разбиваем на подгруппы")
                subgroups = _split_group_into_subgroups(cut_group, max_length=90.0)
                
                # Добавляем подгруппы с разделителями МЕЖДУ ними
                for j, subgroup in enumerate(subgroups):
                    ordered_cuts.extend(subgroup)
                    print(f"[VISUAL]   Добавлена подгруппа #{i+1}.{j+1}: {len(subgroup)} плит")
                    
                    # После каждой подгруппы (кроме последней в этой группе)
                    if j < len(subgroups) - 1 and solid_cuts_list:
                        separator = solid_cuts_list.pop(0)
                        separator['is_separator'] = True
                        ordered_cuts.append(separator)
                        print(f"[VISUAL]   ✓ Разделитель между подгруппами: целая плита")
            else:
                # Группа помещается в 1 дорожку - добавляем как есть
                ordered_cuts.extend(cut_group)
                print(f"[VISUAL] Добавлена группа резов #{i+1} (влезает в дорожку)")
            
            # После каждой ГРУППЫ (кроме последней) добавляем ОПТИМАЛЬНЫЙ разделитель
            if i < len(cut_groups) - 1 and solid_cuts_list:
                # Определяем следующую группу
                next_group = cut_groups[i + 1]
                
                # Умный выбор разделителя по армированию
                best_idx = _choose_best_separator(solid_cuts_list, next_group, reinforcement_map)
                
                if best_idx is not None:
                    # Извлекаем выбранную плиту
                    separator = solid_cuts_list.pop(best_idx)
                    separator['is_separator'] = True  # МЯГКОЕ РЕЗЕРВИРОВАНИЕ: помечаем как разделитель
                    ordered_cuts.append(separator)
                    print(f"[VISUAL] ✓ Разделитель (is_separator=True): целая плита между группами")
                else:
                    # Fallback: если функция не смогла выбрать, берём первую
                    if solid_cuts_list:
                        fallback_sep = solid_cuts_list.pop(0)
                        fallback_sep['is_separator'] = True  # МЯГКОЕ РЕЗЕРВИРОВАНИЕ
                        ordered_cuts.append(fallback_sep)
                        print(f"[VISUAL] ✓ Разделитель: целая плита между группами (fallback, is_separator=True)")
        
        # Оставшиеся целые плиты добавляем в конец
        if solid_cuts_list:
            ordered_cuts.extend(solid_cuts_list)
            print(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")
        
        # ✅ НОВОЕ: Логируем ordered_cuts (плиты после группировки)
        logger.info(f"[TRACE] ===== ШАГ 4: ПЛИТЫ ПОСЛЕ ГРУППИРОВКИ (ordered_cuts) =====")
        total_ordered = sum(c['qty'] for c in ordered_cuts)
        logger.info(f"[TRACE] Всего записей: {len(ordered_cuts)}")
        logger.info(f"[TRACE] Всего плит: {total_ordered}")
        
        # === ПРОВЕРКА ПОСЛЕ ГРУППИРОВКИ ===
        if total_from_primary != total_ordered:
            logger.warning(f"[WARNING] Потеря плит на этапе группировки!")
            logger.warning(f"[WARNING]   До группировки: {total_from_primary}")
            logger.warning(f"[WARNING]   После группировки: {total_ordered}")
            logger.warning(f"[WARNING]   Потеряно: {total_from_primary - total_ordered}")
        
        for i, cut in enumerate(ordered_cuts):
            is_sep = " [РАЗДЕЛИТЕЛЬ]" if cut.get('is_separator') else ""
            logger.info(f"[TRACE]   #{i+1}: width={cut['width']}мм, rest={cut['rest']}мм, qty={cut['qty']}, reinf={cut.get('reinforcement', '?'):.1f}{is_sep}")
        
        # Диагностика: предупреждение при отсутствии армирования в карте (один раз на ключ)
        _missing_reinf_logged = set()
        import logging as _logging
        _log_build = _logging.getLogger(__name__)
        def _warn_missing_reinforcement(length, width_mm, load_code_from_cut):
            if load_code_from_cut is None:
                return
            key = (round(length, 3), width_mm, load_code_from_cut)
            if key in _missing_reinf_logged:
                return
            _missing_reinf_logged.add(key)
            _log_build.warning(
                "[ВИЗУАЛИЗАЦИЯ] Нет армирования в карте для (length=%s, width_mm=%s, load_code=%s)",
                length, width_mm, load_code_from_cut,
            )
        # 1. Первичные резы с вторичными резами внутри остатков
        # ОБРАБАТЫВАЕМ В НОВОМ ПОРЯДКЕ: целая → группа резов → целая → группа резов
        for cut in ordered_cuts:
            width_mm = cut['width']
            rest_mm = cut['rest']
            qty = cut['qty']
            
            # Получаем длины для этих плит
            if use_2d_data and 'lengths' in cut:
                # Используем точные длины из 2D оптимизации
                lengths_for_cut = cut['lengths']
                print(f"[VISUAL] Первичный рез {width_mm}мм: используем точные длины {lengths_for_cut}")
            else:
                # Используем приближение из all_plates_with_lengths
                matching_plates = [p for p in all_plates_with_lengths if p['width'] == width_mm]
                lengths_for_cut = [p['length'] for p in matching_plates[:qty]]
                # Если не хватает, дополняем средней длиной
                while len(lengths_for_cut) < qty:
                    lengths_for_cut.append(6.0 if not matching_plates else matching_plates[0]['length'])
            
            for i in range(qty):
                # Берём длину для этой плиты
                length = lengths_for_cut[i] if i < len(lengths_for_cut) else 6.0
                
                # НОВОЕ: Получаем информацию о КП для этой плиты
                kp_id = cut.get('kp_id')
                customer = cut.get('customer')
                kp_date = cut.get('kp_date')
                plate_name_from_cut = cut.get('plate_name')
                
                # ИСПРАВЛЕНИЕ: Проверяем вторичные резы для КОНКРЕТНОГО остатка (длина + ширина)
                sec_variants = secondary_cuts_info.get(_secondary_geom_cut_key(length, rest_mm)) or []
                
                # ОТЛАДКА: Выводим информацию о поиске вторичных резов
                if rest_mm > 0:
                    found_variant = any(variant['used'] < variant['qty'] for variant in sec_variants)
                    print(f"[VISUAL] Ищем вторичные резы для остатка {length}м x {rest_mm}мм: {'НАЙДЕНО' if found_variant else 'НЕ НАЙДЕНО'}")
                
                # Проверяем, есть ли поперечный рез для этой плиты
                transverse_cut_info = transverse_cut_map.get((length, width_mm))
                
                if transverse_cut_info:
                    # Эта плита режется поперёк - добавляем с mode='transverse'
                    width_m = width_mm / 1000.0
                    target_length = transverse_cut_info['target_length']
                    remainder = transverse_cut_info['remainder']
                    
                    # Получаем армирование из карты (с учётом load_code)
                    load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 8))
                    reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    sequence.append({
                        'length': length,  # Исходная длина плиты
                        'mode': 'transverse',
                        'target_length': target_length,
                        'remainder': remainder,
                        'width': width_m,
                        'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code
                        'label_target': plate_label(target_length, width_m, load_code_from_cut),
                        'label_remainder': f'Остаток {remainder:.2f}м'.replace('.', ',') if remainder > 0.1 else '',
                        'reinforcement': reinforcement,
                        'kp_id': kp_id,
                        'customer': customer,
                        'kp_date': kp_date,
                        'plate_name': plate_name_from_cut
                    })
                    print(f"[VISUAL] Плита с поперечным резом: {length}м x {width_mm}мм -> {target_length}м (остаток {remainder:.2f}м)")
                else:
                    # Обычная плита без поперечного реза
                    main_w = width_mm / 1000.0
                    rest_w = rest_mm / 1000.0
                    fake_rest_override = False
                    
                    # ВАЖНО: Плиты 1.08 м возникают ТОЛЬКО продольным резом из 1.2 м.
                    # Даже если оптимизатор пометил их как "solid" (rest=0),
                    # для человека нужно показать линию реза и остаток 0.12 м.
                    if width_mm == 1080 and rest_mm == 0:
                        rest_mm = 120
                        rest_w = 0.12
                        fake_rest_override = True
                    
                    # Специальная обработка для плит БЕЗ реза (rest = 0)
                    if rest_mm == 0:
                        # Получаем армирование из карты (с учётом load_code)
                        load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 8))
                        reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                        if reinforcement is None:
                            _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                        # #region agent log
                        if abs(length - 7.1) < 0.05 and load_code_from_cut in (12, 12.5, 13):
                            try:
                                open(_DEBUG_LOG, "a", encoding="utf-8").write(__import__("json").dumps({"hypothesisId": "H71item", "location": "layout_sequence:sequence_append_solid", "message": "71-12п item в sequence", "data": {"length": length, "width_mm": width_mm, "load_code_from_cut": load_code_from_cut, "reinforcement": reinforcement}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                            except Exception:
                                pass
                        # #endregion
                        # МЯГКОЕ РЕЗЕРВИРОВАНИЕ: передаём флаг is_separator
                        is_separator = cut.get('is_separator', False)
                        sequence.append({
                            'length': length,
                            'mode': 'solid',
                            'width': main_w,  # ИСПРАВЛЕНИЕ: сохраняем реальную ширину для корректного учёта
                            'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code для различения плит
                            'label': plate_label(length, main_w, load_code_from_cut),
                            'reinforcement': reinforcement,
                            'is_separator': is_separator,  # Для приоритета при разбиении на дорожки
                            'kp_id': kp_id,
                            'customer': customer,
                            'kp_date': kp_date,
                            'plate_name': plate_name_from_cut
                        })
                    else:
                        # Плиты С резом
                        load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 8))
                        # Проверяем, нужны ли вторичные резы для ЭТОЙ плиты
                        secondary_cuts_for_plate = None
                        chosen_variant = None
                        for variant in sec_variants:
                            if variant['used'] < variant['qty']:
                                chosen_variant = variant
                                break

                        if chosen_variant:
                            secondary_cuts_for_plate = []
                            # BUG-4 FIX: сохраняем target_order_key из варианта для точного
                            # match в _count_tracks_for_rescue без fuzzy-поиска
                            _sec_tok = chosen_variant.get('target_order_key')
                            for sec_cut_template in chosen_variant['pattern']:
                                sec_width = sec_cut_template['width']
                                sec_width_mm = sec_cut_template['width_mm']
                                lc = sec_cut_template.get('target_load_code')
                                lc = cfg.normalize_load_code(lc) if lc is not None else load_code_from_cut
                                # ВАЖНО: Проверяем поперечные резы для ВТОРИЧНЫХ плит!
                                sec_transverse = transverse_cut_map.get((length, sec_width_mm))
                                
                                if sec_transverse:
                                    # Вторичная плита с поперечным резом
                                    secondary_cuts_for_plate.append({
                                        'width': sec_width,
                                        'label': f'[2] {plate_label(sec_transverse["target_length"], sec_width, lc)}',
                                        'transverse_cut': True,
                                        'target_length': sec_transverse['target_length'],
                                        'remainder': sec_transverse['remainder'],
                                        'load_code': lc,
                                        'target_order_key': _sec_tok,
                                    })
                                    print(f"[VISUAL] Вторичный рез С поперечным: {length}м x {sec_width_mm}мм -> {sec_transverse['target_length']}м")
                                else:
                                    # Вторичная плита (может быть с поперечным резом!)
                                    # Проверяем, есть ли target_length в шаблоне (для transverse-резов)
                                    target_length = sec_cut_template.get('target_length')
                                    
                                    if target_length:
                                        # Это вторичный рез типа 'transverse' (поперечный + продольный)
                                        # Метка показывает результат ОБОИХ резов
                                        secondary_cuts_for_plate.append({
                                            'width': sec_width,
                                            'label': f'О {plate_label(target_length, sec_width, lc)}',  # О = Остаток
                                            'has_transverse': True,  # Флаг для отрисовки красной линии
                                            'target_length': target_length,  # Длина результата (для правильной отрисовки)
                                            'load_code': lc,
                                            'target_order_key': _sec_tok,
                                        })
                                    else:
                                        # Обычный вторичный рез (включая narrowing)
                                        result_width = sec_cut_template['width']
                                        source_width = sec_cut_template.get('source_width_mm', result_width * 1000) / 1000.0
                                        label_text = plate_label(length, result_width, lc)
                                        if abs(result_width - source_width) > 1e-6:
                                            label_text = f'О {label_text}'  # помечаем, что получено из остатка
                                        secondary_cuts_for_plate.append({
                                            'width': result_width,
                                            'label': label_text,
                                            'load_code': lc,
                                            'target_order_key': _sec_tok,
                                        })
                            chosen_variant['used'] += 1
                        
                        # Получаем армирование из карты (с учётом load_code)
                        reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                        if reinforcement is None:
                            _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                        sequence.append({
                            'length': length,
                            'mode': 'split',
                            'main_w': main_w,
                            'rest_w': rest_w,
                            'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code
                            'label_main': plate_label(length, main_w, load_code_from_cut),
                            'label_rest': (
                                '+0,12' if fake_rest_override else
                                (f'+{rest_w:.2f}'.replace('.', ',') if not secondary_cuts_for_plate else None)
                            ),
                            'secondary_cuts': secondary_cuts_for_plate,
                            'reinforcement': reinforcement,
                            'kp_id': kp_id,
                            'customer': customer,
                            'kp_date': kp_date,
                            'plate_name': plate_name_from_cut
                        })
        
        if sequence:
            # ✅ НОВОЕ: Логируем финальную последовательность
            logger.info(f"[TRACE] ===== ШАГ 5: ФИНАЛЬНАЯ ПОСЛЕДОВАТЕЛЬНОСТЬ (sequence) =====")
            logger.info(f"[TRACE] Всего плит в sequence: {len(sequence)}")
            
            # Группируем по типам
            solid_count = sum(1 for s in sequence if s.get('mode') == 'solid')
            split_count = sum(1 for s in sequence if s.get('mode') == 'split')
            transverse_count = sum(1 for s in sequence if s.get('mode') == 'transverse')
            
            logger.info(f"[TRACE] Плит solid (без реза): {solid_count}")
            logger.info(f"[TRACE] Плит split (с резом): {split_count}")
            logger.info(f"[TRACE] Плит transverse (поперечный рез): {transverse_count}")
            
            # Подсчитываем плиты из вторичных резов
            secondary_count = 0
            for s in sequence:
                if s.get('mode') == 'split' and s.get('secondary_cuts'):
                    secondary_count += len(s['secondary_cuts'])
            logger.info(f"[TRACE] Плит из вторичных резов: {secondary_count}")
            
            total_in_sequence = len(sequence) + secondary_count
            logger.info(f"[TRACE] ИТОГО плит: {total_in_sequence}")
            
            # === КРИТИЧЕСКАЯ ПРОВЕРКА НА ПОТЕРЮ ПЛИТ ===
            if total_from_primary != total_in_sequence:
                logger.error(f"[CRITICAL] ПОТЕРЯ ПЛИТ ОБНАРУЖЕНА!")
                logger.error(f"[CRITICAL]   Запрошено из оптимизации: {total_from_primary}")
                logger.error(f"[CRITICAL]   Получено в sequence:      {total_in_sequence}")
                logger.error(f"[CRITICAL]   ПОТЕРЯНО: {total_from_primary - total_in_sequence} плит(ы)")
                
                # Детальный анализ потерь по ширинам
                requested_by_width = {}
                for cut in all_primary_cuts:
                    w = cut['width']
                    requested_by_width[w] = requested_by_width.get(w, 0) + cut['qty']
                
                result_by_width = {}
                for s in sequence:
                    w = s.get('width', s.get('main_w', 1.2) * 1000 if 'main_w' in s else 1200)
                    result_by_width[w] = result_by_width.get(w, 0) + 1
                    # Добавляем вторичные резы
                    for sec in s.get('secondary_cuts', []):
                        sec_w = sec.get('width', 0)
                        result_by_width[sec_w] = result_by_width.get(sec_w, 0) + 1
                
                logger.error(f"[CRITICAL] Сравнение по ширинам:")
                all_widths = set(requested_by_width.keys()) | set(result_by_width.keys())
                for w in sorted(all_widths):
                    req = requested_by_width.get(w, 0)
                    res = result_by_width.get(w, 0)
                    diff = req - res
                    if diff != 0:
                        logger.error(f"[CRITICAL]   Ширина {w}мм: запрошено {req}, получено {res}, ПОТЕРЯ: {diff}")
            else:
                logger.info(f"[TRACE] ✓ Проверка пройдена: все {total_from_primary} плит в sequence")
            # #region agent log (2d5c43) H3: sequence total vs primary, counts by key
            try:
                _log_2d5c43 = _DEBUG_LOG_2D5C43
                _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
                _seq_by_key = {tuple(tk): 0 for tk in _target_keys}
                _seq_6_530_1200 = []
                for s in sequence:
                    L = round(float(s.get('length', 0) or s.get('target_length', 0)), 2)
                    w = s.get('width') or s.get('main_w') or 1.2
                    w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                    lc = s.get('load_code', 8)
                    try:
                        lc = int(lc) if lc is not None else 8
                    except (TypeError, ValueError):
                        lc = 8
                    for tk in _target_keys:
                        if abs(L - tk[0]) <= 0.02 and w_mm == tk[1] and lc == tk[2]:
                            _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                            break
                    if 5.98 <= L <= 6.02 and w_mm in (530, 1200) and len(_seq_6_530_1200) < 25:
                        _seq_6_530_1200.append({"length": L, "width_mm": w_mm, "mode": s.get('mode'), "label": (s.get('label') or '')[:50]})
                    for sec in s.get('secondary_cuts', []):
                        sw = sec.get('width', 0)
                        sw_mm = round(float(sw) * 1000) if float(sw) < 20 else round(float(sw))
                        sl = round(float(sec.get('target_length') or L), 2)
                        for tk in _target_keys:
                            if abs(sl - tk[0]) <= 0.02 and sw_mm == tk[1]:
                                _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                                break
                with open(_log_2d5c43, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "2d5c43", "hypothesisId": "H3", "location": "layout_sequence:build_layout_sequence:before_return_sequence", "message": "sequence vs primary totals and by key", "data": {"total_from_primary": total_from_primary, "total_in_sequence": len(sequence), "sequence_by_key": dict(_seq_by_key), "sequence_6m_530_1200_sample": _seq_6_530_1200}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
            return sequence
    else:
        print("[VISUAL] ВНИМАНИЕ: OPT_CASCADING_PLAN не найден или пуст, используем старый метод")
    
    # Приоритет 2: OPT_PLAN (старая оптимизация)
    if OPT_PLAN and OPT_PLAN.get('actions'):
        for act in OPT_PLAN['actions']:
            src_type, W1, W2, L, qty, lc, tc = act
            W1_m = W1 / 1000.0; W2_m = W2 / 1000.0 if W2 else 0
            for _ in range(qty):
                if src_type == 'solid':
                    sequence.append({'length': L, 'mode': 'solid', 'label': plate_label(L, W1_m)})
                elif src_type == 'split':
                    rest_w = W2_m if W2_m < W1_m else (1.2 - W1_m)
                    rest_label = f'+{rest_w:.2f}'.replace('.', ',')
                    sequence.append({'length': L, 'mode': 'split', 'main_w': W1_m, 'rest_w': rest_w,
                                     'label_main': plate_label(L, W1_m), 'label_rest': rest_label})
                elif src_type == 'narrow':
                    delta = abs(W2_m - W1_m) if W2_m else 0
                    rest_label = f'-{delta:.2f}'.replace('.', ',') if delta > 0.001 else ''
                    sequence.append({'length': L, 'mode': 'split', 'main_w': W1_m, 'rest_w': delta,
                                     'label_main': plate_label(L, W1_m), 'label_rest': rest_label})
        return sequence
    
    # Приоритет 3: Fallback на старую логику
    for L in cfg.PLATES_1_2:
        sequence.append({'length': L, 'mode': 'solid', 'label': plate_label(L, 1.2)})
    for L in cfg.PLATES_1_5_TO_1_2:
        sequence.append({'length': L, 'mode': 'solid', 'label': plate_label(L, 1.2)})
    for L in cfg.PLATES_1_0:
        sequence.append({'length': L, 'mode': 'split', 'main_w': 1.0, 'rest_w': 0.2,
                         'label_main': plate_label(L, 1.0), 'label_rest': '+0,2'})
    for L in cfg.PLATES_1_08:
        sequence.append({'length': L, 'mode': 'split', 'main_w': 1.08, 'rest_w': 0.12,
                         'label_main': plate_label(L, 1.08), 'label_rest': '+0,12'})
    
    groups_map = {
        '0_32': (cfg.PLATES_0_32, 0.32, 0.88, '+0,88'),
        '0_46': (cfg.PLATES_0_46, 0.46, 0.74, '+0,74'),
        '0_70': (cfg.PLATES_0_70, 0.70, 0.50, '+0,50'),
        '0_72': (cfg.PLATES_0_72, 0.72, 0.48, '+0,48'),
        '0_86': (cfg.PLATES_0_86, 0.86, 0.34, '+0,34'),
    }
    if len(cfg.PLATES_0_74):
        groups_map['0_74'] = (cfg.PLATES_0_74, 0.74, 0.46, '+0,46')
    if len(cfg.PLATES_0_88):
        groups_map['0_88'] = (cfg.PLATES_0_88, 0.88, 0.32, '+0,32')
    if len(cfg.PLATES_0_48):
        groups_map['0_48'] = (cfg.PLATES_0_48, 0.48, 0.72, '+0,72')
    if len(cfg.PLATES_0_50):
        groups_map['0_50'] = (cfg.PLATES_0_50, 0.50, 0.70, '+0,70')
    if len(cfg.PLATES_0_34):
        groups_map['0_34'] = (cfg.PLATES_0_34, 0.34, 0.86, '+0,86')
    
    order = OPT_WIDTH_PRIORITY or list(groups_map.keys())
    for key in order:
        if key not in groups_map:
            continue
        items, main_w, rest_w, rest_label = groups_map[key]
        for L in items:
            sequence.append({'length': L, 'mode': 'split', 'main_w': main_w, 'rest_w': rest_w,
                             'label_main': plate_label(L, main_w), 'label_rest': rest_label})

    return sequence


def _secondary_geom_cut_key(length_m: object, rest_or_source_mm: object) -> tuple[float, int]:
    """Ключ (длина м, остаток мм) для secondary_cuts_info / legacy-lookup.

    Нормализует float/int чтобы (2.7, 880) и (2.7, 880.0) не давали промах dict.
    """
    return (round(float(length_m), 2), int(round(float(rest_or_source_mm))))


def _build_sequence_from_plan(plan, plate_label_func, reinforcement_map=None):
    """
    Вспомогательная функция: строит последовательность плит из плана оптимизации.
    
    Args:
        plan: Результат оптимизации (OPT_CASCADING_PLAN)
        plate_label_func: Функция для создания меток плит
        reinforcement_map: Словарь {(length, width_mm, load_code): reinforcement} для получения армирования
    
    Returns:
        Список сегментов (плит) для визуализации
    """
    if reinforcement_map is None:
        reinforcement_map = {}
    
    sequence = []
    _missing_reinf_logged = set()
    import logging
    _log = logging.getLogger(__name__)

    def _warn_missing_reinforcement(length, width_mm, load_code_from_cut):
        if load_code_from_cut is None:
            return
        key = (round(length, 3), width_mm, load_code_from_cut)
        if key in _missing_reinf_logged:
            return
        _missing_reinf_logged.add(key)
        _log.warning(
            "[ВИЗУАЛИЗАЦИЯ] Нет армирования в карте для (length=%s, width_mm=%s, load_code=%s)",
            length, width_mm, load_code_from_cut,
        )

    # Проверяем, есть ли 2D данные (plate_assignments)
    use_2d_data = 'plate_assignments' in plan and plan['plate_assignments']
    
    if not use_2d_data:
        print("[VISUAL] ⚠️ 2D данных нет, используем приближение")
        # Собираем все плиты с их длинами из cfg (модуль уже импортирован в начале файла)
        all_plates_with_lengths = []
        for plates, width_mm in [
            (cfg.PLATES_1_2, 1200), (cfg.PLATES_1_08, 1080),
            (cfg.PLATES_0_32, 320), (cfg.PLATES_0_46, 460), (cfg.PLATES_0_70, 700),
            (cfg.PLATES_0_72, 720), (cfg.PLATES_0_86, 860), (cfg.PLATES_0_88, 880),
            (cfg.PLATES_0_74, 740), (cfg.PLATES_0_48, 480), (cfg.PLATES_0_50, 500),
            (cfg.PLATES_0_34, 340)
        ]:
            for length in plates:
                all_plates_with_lengths.append({'length': length, 'width': width_mm})
        all_plates_with_lengths.sort(key=lambda x: (-x['width'], -x['length']))
    
    # Создаём карту поперечных резов
    transverse_cut_map = {}
    if plan.get('transverse_cuts'):
        for tcut in plan['transverse_cuts']:
            key = (tcut['source_length'], tcut['source_width'])
            transverse_cut_map[key] = {
                'target_length': tcut['target_length'],
                'remainder': tcut['remainder']
            }
    
    # Создаём карту вторичных резов
    secondary_cuts_info = {}
    secondary_cuts_by_parent: dict[str, list[dict]] = defaultdict(list)
    secondary_total_from_plan = 0
    secondary_attached_total = 0
    unmatched_by_reason = defaultdict(int)
    legacy_secondary_match_used = 0
    if plan.get('secondary_cuts'):
        for sec_cut in plan['secondary_cuts']:
            source_mm = sec_cut['source']
            pieces = sec_cut.get('pieces', 1)
            cuts_list = sec_cut.get('cuts', [])
            qty = sec_cut['qty']
            secondary_total_from_plan += int(qty) * max(1, int(pieces))
            
            source_lengths_list = sec_cut.get('source_lengths', [])
            target_lengths_list = sec_cut.get('lengths', [])
            target_order_key = sec_cut.get('target_order_key')
            target_load_code = cfg.normalize_load_code(target_order_key[2]) if (target_order_key and len(target_order_key) > 2) else None
            parent_ids_list = sec_cut.get('parent_instance_ids') or []
            secondary_ids_list = sec_cut.get('secondary_instance_ids') or []
            
            pattern = []
            if cuts_list:
                target_width_mm = cuts_list[0]
                for _ in range(pieces):
                    pattern.append({
                        'width': target_width_mm / 1000.0,
                        'width_mm': target_width_mm,
                        'source_width_mm': source_mm,
                        'label': None,
                        'target_length': target_lengths_list[0] if target_lengths_list else None,
                        'target_load_code': target_load_code
                    })
            
            for i in range(qty):
                source_length = source_lengths_list[i] if i < len(source_lengths_list) else 6.0
                key = _secondary_geom_cut_key(source_length, source_mm)
                parent_instance_id = (
                    parent_ids_list[i] if i < len(parent_ids_list)
                    else sec_cut.get('parent_instance_id')
                )
                secondary_instance_id = (
                    secondary_ids_list[i] if i < len(secondary_ids_list)
                    else sec_cut.get('secondary_instance_id')
                )
                
                if key not in secondary_cuts_info:
                    secondary_cuts_info[key] = []
                
                variant = {
                    'pattern': [segment.copy() for segment in pattern],
                    'qty': 1,
                    'used': 0,
                    # BUG-4 FIX: canonical ключ заказа для точного match в rescue
                    'target_order_key': target_order_key,
                    'parent_instance_id': parent_instance_id,
                    'secondary_instance_id': secondary_instance_id,
                }
                secondary_cuts_info[key].append(variant)
                if parent_instance_id:
                    secondary_cuts_by_parent[str(parent_instance_id)].append(variant)
    
    # ========== НОВАЯ ЛОГИКА: РАЗДЕЛИТЕЛИ МЕЖДУ ГРУППАМИ РЕЗОВ ==========
    # Требования завода (ОБНОВЛЁННЫЕ):
    # 1. Первая плита ДОЛЖНА быть целой с МИНИМАЛЬНЫМ армированием
    # 2. Плиты с одинаковым резом И одинаковым армированием идут подряд
    # 3. Группы сортируются по армированию (от меньшего к большему)
    # 4. Между группами с РАЗНЫМ резом/армированием должна быть целая плита-разделитель с мин. армированием
    all_primary_cuts = plan.get('primary_cuts', [])
    # ИСПРАВЛЕНО: Целая плита = любая плита без реза (rest=0)
    # Раньше фильтровали только 1200мм, из-за чего плиты 1080мм и другие терялись!
    solid_cuts = [cut for cut in all_primary_cuts 
                  if cut['rest'] == 0]
    
    # НОВОЕ: Вычисляем армирование для каждой целой плиты и сортируем по армированию (мин. первый)
    for cut in solid_cuts:
        length = cut['lengths'][0] if cut.get('lengths') else 6.0
        width_mm = cut['width']
        lc = cfg.normalize_load_code(cut.get('load_code', 8))
        cut['reinforcement'] = _get_reinforcement_from_map(reinforcement_map, length, width_mm, lc) or 999.0
    solid_cuts.sort(key=lambda x: (x.get('reinforcement', 999.0), -x['lengths'][0] if x.get('lengths') else 0))
    
    # Плиты с резом - вычисляем армирование для группировки
    cut_with_rest_raw = [cut for cut in all_primary_cuts if cut['rest'] > 0]
    
    # НОВОЕ: Добавляем армирование к каждой плите с резом для группировки
    for cut in cut_with_rest_raw:
        length = cut['lengths'][0] if cut.get('lengths') else 6.0
        width_mm = cut['width']
        lc = cfg.normalize_load_code(cut.get('load_code', 8))
        cut['reinforcement'] = _get_reinforcement_from_map(reinforcement_map, length, width_mm, lc) or 999.0
    
    # Сортируем плиты с резом по армированию (мин. первый), потом по типу реза
    cut_with_rest = sorted(
        cut_with_rest_raw,
        key=lambda x: (x.get('reinforcement', 999.0), x['width'], x['rest'])
    )
    
    print(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
    if solid_cuts:
        print(f"[VISUAL] Целые плиты (сортировка по армированию): {[(c['width'], c['qty'], c.get('reinforcement', '?')) for c in solid_cuts[:5]]}")
    
    # НОВОЕ: Группируем плиты с резом по (width, rest, reinforcement)
    # Это гарантирует что плиты с одинаковым резом И армированием идут вместе
    from itertools import groupby
    cut_groups = [list(group) for key, group in groupby(cut_with_rest, key=lambda x: (x['width'], x['rest'], x.get('reinforcement', 999.0)))]
    
    if cut_groups:
        print(f"[VISUAL] Найдено {len(cut_groups)} групп резов (сгруппировано по рез+армирование):")
        for i, group in enumerate(cut_groups, 1):
            print(f"[VISUAL]   Группа {i}: width={group[0]['width']}мм, rest={group[0]['rest']}мм, армирование={group[0].get('reinforcement', '?'):.1f}, плит={sum(c['qty'] for c in group)}")
    
    # Формируем последовательность с разделителями
    ordered_cuts = []
    
    # ВАЖНО: Разворачиваем записи целых плит в отдельные плиты
    # Из записи {qty: 5, lengths: [6.15, 6.15, ...]} создаём 5 записей {qty: 1, lengths: [6.15]}
    solid_cuts_list = []
    for cut in solid_cuts:
        lengths = cut.get('lengths', [])
        for i in range(cut['qty']):
            single_cut = cut.copy()
            single_cut['qty'] = 1  # Каждая запись = 1 плита
            single_cut['lengths'] = [lengths[i]] if i < len(lengths) else [lengths[0] if lengths else 6.0]
            solid_cuts_list.append(single_cut)
    
    print(f"[VISUAL] Развёрнуто {len(solid_cuts_list)} отдельных целых плит для разделителей")
    
    # Правило 1: Первая плита ОБЯЗАТЕЛЬНО целая
    if solid_cuts_list:
        first_plate = solid_cuts_list.pop(0)
        ordered_cuts.append(first_plate)
        first_width = first_plate.get('width', 1200)
        print(f"[VISUAL] ✓ Первая плита: целая {first_width}мм")
    
    # Правило 2 и 3: Чередуем группы резов и целые плиты-разделители
    for i, cut_group in enumerate(cut_groups):
        # Рассчитываем суммарную длину группы
        total_group_length = sum(
            cut['lengths'][0] * cut['qty'] 
            for cut in cut_group
            if cut.get('lengths')
        )
        
        print(f"[VISUAL] Группа резов #{i+1}: width={cut_group[0]['width']}мм, rest={cut_group[0]['rest']}мм, "
              f"типов={len(cut_group)}, длина={total_group_length:.1f}м")
        
        # Если группа слишком большая (>90м), разбиваем на подгруппы
        if total_group_length > 90.0:
            print(f"[VISUAL] ⚠️ Группа #{i+1} слишком большая ({total_group_length:.1f}м > 90м), разбиваем на подгруппы")
            subgroups = _split_group_into_subgroups(cut_group, max_length=90.0)
            
            # Добавляем подгруппы с разделителями МЕЖДУ ними
            for j, subgroup in enumerate(subgroups):
                ordered_cuts.extend(subgroup)
                print(f"[VISUAL]   Добавлена подгруппа #{i+1}.{j+1}: {len(subgroup)} плит")
                
                # После каждой подгруппы (кроме последней в этой группе)
                if j < len(subgroups) - 1 and solid_cuts_list:
                    separator = solid_cuts_list.pop(0)
                    separator['is_separator'] = True
                    ordered_cuts.append(separator)
                    print(f"[VISUAL]   ✓ Разделитель между подгруппами: целая плита")
        else:
            # Группа помещается в 1 дорожку - добавляем как есть
            ordered_cuts.extend(cut_group)
            print(f"[VISUAL] Добавлена группа резов #{i+1} (влезает в дорожку)")
        
        # После каждой ГРУППЫ (кроме последней) добавляем ОПТИМАЛЬНЫЙ разделитель
        if i < len(cut_groups) - 1 and solid_cuts_list:
            # Определяем следующую группу
            next_group = cut_groups[i + 1]
            
            # Умный выбор разделителя по армированию
            best_idx = _choose_best_separator(solid_cuts_list, next_group, reinforcement_map)
            
            if best_idx is not None:
                # Извлекаем выбранную плиту
                separator = solid_cuts_list.pop(best_idx)
                separator['is_separator'] = True  # МЯГКОЕ РЕЗЕРВИРОВАНИЕ: помечаем как разделитель
                ordered_cuts.append(separator)
                print(f"[VISUAL] ✓ Разделитель (is_separator=True): целая плита между группами")
            else:
                # Fallback: если функция не смогла выбрать, берём первую
                if solid_cuts_list:
                    fallback_sep = solid_cuts_list.pop(0)
                    fallback_sep['is_separator'] = True  # МЯГКОЕ РЕЗЕРВИРОВАНИЕ
                    ordered_cuts.append(fallback_sep)
                    print(f"[VISUAL] ✓ Разделитель: целая плита между группами (fallback, is_separator=True)")
    
    # Оставшиеся целые плиты добавляем в конец
    if solid_cuts_list:
        ordered_cuts.extend(solid_cuts_list)
        print(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")
    
    # Обрабатываем первичные резы В НОВОМ ПОРЯДКЕ: целая → группа резов → целая → группа резов
    for cut in ordered_cuts:
        width_mm = cut['width']
        rest_mm = cut['rest']
        qty = cut['qty']
        primary_instance_ids = cut.get('primary_instance_ids') or []
        if cut.get('primary_instance_id') and not primary_instance_ids:
            primary_instance_ids = [cut.get('primary_instance_id')]
        
        # Получаем длины для этих плит
        if use_2d_data and 'lengths' in cut:
            lengths_for_cut = cut['lengths']
        else:
            matching_plates = [p for p in all_plates_with_lengths if p['width'] == width_mm]
            lengths_for_cut = [p['length'] for p in matching_plates[:qty]]
            while len(lengths_for_cut) < qty:
                lengths_for_cut.append(6.0 if not matching_plates else matching_plates[0]['length'])
        
        for i in range(qty):
            length = lengths_for_cut[i] if i < len(lengths_for_cut) else 6.0
            parent_instance_id = (
                primary_instance_ids[i] if i < len(primary_instance_ids)
                else cut.get('primary_instance_id')
            )
            
            # НОВОЕ: Получаем информацию о КП для этой плиты
            kp_id = cut.get('kp_id')
            customer = cut.get('customer')
            kp_date = cut.get('kp_date')
            plate_name_from_cut = cut.get('plate_name')
            
            transverse_cut_info = transverse_cut_map.get((length, width_mm))
            
            if transverse_cut_info:
                # Поперечный рез
                width_m = width_mm / 1000.0
                # Получаем армирование из карты (с учётом load_code)
                load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 800))
                reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                if reinforcement is None:
                    _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                sequence.append({
                    'length': length,
                    'mode': 'transverse',
                    'target_length': transverse_cut_info['target_length'],
                    'remainder': transverse_cut_info['remainder'],
                    'width': width_m,
                    'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code
                    'label_target': plate_label_func(transverse_cut_info['target_length'], width_m, load_code_from_cut),
                    'label_remainder': f'Остаток {transverse_cut_info["remainder"]:.2f}м'.replace('.', ',') if transverse_cut_info['remainder'] > 0.1 else '',
                    'reinforcement': reinforcement,
                    'kp_id': kp_id,
                    'customer': customer,
                    'kp_date': kp_date,
                    'plate_name': plate_name_from_cut,
                    'unit_id': parent_instance_id,
                    'layout_uid': str(parent_instance_id) if parent_instance_id else f"transverse:{len(sequence)}",
                })
            else:
                # Обычная плита
                main_w = width_mm / 1000.0
                rest_w = rest_mm / 1000.0
                fake_rest_override = False
                
                # Обработка плит 1.08м
                if width_mm == 1080 and rest_mm == 0:
                    rest_mm = 120
                    rest_w = 0.12
                    fake_rest_override = True
                
                if rest_mm == 0:
                    # Плита без реза
                    # Получаем армирование из карты по (length, width_mm, load_code)
                    load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 800))
                    reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    # #region agent log
                    if abs(length - 7.1) < 0.05 and load_code_from_cut in (12, 12.5, 13):
                        try:
                            open(_DEBUG_LOG, "a", encoding="utf-8").write(__import__("json").dumps({"hypothesisId": "H71item2d", "location": "layout_sequence:sequence_append_solid_2d", "message": "71-12п item в sequence (2d)", "data": {"length": length, "width_mm": width_mm, "load_code_from_cut": load_code_from_cut, "reinforcement": reinforcement}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                    # #endregion
                    if not reinforcement:
                        print(f"[VISUAL] ⚠️ Армирование не найдено для целой плиты: {length}м x {width_mm}мм (load_code={load_code_from_cut})")
                        print(f"[VISUAL]    Доступные ключи в карте: {list(reinforcement_map.keys())[:5]}")
                    else:
                        print(f"[VISUAL] ✓ Армирование найдено для целой плиты: {length}м x {width_mm}мм = {reinforcement:.1f}")
                    # МЯГКОЕ РЕЗЕРВИРОВАНИЕ: передаём флаг is_separator
                    is_separator = cut.get('is_separator', False)
                    sequence.append({
                        'length': length,
                        'mode': 'solid',
                        'width': main_w,  # ИСПРАВЛЕНИЕ: сохраняем реальную ширину для корректного учёта
                        'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code для различения плит
                        'label': plate_label_func(length, main_w, load_code_from_cut),
                        'reinforcement': reinforcement,
                        'is_separator': is_separator,  # Для приоритета при разбиении на дорожки
                        'kp_id': kp_id,
                        'customer': customer,
                        'kp_date': kp_date,
                        'plate_name': plate_name_from_cut,
                        'unit_id': parent_instance_id,
                        'layout_uid': str(parent_instance_id) if parent_instance_id else f"solid:{len(sequence)}",
                    })
                else:
                    # Плита с резом
                    load_code_from_cut = cfg.normalize_load_code(cut.get('load_code', 800))
                    secondary_cuts_for_plate = None
                    chosen_variant = None
                    # Сначала варианты по parent_instance_id; если не выбрали — геометрия
                    # (иначе непустой, но «неподходящий» by_parent блокирует legacy_match_used=0).
                    _pool_parent: list = []
                    if parent_instance_id:
                        _pool_parent = secondary_cuts_by_parent.get(str(parent_instance_id)) or []
                    for variant in _pool_parent:
                        if variant['used'] < variant['qty'] and (variant.get('pattern') or []):
                            chosen_variant = variant
                            break
                    if not chosen_variant:
                        _pool_geom = secondary_cuts_info.get(_secondary_geom_cut_key(length, rest_mm)) or []
                        if _pool_geom:
                            legacy_secondary_match_used += 1
                            for variant in _pool_geom:
                                if variant['used'] < variant['qty'] and (variant.get('pattern') or []):
                                    chosen_variant = variant
                                    break
                    
                    if chosen_variant:
                        secondary_cuts_for_plate = []
                        # BUG-4 FIX: canonical ключ заказа из оптимизатора для точного match
                        _sec_tok_plan = chosen_variant.get('target_order_key')
                        for sec_idx, sec_cut_template in enumerate(chosen_variant['pattern']):
                            sec_unit_id = chosen_variant.get('secondary_instance_id')
                            if sec_unit_id is None:
                                sec_unit_id = f"{parent_instance_id or 'secondary'}:{length}:{rest_mm}:{sec_idx}:{chosen_variant['used']}"
                            sec_width = sec_cut_template['width']
                            sec_width_mm = sec_cut_template['width_mm']
                            lc = sec_cut_template.get('target_load_code')
                            lc = cfg.normalize_load_code(lc) if lc is not None else load_code_from_cut
                            
                            sec_transverse = transverse_cut_map.get((length, sec_width_mm))
                            
                            if sec_transverse:
                                secondary_cuts_for_plate.append({
                                    'width': sec_width,
                                    'label': f'[2] {plate_label_func(sec_transverse["target_length"], sec_width, lc)}',
                                    'transverse_cut': True,
                                    'target_length': sec_transverse['target_length'],
                                    'remainder': sec_transverse['remainder'],
                                    'load_code': lc,
                                    'target_order_key': _sec_tok_plan,
                                    'parent_unit_id': parent_instance_id,
                                    'unit_id': sec_unit_id,
                                })
                            else:
                                target_length = sec_cut_template.get('target_length')
                                if target_length:
                                    secondary_cuts_for_plate.append({
                                        'width': sec_width,
                                        'label': f'О {plate_label_func(target_length, sec_width, lc)}',
                                        'has_transverse': True,
                                        'target_length': target_length,
                                        'load_code': lc,
                                        'target_order_key': _sec_tok_plan,
                                        'parent_unit_id': parent_instance_id,
                                        'unit_id': sec_unit_id,
                                    })
                                else:
                                    result_width = sec_cut_template['width']
                                    source_width = sec_cut_template.get('source_width_mm', result_width * 1000) / 1000.0
                                    label_text = plate_label_func(length, result_width, lc)
                                    if abs(result_width - source_width) > 1e-6:
                                        label_text = f'О {label_text}'
                                    secondary_cuts_for_plate.append({
                                        'width': result_width,
                                        'label': label_text,
                                        'load_code': lc,
                                        'target_order_key': _sec_tok_plan,
                                        'parent_unit_id': parent_instance_id,
                                        'unit_id': sec_unit_id,
                                    })
                        chosen_variant['used'] += 1
                        secondary_attached_total += len(secondary_cuts_for_plate)
                    elif rest_mm > 0:
                        if parent_instance_id:
                            unmatched_by_reason["parent_instance_id_not_found"] += 1
                        else:
                            unmatched_by_reason["key_not_found"] += 1
                    
                    # Получаем армирование из карты (с учётом load_code)
                    reinforcement = _get_reinforcement_from_map(reinforcement_map, length, width_mm, load_code_from_cut)
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    if reinforcement:
                        print(f"[VISUAL] ✓ Армирование найдено для плиты с резом: {length}м x {width_mm}мм = {reinforcement:.1f}")
                    sequence.append({
                        'length': length,
                        'mode': 'split',
                        'main_w': main_w,
                        'rest_w': rest_w,
                        'load_code': load_code_from_cut,  # ИСПРАВЛЕНИЕ: добавляем load_code
                        'label_main': plate_label_func(length, main_w, load_code_from_cut),
                        'label_rest': (
                            '+0,12' if fake_rest_override else
                            (f'+{rest_w:.2f}'.replace('.', ',') if not secondary_cuts_for_plate else None)
                        ),
                        'secondary_cuts': secondary_cuts_for_plate,
                        'reinforcement': reinforcement,
                        'kp_id': kp_id,
                        'customer': customer,
                        'kp_date': kp_date,
                        'plate_name': plate_name_from_cut,
                        'unit_id': parent_instance_id,
                        'layout_uid': str(parent_instance_id) if parent_instance_id else f"split:{len(sequence)}",
                    })
    
    _ensure_sequence_layout_uid(sequence, prefix="built")
    secondary_unmatched_total = max(0, secondary_total_from_plan - secondary_attached_total)
    if secondary_unmatched_total or legacy_secondary_match_used:
        _log.warning(
            "[LAYOUT_SEQUENCE] secondary mapping report: total=%s attached=%s unmatched=%s reasons=%s legacy_match_used=%s",
            secondary_total_from_plan,
            secondary_attached_total,
            secondary_unmatched_total,
            dict(unmatched_by_reason),
            legacy_secondary_match_used,
        )
    return sequence

