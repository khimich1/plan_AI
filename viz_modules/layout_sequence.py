#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль построения последовательности раскладки плит:
- Формирование последовательности сегментов вдоль дорожки
"""
import core.config_and_data as cfg
from core.optimization import OPT_PLAN, OPT_WIDTH_PRIORITY


def build_layout_sequence():
    """Формирует последовательность сегментов вдоль дорожки, РАЗДЕЛЁННУЮ ПО НАГРУЗКАМ."""
    from core.optimization import OPT_CASCADING_PLAN, OPT_CASCADING_PLAN_BY_LOAD
    from core.reinforcement_db import get_reinforcement
    from pathlib import Path
    
    # Создаём глобальную карту армирования из PLATE_LOAD_DETAILS
    reinforcement_map = {}  # {(length, width_mm): reinforcement}
    
    if cfg.PLATE_LOAD_DETAILS:
        db_path = Path(__file__).parent.parent / "pb.db"
        print(f"[VISUAL] Начинаем создание карты армирования из {len(cfg.PLATE_LOAD_DETAILS)} записей")
        for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
            width_mm = int(round(width_m * 1000))
            reinforcement = get_reinforcement(
                length_m=length,
                load_code=load_code,
                source='series',
                db_path=db_path,
                allow_fallback=True
            )
            key = (length, width_mm)
            if reinforcement and reinforcement < 999:
                reinforcement_map[key] = reinforcement
                print(f"[VISUAL]   Добавлено: ({length}м, {width_mm}мм) → армирование {reinforcement:.1f}")
    
    print(f"[VISUAL] Создана карта армирования: {len(reinforcement_map)} записей")
    sequence = []

    def plate_label(L: float, W: float) -> str:
        Ldm = int(round(L * 10))
        Wdm_val = round(W * 10, 1)
        if abs(Wdm_val - int(Wdm_val)) < 1e-6:
            Wdm = str(int(Wdm_val))
        else:
            Wdm = str(Wdm_val).replace('.', ',')
        return f'ПБ {Ldm}-{Wdm}-8п'
    
    # ✅ НОВЫЙ ПРИОРИТЕТ 0: OPT_CASCADING_PLAN_BY_LOAD (группировка по нагрузкам)
    print(f"[VISUAL] Проверяем OPT_CASCADING_PLAN_BY_LOAD: {bool(OPT_CASCADING_PLAN_BY_LOAD)}")
    if OPT_CASCADING_PLAN_BY_LOAD:
        print(f"[VISUAL] ✅ Используем группировку по нагрузкам! Групп: {len(OPT_CASCADING_PLAN_BY_LOAD)}")
        
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
        
        # НОВЫЙ ФОРМАТ ВОЗВРАТА: список групп по нагрузкам
        return all_sequences
    
    # Приоритет 1: OPT_CASCADING_PLAN (старый формат, без группировки)
    print(f"[VISUAL] Проверяем OPT_CASCADING_PLAN: {OPT_CASCADING_PLAN is not None}")
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
                            'target_length': target_lengths_list[0] if target_lengths_list else None  # Для поперечных резов
                        })
                
                # ИСПРАВЛЕНИЕ: Создаём запись для КАЖДОЙ ИСХОДНОЙ длины отдельно
                for i in range(qty):
                    source_length = source_lengths_list[i] if i < len(source_lengths_list) else 6.0
                    key = (source_length, source_mm)

                    if key not in secondary_cuts_info:
                        secondary_cuts_info[key] = []

                    secondary_cuts_info[key].append({
                        'pattern': [segment.copy() for segment in pattern],
                        'qty': 1,
                        'used': 0
                    })
        
        print(f"[VISUAL] Создано {len(secondary_cuts_info)} вариантов вторичных резов:")
        for (src_len, src_w), variants in secondary_cuts_info.items():
            for idx, info in enumerate(variants, start=1):
                pattern_desc = ", ".join([f"{c['width_mm']}мм" for c in info['pattern']])
                print(f"  Остаток {src_len}м x {src_w}мм: вариант #{idx} -> [{pattern_desc}]")
        
        # ========== НОВАЯ ЛОГИКА: РАЗДЕЛИТЕЛИ МЕЖДУ ГРУППАМИ РЕЗОВ ==========
        # Требования завода:
        # 1. Первая плита ДОЛЖНА быть целой (без реза)
        # 2. Плиты с одинаковым резом идут подряд
        # 3. Между группами с РАЗНЫМ резом должна быть целая плита-разделитель
        all_primary_cuts = OPT_CASCADING_PLAN.get('primary_cuts', [])
        solid_cuts = [cut for cut in all_primary_cuts if cut['rest'] == 0]
        
        # Сортируем целые плиты
        solid_cuts.sort(key=lambda x: (-x['width'], -x['lengths'][0] if x.get('lengths') else 0))
        
        # Сортируем плиты с резом по типу реза
        cut_with_rest = sorted(
            [cut for cut in all_primary_cuts if cut['rest'] > 0],
            key=lambda x: (-x['rest'], -x['width'])
        )
        
        print(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
        if solid_cuts:
            print(f"[VISUAL] Целые плиты: {[(c['width'], c['qty']) for c in solid_cuts]}")
        
        # Группируем плиты с резом по типу реза (width, rest)
        # ВАЖНО: учитываем И ширину И остаток, т.к. это разные настройки станка!
        from itertools import groupby
        cut_groups = [list(group) for key, group in groupby(cut_with_rest, key=lambda x: (x['width'], x['rest']))]
        
        if cut_groups:
            print(f"[VISUAL] Найдено {len(cut_groups)} групп резов:")
            for i, group in enumerate(cut_groups, 1):
                print(f"[VISUAL]   Группа {i}: width={group[0]['width']}мм, rest={group[0]['rest']}мм, типов={len(group)}")
        
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
            ordered_cuts.append(solid_cuts_list.pop(0))
            print(f"[VISUAL] ✓ Первая плита: целая 1200мм")
        
        # Правило 2 и 3: Чередуем группы резов и целые плиты-разделители
        for i, cut_group in enumerate(cut_groups):
            ordered_cuts.extend(cut_group)
            print(f"[VISUAL] Добавлена группа резов #{i+1}: width={cut_group[0]['width']}мм, rest={cut_group[0]['rest']}мм, типов={len(cut_group)}")
            
            # После каждой группы (кроме последней) добавляем целую плиту-разделитель
            if i < len(cut_groups) - 1 and solid_cuts_list:
                ordered_cuts.append(solid_cuts_list.pop(0))
                print(f"[VISUAL] ✓ Разделитель: целая плита 1200мм между группами")
        
        # Оставшиеся целые плиты добавляем в конец
        if solid_cuts_list:
            ordered_cuts.extend(solid_cuts_list)
            print(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")
        
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
                
                # ИСПРАВЛЕНИЕ: Проверяем вторичные резы для КОНКРЕТНОГО остатка (длина + ширина)
                sec_variants = secondary_cuts_info.get((length, rest_mm)) or []
                
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
                    
                    # Получаем армирование из карты
                    reinforcement = reinforcement_map.get((length, width_mm))
                    
                    sequence.append({
                        'length': length,  # Исходная длина плиты
                        'mode': 'transverse',
                        'target_length': target_length,
                        'remainder': remainder,
                        'width': width_m,
                        'label_target': plate_label(target_length, width_m),
                        'label_remainder': f'Остаток {remainder:.2f}м'.replace('.', ',') if remainder > 0.1 else '',
                        'reinforcement': reinforcement
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
                        # Получаем армирование из карты
                        reinforcement = reinforcement_map.get((length, width_mm))
                        sequence.append({
                            'length': length,
                            'mode': 'solid',
                            'label': plate_label(length, main_w),
                            'reinforcement': reinforcement
                        })
                    else:
                        # Плиты С резом
                        # Проверяем, нужны ли вторичные резы для ЭТОЙ плиты
                        secondary_cuts_for_plate = None
                        chosen_variant = None
                        for variant in sec_variants:
                            if variant['used'] < variant['qty']:
                                chosen_variant = variant
                                break

                        if chosen_variant:
                            secondary_cuts_for_plate = []
                            for sec_cut_template in chosen_variant['pattern']:
                                sec_width = sec_cut_template['width']
                                sec_width_mm = sec_cut_template['width_mm']
                                
                                # ВАЖНО: Проверяем поперечные резы для ВТОРИЧНЫХ плит!
                                sec_transverse = transverse_cut_map.get((length, sec_width_mm))
                                
                                if sec_transverse:
                                    # Вторичная плита с поперечным резом
                                    secondary_cuts_for_plate.append({
                                        'width': sec_width,
                                        'label': f'[2] {plate_label(sec_transverse["target_length"], sec_width)}',
                                        'transverse_cut': True,
                                        'target_length': sec_transverse['target_length'],
                                        'remainder': sec_transverse['remainder']
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
                                            'label': f'О {plate_label(target_length, sec_width)}',  # О = Остаток
                                            'has_transverse': True,  # Флаг для отрисовки красной линии
                                            'target_length': target_length  # Длина результата (для правильной отрисовки)
                                        })
                                    else:
                                        # Обычный вторичный рез (включая narrowing)
                                        result_width = sec_cut_template['width']
                                        source_width = sec_cut_template.get('source_width_mm', result_width * 1000) / 1000.0
                                        label_text = plate_label(length, result_width)
                                        if abs(result_width - source_width) > 1e-6:
                                            label_text = f'О {label_text}'  # помечаем, что получено из остатка
                                        secondary_cuts_for_plate.append({
                                            'width': result_width,
                                            'label': label_text
                                        })
                            chosen_variant['used'] += 1
                        
                        # Получаем армирование из карты
                        reinforcement = reinforcement_map.get((length, width_mm))
                        sequence.append({
                            'length': length,
                            'mode': 'split',
                            'main_w': main_w,
                            'rest_w': rest_w,
                            'label_main': plate_label(length, main_w),
                            'label_rest': (
                                '+0,12' if fake_rest_override else
                                (f'+{rest_w:.2f}'.replace('.', ',') if not secondary_cuts_for_plate else None)
                            ),
                            'secondary_cuts': secondary_cuts_for_plate,
                            'reinforcement': reinforcement
                        })
        
        if sequence:
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


def _build_sequence_from_plan(plan, plate_label_func, reinforcement_map=None):
    """
    Вспомогательная функция: строит последовательность плит из плана оптимизации.
    
    Args:
        plan: Результат оптимизации (OPT_CASCADING_PLAN)
        plate_label_func: Функция для создания меток плит
        reinforcement_map: Словарь {(length, width_mm): reinforcement} для получения армирования
    
    Returns:
        Список сегментов (плит) для визуализации
    """
    if reinforcement_map is None:
        reinforcement_map = {}
    
    sequence = []
    
    # Проверяем, есть ли 2D данные (plate_assignments)
    use_2d_data = 'plate_assignments' in plan and plan['plate_assignments']
    
    if not use_2d_data:
        print("[VISUAL] ⚠️ 2D данных нет, используем приближение")
        # Собираем все плиты с их длинами из cfg
        import core.config_and_data as cfg
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
    if plan.get('secondary_cuts'):
        for sec_cut in plan['secondary_cuts']:
            source_mm = sec_cut['source']
            pieces = sec_cut.get('pieces', 1)
            cuts_list = sec_cut.get('cuts', [])
            qty = sec_cut['qty']
            
            source_lengths_list = sec_cut.get('source_lengths', [])
            target_lengths_list = sec_cut.get('lengths', [])
            
            pattern = []
            if cuts_list:
                target_width_mm = cuts_list[0]
                for _ in range(pieces):
                    pattern.append({
                        'width': target_width_mm / 1000.0,
                        'width_mm': target_width_mm,
                        'source_width_mm': source_mm,
                        'label': None,
                        'target_length': target_lengths_list[0] if target_lengths_list else None
                    })
            
            for i in range(qty):
                source_length = source_lengths_list[i] if i < len(source_lengths_list) else 6.0
                key = (source_length, source_mm)
                
                if key not in secondary_cuts_info:
                    secondary_cuts_info[key] = []
                
                secondary_cuts_info[key].append({
                    'pattern': [segment.copy() for segment in pattern],
                    'qty': 1,
                    'used': 0
                })
    
    # ========== НОВАЯ ЛОГИКА: РАЗДЕЛИТЕЛИ МЕЖДУ ГРУППАМИ РЕЗОВ ==========
    # Требования завода:
    # 1. Первая плита ДОЛЖНА быть целой (без реза)
    # 2. Плиты с одинаковым резом идут подряд
    # 3. Между группами с РАЗНЫМ резом должна быть целая плита-разделитель
    all_primary_cuts = plan.get('primary_cuts', [])
    solid_cuts = [cut for cut in all_primary_cuts if cut['rest'] == 0]
    
    # Сортируем целые плиты
    solid_cuts.sort(key=lambda x: (-x['width'], -x['lengths'][0] if x.get('lengths') else 0))
    
    # Сортируем плиты с резом по типу реза
    cut_with_rest = sorted(
        [cut for cut in all_primary_cuts if cut['rest'] > 0],
        key=lambda x: (-x['rest'], -x['width'])
    )
    
    print(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
    if solid_cuts:
        print(f"[VISUAL] Целые плиты: {[(c['width'], c['qty']) for c in solid_cuts]}")
    
    # Группируем плиты с резом по типу реза (width, rest)
    # ВАЖНО: учитываем И ширину И остаток, т.к. это разные настройки станка!
    from itertools import groupby
    cut_groups = [list(group) for key, group in groupby(cut_with_rest, key=lambda x: (x['width'], x['rest']))]
    
    if cut_groups:
        print(f"[VISUAL] Найдено {len(cut_groups)} групп резов:")
        for i, group in enumerate(cut_groups, 1):
            print(f"[VISUAL]   Группа {i}: width={group[0]['width']}мм, rest={group[0]['rest']}мм, типов={len(group)}")
    
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
        ordered_cuts.append(solid_cuts_list.pop(0))
        print(f"[VISUAL] ✓ Первая плита: целая 1200мм")
    
    # Правило 2 и 3: Чередуем группы резов и целые плиты-разделители
    for i, cut_group in enumerate(cut_groups):
        ordered_cuts.extend(cut_group)
        print(f"[VISUAL] Добавлена группа резов #{i+1}: width={cut_group[0]['width']}мм, rest={cut_group[0]['rest']}мм, типов={len(cut_group)}")
        
        # После каждой группы (кроме последней) добавляем целую плиту-разделитель
        if i < len(cut_groups) - 1 and solid_cuts_list:
            ordered_cuts.append(solid_cuts_list.pop(0))
            print(f"[VISUAL] ✓ Разделитель: целая плита 1200мм между группами")
    
    # Оставшиеся целые плиты добавляем в конец
    if solid_cuts_list:
        ordered_cuts.extend(solid_cuts_list)
        print(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")
    
    # Обрабатываем первичные резы В НОВОМ ПОРЯДКЕ: целая → группа резов → целая → группа резов
    for cut in ordered_cuts:
        width_mm = cut['width']
        rest_mm = cut['rest']
        qty = cut['qty']
        
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
            
            sec_variants = secondary_cuts_info.get((length, rest_mm)) or []
            transverse_cut_info = transverse_cut_map.get((length, width_mm))
            
            if transverse_cut_info:
                # Поперечный рез
                width_m = width_mm / 1000.0
                # Получаем армирование из карты
                reinforcement = reinforcement_map.get((length, width_mm))
                sequence.append({
                    'length': length,
                    'mode': 'transverse',
                    'target_length': transverse_cut_info['target_length'],
                    'remainder': transverse_cut_info['remainder'],
                    'width': width_m,
                    'label_target': plate_label_func(transverse_cut_info['target_length'], width_m),
                    'label_remainder': f'Остаток {transverse_cut_info["remainder"]:.2f}м'.replace('.', ',') if transverse_cut_info['remainder'] > 0.1 else '',
                    'reinforcement': reinforcement
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
                    # Получаем армирование из карты по (length, width_mm)
                    reinforcement = reinforcement_map.get((length, width_mm))
                    if not reinforcement:
                        print(f"[VISUAL] ⚠️ Армирование не найдено для целой плиты: {length}м x {width_mm}мм")
                        print(f"[VISUAL]    Доступные ключи в карте: {list(reinforcement_map.keys())[:5]}")
                    else:
                        print(f"[VISUAL] ✓ Армирование найдено для целой плиты: {length}м x {width_mm}мм = {reinforcement:.1f}")
                    sequence.append({
                        'length': length,
                        'mode': 'solid',
                        'label': plate_label_func(length, main_w),
                        'reinforcement': reinforcement
                    })
                else:
                    # Плита с резом
                    secondary_cuts_for_plate = None
                    chosen_variant = None
                    for variant in sec_variants:
                        if variant['used'] < variant['qty']:
                            chosen_variant = variant
                            break
                    
                    if chosen_variant:
                        secondary_cuts_for_plate = []
                        for sec_cut_template in chosen_variant['pattern']:
                            sec_width = sec_cut_template['width']
                            sec_width_mm = sec_cut_template['width_mm']
                            
                            sec_transverse = transverse_cut_map.get((length, sec_width_mm))
                            
                            if sec_transverse:
                                secondary_cuts_for_plate.append({
                                    'width': sec_width,
                                    'label': f'[2] {plate_label_func(sec_transverse["target_length"], sec_width)}',
                                    'transverse_cut': True,
                                    'target_length': sec_transverse['target_length'],
                                    'remainder': sec_transverse['remainder']
                                })
                            else:
                                target_length = sec_cut_template.get('target_length')
                                if target_length:
                                    secondary_cuts_for_plate.append({
                                        'width': sec_width,
                                        'label': f'О {plate_label_func(target_length, sec_width)}',
                                        'has_transverse': True,
                                        'target_length': target_length
                                    })
                                else:
                                    result_width = sec_cut_template['width']
                                    source_width = sec_cut_template.get('source_width_mm', result_width * 1000) / 1000.0
                                    label_text = plate_label_func(length, result_width)
                                    if abs(result_width - source_width) > 1e-6:
                                        label_text = f'О {label_text}'
                                    secondary_cuts_for_plate.append({
                                        'width': result_width,
                                        'label': label_text
                                    })
                        chosen_variant['used'] += 1
                    
                    # Получаем армирование из карты
                    reinforcement = reinforcement_map.get((length, width_mm))
                    if reinforcement:
                        print(f"[VISUAL] ✓ Армирование найдено для плиты с резом: {length}м x {width_mm}мм = {reinforcement:.1f}")
                    sequence.append({
                        'length': length,
                        'mode': 'split',
                        'main_w': main_w,
                        'rest_w': rest_w,
                        'label_main': plate_label_func(length, main_w),
                        'label_rest': (
                            '+0,12' if fake_rest_override else
                            (f'+{rest_w:.2f}'.replace('.', ',') if not secondary_cuts_for_plate else None)
                        ),
                        'secondary_cuts': secondary_cuts_for_plate,
                        'reinforcement': reinforcement
                    })
    
    return sequence

