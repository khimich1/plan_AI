#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль визуализации и работы с ценами:
- Загрузка прайса из XLSX
- Работа с базой цен SQLite
- Построение сметы
- Визуализация раскладки плит
"""
import os
from collections import Counter
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.lines import Line2D

# Относительные импорты внутри core/
from . import config_and_data as cfg
from .optimization import optimize_cuts_pulp
from .price_db import init_schema, import_from_xlsx

# Импорты из новых модулей
from viz_modules.price_utils import load_price_table_from_xlsx
from viz_modules.procurement import (
    build_procurement_items,
    build_price_rows,
    build_component_breakdown,
    get_orders_from_opt_plan
)
from viz_modules.layout_sequence import build_layout_sequence
from viz_modules.visualization_drawing import (
    _draw_segment,
    _draw_split_plate,
    _draw_transverse_cut
)

try:
    import pandas as pd
except Exception:
    pd = None


# Реэкспорт для обратной совместимости
__all__ = ['visualize_plan', 'build_layout_sequence']


def visualize_plan(output_dir: str = 'Визуализация_Раскладки'):
    """Создаёт визуализацию раскладки плит и сохраняет файлы"""
    try:
        optimized = optimize_cuts_pulp({300: 4, 500: 3, 700: 2, 900: 2})
        print("Оптимальные резы:", optimized)
    except Exception as e:
        print("[OPT] Ошибка при оптимизации:", e)

    os.makedirs(output_dir, exist_ok=True)

    price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
    try:
        init_schema(cfg.PRICE_DB_PATH)
        written = import_from_xlsx(cfg.PRICE_XLSX_PATH, cfg.PRICE_DB_PATH)
        if written:
            print(f'[ПРАЙС->БД] записано строк: {written}')
    except Exception:
        pass

    price_rows, total_sum = build_price_rows(price_table)
    breakdown_tables = build_component_breakdown(price_table, price_rows)
    
    # Отладочная информация
    print(f'[DEBUG] breakdown_tables count: {len(breakdown_tables) if breakdown_tables else 0}')
    if breakdown_tables:
        for i, bt in enumerate(breakdown_tables):
            print(f'[DEBUG] breakdown_tables[{i}]: name={bt.get("name")}, rows={len(bt.get("rows", []))}')
    
    seq = build_layout_sequence()
    total_length = sum(s['length'] for s in seq)

    # Разбиваем плиты на несколько дорожек по 101 метру
    MAX_TRACK_LENGTH = 101.0  # Максимальная длина одной дорожки
    
    tracks = []  # Список дорожек
    current_track = []
    current_track_length = 0.0
    
    for item in seq:
        item_length = item['length']
        
        # Если плита не помещается - создаём новую дорожку
        if current_track_length + item_length > MAX_TRACK_LENGTH and current_track:
            tracks.append({
                'items': current_track,
                'length': current_track_length
            })
            current_track = []
            current_track_length = 0.0
        
        current_track.append(item)
        current_track_length += item_length
    
    # Добавляем последнюю дорожку
    if current_track:
        tracks.append({
            'items': current_track,
            'length': current_track_length
        })
    
    num_tracks = len(tracks)
    print(f"[ВИЗУАЛИЗАЦИЯ] Плиты разбиты на {num_tracks} дорожек по {MAX_TRACK_LENGTH}м")

    # Убрали секцию детальной разбивки - она теперь в отдельном Excel файле
    # Увеличиваем высоту секции дорожек пропорционально их количеству
    track_section_height = 3.0 + (num_tracks - 1) * 2.5
    num_sections = 4
    height_ratios = [track_section_height, 1.0, 1.4, 1.8]
    
    # Увеличиваем общую высоту окна
    total_fig_height = 16 + (num_tracks - 1) * 5
    
    fig = plt.figure(figsize=(22, total_fig_height))
    gs = fig.add_gridspec(num_sections, 1, height_ratios=height_ratios)
    ax_track = fig.add_subplot(gs[0, 0])
    ax_strips = fig.add_subplot(gs[1, 0])
    ax_table = fig.add_subplot(gs[2, 0])
    ax_price = fig.add_subplot(gs[3, 0])
    
    # Заголовок с количеством дорожек
    if num_tracks == 1:
        fig.suptitle('КЗ: Дорожка 1 (ширина 1.2 м) — раскладка, резы, ведомости и смета', 
                     fontsize=16, fontweight='bold')
    else:
        fig.suptitle(f'КЗ: Дорожки 1-{num_tracks} (ширина 1.2 м, по {MAX_TRACK_LENGTH}м) — раскладка, резы, ведомости и смета', 
                     fontsize=16, fontweight='bold')

    # Настройка осей для множественных дорожек
    track_height = cfg.TRACK_WIDTH_M  # 1.2 м
    track_spacing = 0.3  # Отступ между дорожками
    total_height = num_tracks * (track_height + track_spacing)
    
    ax_track.set_xlim(0, MAX_TRACK_LENGTH + 2)
    ax_track.set_ylim(0, total_height)
    ax_track.set_aspect('auto')
    ax_track.spines['top'].set_visible(False)
    ax_track.spines['right'].set_visible(False)
    
    # Метки по оси Y (номера дорожек)
    y_ticks = []
    y_labels = []
    for i in range(num_tracks):
        y_pos = i * (track_height + track_spacing) + track_height / 2
        y_ticks.append(y_pos)
        y_labels.append(f'Дор.{i+1}')
    ax_track.set_yticks(y_ticks)
    ax_track.set_yticklabels(y_labels)
    
    ax_track.set_xlabel('Длина (м)')
    ax_track.set_xticks(range(0, int(MAX_TRACK_LENGTH) + 1, 5))
    ax_track.grid(axis='x', linestyle=':', linewidth=0.5, alpha=0.5)

    # Рисуем каждую дорожку
    for track_idx, track_data in enumerate(tracks):
        # Y-координата текущей дорожки
        y_base = track_idx * (track_height + track_spacing)
        
        # Рамка дорожки
        track_rect = patches.Rectangle(
            (0, y_base), 
            MAX_TRACK_LENGTH, 
            track_height, 
            linewidth=2, 
            edgecolor='black', 
            facecolor='none', 
            linestyle='--'
        )
        ax_track.add_patch(track_rect)
        
        # Рисуем плиты в этой дорожке
        x = 0.0
        for item in track_data['items']:
            if item.get('mode') == 'solid':
                _draw_segment(ax_track, x, item['length'], '#2ecc71', item['label'], 
                            y=y_base, height=track_height)
            elif item.get('mode') == 'transverse':
                # Плита с поперечным резом (по длине)
                _draw_transverse_cut(
                    ax_track, x, 
                    total_length=item['length'],
                    target_length=item['target_length'],
                    width=item['width'],
                    label_target=item['label_target'],
                    remainder_length=item['remainder'],
                    y_base=y_base
                )
            else:
                # Плиты с резами (первичными и возможными вторичными)
                _draw_split_plate(
                    ax_track, x, item['length'],
                    main_w=item['main_w'], rest_w=item['rest_w'],
                    label_main=item['label_main'], label_rest=item.get('label_rest'),
                    secondary_cuts=item.get('secondary_cuts'),
                    y_base=y_base
                )
            x += item['length']

    legend_patches = [
        patches.Patch(facecolor='#2ecc71', edgecolor='black', label='🟢 Основа (первичный рез)'),
        patches.Patch(facecolor='#3498db', edgecolor='black', label='🔵 Вторичный рез (из остатка)'),
        patches.Patch(facecolor='#95a5a6', edgecolor='gray', label='⬛ Отход'),
        patches.Patch(facecolor='#ecf0f1', edgecolor='black', label='⬜ Остаток (не использован)'),
        Line2D([0], [0], color='blue', linestyle='-', linewidth=2.5, label='━ Продольный рез (первичный)'),
        Line2D([0], [0], color='orange', linestyle='-', linewidth=2.0, label='━ Продольный рез (вторичный)'),
        Line2D([0], [0], color='red', linestyle='--', linewidth=2.5, label='┊ Поперечный рез (по длине)'),
    ]
    ax_track.legend(handles=legend_patches, loc='upper right', fontsize=9)

    ax_strips.set_xlim(0, 100)
    ax_strips.set_ylim(0, 1)
    ax_strips.axis('off')

    # Формируем сводку с учётом каскадной оптимизации
    from .optimization import OPT_CASCADING_PLAN
    txt = (
        f"Длина по плану: {total_length:.1f} м ({num_tracks} дорожек)  |  Продольных резов: {cfg.LONGITUDINAL_CUTS}  |  Подрезов по длине: {cfg.LENGTH_TRIMS}\n"
        f"Остатки лент 0.3: {cfg.UNUSED_STRIPS_0_3_M_TOTAL:.1f} пог.м  |  Обрезки 0.2: {cfg.SCRAP_STRIPS_0_2_M_TOTAL:.1f} пог.м (≈ {cfg.WASTE_AREA_M2:.2f} м²)"
    )
    
    # Добавляем информацию о каскадной оптимизации, если она была использована
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_plates', 0) > 0:
        txt += f"\n\nОПТИМИЗАЦИЯ: Плит потребуется {OPT_CASCADING_PLAN['total_plates']} шт (с каскадными резами)"
        txt += f" | Отходы: {OPT_CASCADING_PLAN.get('waste_width', 0)} мм"
    
    ax_strips.text(0.02, 0.6, txt, ha='left', va='center', fontsize=12,
                   bbox=dict(boxstyle='round,pad=0.6', facecolor='#f8f9fa', edgecolor='#bdc3c7'))

    # Формируем детальный план резов или остатков
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_plates', 0) > 0:
        # Показываем детальный план каскадных резов
        details = "[PLAN] ДЕТАЛЬНЫЙ ПЛАН РЕЗОВ:\n\n"
        
        # Первичные резы
        if OPT_CASCADING_PLAN.get('primary_cuts'):
            details += "[1] Первичные резы (из 1200 мм):\n"
            for cut in OPT_CASCADING_PLAN['primary_cuts']:
                details += f"  • {cut['qty']} плит → {cut['width']} мм + остаток {cut['rest']} мм\n"
        
        # Вторичные резы
        if OPT_CASCADING_PLAN.get('secondary_cuts'):
            details += "\n[2] Вторичные резы (из остатков):\n"
            for cut in OPT_CASCADING_PLAN['secondary_cuts']:
                if cut.get('pieces', 1) > 1:
                    details += f"  • {cut['qty']} остатков {cut['source']} мм → {cut['pieces']} частей по {cut['cuts'][0]} мм"
                    if cut.get('waste', 0) > 0:
                        details += f" (отход {cut['waste']} мм)"
                    details += "\n"
                else:
                    cuts_str = ' + '.join(str(c) for c in cut['cuts'])
                    details += f"  • {cut['qty']} остатков {cut['source']} мм → {cuts_str} мм"
                    if cut.get('waste', 0) > 0:
                        details += f" (отход {cut['waste']} мм)"
                    details += "\n"
        
        # Поперечные резы
        if OPT_CASCADING_PLAN.get('transverse_cuts'):
            details += "\n[RED] Поперечные резы (по длине):\n"
            for tcut in OPT_CASCADING_PLAN['transverse_cuts']:
                details += f"  • Плита {tcut['source_length']}м x {tcut['source_width']}мм -> {tcut['target_length']}м"
                if tcut.get('remainder', 0) > 0.1:
                    details += f" (остаток {tcut['remainder']:.2f}м)"
                details += "\n"
        
        ax_strips.text(0.02, 0.15, details, ha='left', va='center', fontsize=10,
                       bbox=dict(boxstyle='round,pad=0.5', facecolor='#e8f5e9', edgecolor='#66bb6a'))
    else:
        # Fallback: показываем старую информацию об остатках
        leftovers = (
            "Остатки/обрезки:\n"
            "Ленты 0.3: 3x3.8 м; 1x2.9 м\n"
            f"Ленты 0.2 (обрезки): {', '.join(f'{L:.1f} м' for L in cfg.PLATES_1_0)}"
        )
        ax_strips.text(0.02, 0.15, leftovers, ha='left', va='center', fontsize=11,
                       bbox=dict(boxstyle='round,pad=0.5', facecolor='#eef7ff', edgecolor='#a3c9ff'))

    ax_table.axis('off')

    # Формируем список заказа из реальных данных
    order_list = []
    
    plan_orders = get_orders_from_opt_plan()
    if plan_orders:
        plan_counter = Counter()
        for order in plan_orders:
            key = (round(float(order['length']), 3), int(order['width']))
            plan_counter[key] += order['qty']
        for (length, width_mm), qty in sorted(plan_counter.items(), key=lambda x: (x[0][0], x[0][1])):
            order_list.append(f"Заказ {length:.2f}м×{width_mm}мм: {qty} шт")
    else:
        # Собираем все плиты из заказа (legacy путь через cfg)
        all_orders = []
        for width_mm, plates_list in [
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34), (1080, cfg.PLATES_1_08)
        ]:
            if plates_list:
                length_counts = Counter(plates_list)
                for length, qty in sorted(length_counts.items(), key=lambda x: (-x[0], -x[1])):
                    all_orders.append({
                        'length': length,
                        'width': width_mm,
                        'qty': qty
                    })
        if all_orders:
            for order in all_orders:
                order_list.append(f"Заказ {order['length']:.1f}м×{order['width']}мм: {order['qty']} шт")
        else:
            order_list.append('Заказ не найден')
    
    # Формируем список использования из оптимизации
    used_list = []
    
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_plates', 0) > 0:
        # Итого плит
        used_list.append(f"Плит 1200мм потребуется: {OPT_CASCADING_PLAN['total_plates']} шт")
        
        # Первичные резы
        if OPT_CASCADING_PLAN.get('primary_cuts'):
            primary_info = []
            for cut in OPT_CASCADING_PLAN['primary_cuts']:
                primary_info.append(f"{cut['qty']}x({cut['width']}мм+{cut['rest']}мм)")
            if primary_info:
                used_list.append(f"Первичные резы: {'; '.join(primary_info)}")
        
        # Вторичные резы
        if OPT_CASCADING_PLAN.get('secondary_cuts'):
            secondary_info = []
            for cut in OPT_CASCADING_PLAN['secondary_cuts']:
                if cut.get('pieces', 1) > 1:
                    secondary_info.append(f"{cut['qty']}x{cut['source']}мм->{cut['pieces']}x{cut['cuts'][0]}мм")
                else:
                    secondary_info.append(f"{cut['qty']}x{cut['source']}мм->{cut['cuts'][0]}мм")
            if secondary_info:
                used_list.append(f"Вторичные резы: {'; '.join(secondary_info)}")
        
        # Поперечные резы
        if OPT_CASCADING_PLAN.get('transverse_cuts'):
            trans_count = len(OPT_CASCADING_PLAN['transverse_cuts'])
            used_list.append(f"Поперечных резов: {trans_count}")
        
        # Отходы
        if OPT_CASCADING_PLAN.get('waste_width', 0) > 0:
            used_list.append(f"Отходы: {OPT_CASCADING_PLAN['waste_width']} мм по ширине")
    else:
        # Старый формат, если оптимизация не использовалась
        used_list.append('1.2 без реза: 6.3x2; 3.8x2')
        used_list.append('1.5->1.2: 3.8x3; 2.9x1 (остаток 0.3)')
        used_list.append(f'Резы: продольных {cfg.LONGITUDINAL_CUTS}; подрезов {cfg.LENGTH_TRIMS}')

    rows = max(len(order_list), len(used_list))
    table_rows = []
    for i in range(rows):
        left = order_list[i] if i < len(order_list) else ''
        right = used_list[i] if i < len(used_list) else ''
        table_rows.append([left, right])

    col_labels = ['Список плит по заказу', 'Использовано (с учётом резов) / остатки / обрезки']

    table = ax_table.table(cellText=table_rows, colLabels=col_labels, loc='center', cellLoc='left', colLoc='center')
    table.auto_set_font_size(False)
    # Уменьшаем шрифт если таблица большая
    table_font_size = 8 if len(table_rows) > 15 else 11
    table.set_fontsize(table_font_size)
    table.scale(1, 1.5)

    ax_price.axis('off')
    price_headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Неделя', 'Контрагент', 'Вес(кг)', 'Цена', 'Сумма']
    price_table = ax_price.table(cellText=price_rows, colLabels=price_headers, loc='center', cellLoc='center', colLoc='center')
    price_table.auto_set_font_size(False)
    # Уменьшаем шрифт если таблица большая
    price_font_size = 7 if len(price_rows) > 20 else 10
    price_table.set_fontsize(price_font_size)
    price_table.scale(1, 1.4)
    price_col_idx = price_headers.index('Цена')
    not_priced = any(row[price_col_idx].strip().startswith('0') for row in price_rows)
    
    # Используем стоимость из каскадной оптимизации, если она доступна
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get('total_cost', 0) > 0:
        optimized_cost = OPT_CASCADING_PLAN['total_cost']
        title = f'Итоговая стоимость: {optimized_cost:,.2f} ₽ (оптимизировано)'.replace(',', ' ').replace('.', ',')
    else:
        title = f'Итоговая стоимость: {total_sum:,.2f} ₽'.replace(',', ' ').replace('.', ',')
    
    if not_priced:
        title += ' (внимание: не найдены цены для некоторых позиций — проверьте прайс)'
    ax_price.set_title(title, fontsize=12, pad=10)

    # Детальная разбивка теперь сохраняется в отдельный Excel файл, а не отображается на графике

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    csv_path = os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.csv')
    with open(csv_path, 'w', encoding='utf-8') as f:
        f.write('Список плит по заказу;Использовано (с учётом резов) / остатки / обрезки\n')
        for left, right in table_rows:
            f.write(f'{left};{right}\n')

    if pd is not None:
        try:
            df_v = pd.DataFrame(table_rows, columns=col_labels)
            xlsx_path_v = os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.xlsx')
            df_v.to_excel(xlsx_path_v, index=False)
            print(f'[DEBUG] Ведомость сохранена: {xlsx_path_v}')

            # Смета по дорожке
            df_p = pd.DataFrame(price_rows, columns=price_headers)

            # Добавляем итоговую строку по всему заказу (общая сумма по столбцу "Сумма")
            total_sum_str = f'{total_sum:,.2f}'.replace(',', ' ').replace('.', ',')
            total_row = {
                '№': '',
                'Наименование': 'ИТОГО',
                'Кол-во': '',
                'Ед.': '',
                'Неделя': '',
                'Контрагент': '',
                'Вес(кг)': '',
                'Цена': '',
                'Сумма': total_sum_str,
            }
            df_p = pd.concat([df_p, pd.DataFrame([total_row])], ignore_index=True)

            xlsx_path_p = os.path.join(output_dir, f'Смета_Дорожка_1_{timestamp}.xlsx')
            with pd.ExcelWriter(xlsx_path_p, engine='openpyxl') as writer:
                df_p.to_excel(writer, index=False, sheet_name='Смета')
                df_v.to_excel(writer, index=False, sheet_name='Ведомость')
            print(f'[DEBUG] Смета сохранена: {xlsx_path_p}')
            
            # Сохраняем детальную разбивку компонентов в отдельный Excel файл
            if breakdown_tables:
                breakdown_headers = ['Компонент', 'Расчёт', 'Сумма']
                all_breakdown_rows = []
                for breakdown in breakdown_tables:
                    # Заголовок с наименованием
                    all_breakdown_rows.append([breakdown['name'], '', ''])
                    # Строки таблицы
                    for row in breakdown['rows']:
                        all_breakdown_rows.append(row)
                    # Пустая строка между таблицами
                    all_breakdown_rows.append(['', '', ''])
                
                # Удаляем последнюю пустую строку
                if all_breakdown_rows and all_breakdown_rows[-1] == ['', '', '']:
                    all_breakdown_rows.pop()
                
                df_breakdown = pd.DataFrame(all_breakdown_rows, columns=breakdown_headers)
                xlsx_path_breakdown = os.path.join(output_dir, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
                df_breakdown.to_excel(xlsx_path_breakdown, index=False)
                print(f'[DEBUG] Детальная разбивка сохранена в Excel: {xlsx_path_breakdown}')
            else:
                print('[DEBUG] breakdown_tables пустой - файл детальной разбивки не создан')
        except Exception as e:
            print(f'[ОШИБКА] При сохранении Excel файлов: {e}')
            import traceback
            traceback.print_exc()
            pass
    else:
        print('[ПРЕДУПРЕЖДЕНИЕ] pandas не установлен - Excel файлы не будут созданы!')
        print('[ПРЕДУПРЕЖДЕНИЕ] Установите: pip install pandas openpyxl')

    png_path = os.path.join(output_dir, f'Схема_Дорожка_1_КЗ_{timestamp}.png')
    pdf_path = os.path.join(output_dir, f'Схема_Дорожка_1_КЗ_{timestamp}.pdf')
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(png_path, dpi=300)
    fig.savefig(pdf_path)
    plt.close(fig)

    print('[ГОТОВО] Визуализация и файлы сохранены:')
    print('  PNG:', png_path)
    print('  PDF:', pdf_path)
    print('  CSV:', csv_path)
    if pd is not None:
        print('  XLSX (ведомость):', os.path.join(output_dir, f'Ведомость_Дорожка_1_{timestamp}.xlsx'))
        print('  XLSX (смета):', os.path.join(output_dir, f'Смета_Дорожка_1_{timestamp}.xlsx'))
        if breakdown_tables:
            print('  XLSX (детальная разбивка):', os.path.join(output_dir, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx'))
    return png_path, pdf_path


if __name__ == '__main__':
    visualize_plan()
