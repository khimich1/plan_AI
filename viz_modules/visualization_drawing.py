#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль функций рисования для визуализации:
- Отрисовка сегментов плит
- Отрисовка плит с резами
- Отрисовка поперечных резов
"""
import matplotlib.patches as patches

from core.config.constants import TRACK_WIDTH_M


def _draw_segment(ax, x0: float, length: float, color: str, label: str, y: float = 0.0, height: float = TRACK_WIDTH_M, reinforcement: float = None):
    """Рисует простой сегмент плиты"""
    rect = patches.Rectangle((x0, y), length, height, linewidth=1, edgecolor='black', facecolor=color, alpha=0.85)
    ax.add_patch(rect)
    
    # ✅ Добавляем армирование к метке
    if reinforcement and reinforcement < 999:
        label_with_reinf = f"{label} ({reinforcement:.1f})"
    else:
        label_with_reinf = label
    
    # ✅ Поворачиваем текст на 90° (вдоль плиты)
    ax.text(x0 + length/2, y + height/2, label_with_reinf, ha='center', va='center', fontsize=8, color='white', weight='bold', rotation=90)


def _draw_split_plate(ax, x0: float, length: float, main_w: float, rest_w: float, label_main: str, label_rest: str | None = None, secondary_cuts: list = None, y_base: float = 0.0, reinforcement: float = None):
    """Рисует плиту с продольным резом и возможными вторичными резами в остатке"""
    # Защитная проверка: сумма ширин не должна превышать ширину дорожки (1200мм = 1.2м)
    total_width = main_w + rest_w
    if total_width > TRACK_WIDTH_M + 0.001:
        print(f"[WARNING] Ширина превышена: {main_w:.3f} + {rest_w:.3f} = {total_width:.3f} > {TRACK_WIDTH_M}")
        # Корректируем rest_w чтобы уместиться в дорожку
        rest_w = max(0, TRACK_WIDTH_M - main_w)
    
    # Фон всей плиты (1.2 м)
    rect = patches.Rectangle((x0, y_base), length, TRACK_WIDTH_M, linewidth=1.2, edgecolor='black', facecolor='#ecf0f1', alpha=1.0)
    ax.add_patch(rect)
    
    # Основная часть (зелёная)
    main_rect = patches.Rectangle((x0, y_base), length, main_w, linewidth=0.8, edgecolor='black', facecolor='#2ecc71', alpha=0.9)
    ax.add_patch(main_rect)
    
    # ПРОДОЛЬНЫЙ РЕЗ (по ширине) - СИНЯЯ ГОРИЗОНТАЛЬНАЯ ЛИНИЯ
    ax.plot([x0, x0 + length], [y_base + main_w, y_base + main_w], color='blue', linestyle='-', linewidth=2.5, alpha=0.8, zorder=10)
    
    # ✅ Добавляем армирование к основной метке
    if reinforcement and reinforcement < 999:
        label_main_with_reinf = f"{label_main} ({reinforcement:.1f})"
    else:
        label_main_with_reinf = label_main
    
    # ✅ Поворачиваем текст на 90° (вдоль плиты)
    ax.text(x0 + length/2, y_base + main_w/2, label_main_with_reinf, ha='center', va='center', fontsize=8, color='white', weight='bold', rotation=90)
    
    # Если есть вторичные резы в остатке
    if secondary_cuts and rest_w > 0.02:
        y_offset = y_base + main_w
        for i, sec_cut in enumerate(secondary_cuts):
            sec_w = sec_cut['width']
            sec_label = sec_cut['label']
            
            # ВТОРИЧНЫЙ ПРОДОЛЬНЫЙ РЕЗ - ОРАНЖЕВАЯ ЛИНИЯ (перед сегментом)
            if i == 0:
                # Первый вторичный рез (граница между первичным остатком и вторичным сегментом)
                # Линия только на длину вторичного реза
                first_sec_length = sec_cut.get('target_length', length)
                ax.plot([x0, x0 + first_sec_length], [y_offset, y_offset], color='orange', linestyle='-', linewidth=2.0, alpha=0.8, zorder=10)
            
            # Проверяем, есть ли поперечный рез для этой вторичной плиты
            if sec_cut.get('transverse_cut'):
                # Вторичная плита С поперечным резом
                target_length = sec_cut['target_length']
                remainder = sec_cut.get('remainder', 0)
                
                # Рисуем целевую часть (голубая)
                sec_rect = patches.Rectangle((x0, y_offset), target_length, sec_w, linewidth=0.8, edgecolor='black', facecolor='#3498db', alpha=0.9)
                ax.add_patch(sec_rect)
                # ✅ Поворачиваем текст на 90° (вдоль плиты)
                ax.text(x0 + target_length/2, y_offset + sec_w/2, sec_label, ha='center', va='center', fontsize=7, color='white', weight='bold', rotation=90)
                
                # Остаток по длине (светло-серый)
                if remainder > 0.1:
                    remainder_rect = patches.Rectangle((x0 + target_length, y_offset), remainder, sec_w, 
                                                       linewidth=0.8, edgecolor='gray', facecolor='#bdc3c7', alpha=0.7)
                    ax.add_patch(remainder_rect)
                    if remainder > 0.3:
                        ax.text(x0 + target_length + remainder/2, y_offset + sec_w/2, 
                               f'ост.\n{remainder:.2f}м', ha='center', va='center', fontsize=5, color='#2c3e50')
                
                # КРАСНАЯ ВЕРТИКАЛЬНАЯ ЛИНИЯ - поперечный рез!
                ax.plot([x0 + target_length, x0 + target_length], 
                       [y_offset, y_offset + sec_w],
                       color='red', linestyle='--', linewidth=2.5, alpha=0.8, zorder=10)
            else:
                # Обычный вторичный рез (может быть с укорочением по длине)
                # Проверяем, указана ли целевая длина (для transverse-резов)
                sec_length = sec_cut.get('target_length', length)  # По умолчанию - длина основной плиты
                
                sec_rect = patches.Rectangle((x0, y_offset), sec_length, sec_w, linewidth=0.8, edgecolor='black', facecolor='#3498db', alpha=0.9)
                ax.add_patch(sec_rect)
                # ✅ Поворачиваем текст на 90° (вдоль плиты)
                ax.text(x0 + sec_length/2, y_offset + sec_w/2, sec_label, ha='center', va='center', fontsize=7, color='white', weight='bold', rotation=90)
                
                # Если это укороченная плита, рисуем остаток по длине
                if sec_length < length - 0.1:
                    remainder_length = length - sec_length
                    remainder_rect = patches.Rectangle((x0 + sec_length, y_offset), remainder_length, sec_w,
                                                       linewidth=0.8, edgecolor='gray', facecolor='#bdc3c7', alpha=0.7)
                    ax.add_patch(remainder_rect)
                    if remainder_length > 0.3:
                        ax.text(x0 + sec_length + remainder_length/2, y_offset + sec_w/2,
                               f'ост.\n{remainder_length:.2f}м', ha='center', va='center', fontsize=5, color='#2c3e50')
                    
                    # КРАСНАЯ ВЕРТИКАЛЬНАЯ ЛИНИЯ - поперечный рез!
                    ax.plot([x0 + sec_length, x0 + sec_length],
                           [y_offset, y_offset + sec_w],
                           color='red', linestyle='--', linewidth=2.5, alpha=0.8, zorder=10)
            
            y_offset += sec_w
            
            # Линия между вторичными резами (если их несколько)
            if i < len(secondary_cuts) - 1:
                # Линия только на длину вторичного реза (не на всю основную плиту)
                sec_length = sec_cut.get('target_length', length)
                ax.plot([x0, x0 + sec_length], [y_offset, y_offset], color='orange', linestyle='-', linewidth=2.0, alpha=0.8, zorder=10)
        
        # Остаток (отход) - тёмно-серый
        remaining_w = rest_w - sum(sc['width'] for sc in secondary_cuts)
        if remaining_w > 0.01:
            # Определяем минимальную длину среди всех вторичных резов
            min_sec_length = min(sc.get('target_length', length) for sc in secondary_cuts)
            
            # Линия перед отходом (граница) - только на минимальную длину вторичных резов
            ax.plot([x0, x0 + min_sec_length], [y_offset, y_offset], color='gray', linestyle='--', linewidth=1.5, alpha=0.8, zorder=10)
            waste_rect = patches.Rectangle((x0, y_offset), min_sec_length, remaining_w, linewidth=0.5, edgecolor='gray', facecolor='#95a5a6', alpha=0.7)
            ax.add_patch(waste_rect)
            ax.text(x0 + min_sec_length/2, y_offset + remaining_w/2, f'отход {remaining_w*1000:.0f}мм', ha='center', va='center', fontsize=6, color='white')
    elif label_rest and rest_w > 0.02:
        # Обычный остаток без вторичных резов
        # Линия перед остатком уже нарисована (синяя), добавляем только метку
        ax.text(x0 + length/2, y_base + main_w + rest_w/2, label_rest, ha='center', va='center', fontsize=7, color='#2c3e50')
        # Подпись "остаток"
        ax.text(x0 + length - 0.2, y_base + main_w + rest_w/2, f'остаток\n{rest_w*1000:.0f}мм', ha='right', va='center', fontsize=6, color='#7f8c8d', style='italic')


def _draw_transverse_cut(ax, x0: float, total_length: float, target_length: float, 
                         width: float, label_target: str, remainder_length: float, y_base: float = 0.0, reinforcement: float = None):
    """
    Рисует плиту с поперечным резом (по длине)
    
    ├─────────┬──────┤
    │ 3.31м   │ост.  │
    │ нужна   │0.01м │
    └─────────┴──────┘
         ↑
    поперечный рез (красная вертикальная линия)
    """
    # Фон всей плиты
    rect = patches.Rectangle((x0, y_base), total_length, width, 
                            linewidth=1.2, edgecolor='black', 
                            facecolor='#ecf0f1', alpha=1.0)
    ax.add_patch(rect)
    
    # Левая часть (целевая плита) - зелёная
    target_rect = patches.Rectangle((x0, y_base), target_length, width,
                                   linewidth=0.8, edgecolor='black',
                                   facecolor='#27ae60', alpha=0.9)
    ax.add_patch(target_rect)
    
    # ✅ Добавляем армирование к метке
    if reinforcement and reinforcement < 999:
        label_target_with_reinf = f"{label_target} ({reinforcement:.1f})"
    else:
        label_target_with_reinf = label_target
    
    # ✅ Поворачиваем текст на 90° (вдоль плиты)
    ax.text(x0 + target_length/2, y_base + width/2, label_target_with_reinf,
           ha='center', va='center', fontsize=8, color='white', weight='bold', rotation=90)
    
    # Правая часть (остаток) - светло-серая
    if remainder_length > 0.01:
        remainder_rect = patches.Rectangle((x0 + target_length, y_base), 
                                          remainder_length, width,
                                          linewidth=0.8, edgecolor='gray',
                                          facecolor='#bdc3c7', alpha=0.7)
        ax.add_patch(remainder_rect)
        
        # Метка остатка по длине
        if remainder_length > 0.3:  # Показываем метку только если остаток заметный
            ax.text(x0 + target_length + remainder_length/2, y_base + width/2,
                   f'остаток\nпо длине\n{remainder_length:.2f}м',
                   ha='center', va='center', fontsize=6, color='#2c3e50', weight='bold')
    
    # КРАСНАЯ ВЕРТИКАЛЬНАЯ ЛИНИЯ - поперечный рез!
    ax.plot([x0 + target_length, x0 + target_length], 
           [y_base, y_base + width],
           color='red', linestyle='--', linewidth=2.5, alpha=0.8)

