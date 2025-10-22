#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Оптимизация раскладки плит для дорожек
Минимизация обрезков и разрезов
"""

import itertools
from typing import List, Tuple, Dict
from dataclasses import dataclass
import math

@dataclass
class PlateSize:
    """Размер плиты"""
    width: float
    length: float
    name: str

@dataclass
class CutOption:
    """Вариант разреза плиты"""
    original_plate: PlateSize
    pieces: List[Tuple[float, float]]  # (ширина, длина) для каждого куска
    waste: float  # отходы в процентах
    cut_info: str = ""  # информация о разрезе

class PlateLayoutOptimizer:
    """Оптимизатор раскладки плит"""
    
    def __init__(self):
        self.road_width = 1.2  # ширина дорожки в метрах
        self.road_length = 101  # длина дорожки в метрах
        self.num_roads = 3  # количество дорожек
        self.total_length = self.road_length * self.num_roads
        
        # Стандартные размеры плит ПБ (основная плита 1.2×2.4м)
        self.available_plates = [
            PlateSize(1.2, 2.4, "ПБ 1.2x2.4"),  # стандартная плита ПБ
        ]
        
        # Варианты разрезов (можно дополнить из таблицы резов)
        self.cut_options = self._generate_cut_options()
    
    def _generate_cut_options(self) -> List[CutOption]:
        """Генерирует варианты разрезов плит с точными границами"""
        options = []
        
        # Стандартная ширина ПБ: 1200 мм (1.2 м)
        plate_1_2 = PlateSize(1.2, 2.4, "ПБ 1.2x2.4")
        
        # Точная таблица допустимых резов с границами:
        
        # 1. Рез 300мм (260-320мм): остаток 880-940мм
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.26, 2.4), (0.94, 2.4)],  # минимальные границы
            waste=0.0,
            cut_info="Рез 300мм (260-320мм) -> остаток 880-940мм"
        ))
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.32, 2.4), (0.88, 2.4)],  # максимальные границы
            waste=0.0,
            cut_info="Рез 300мм (260-320мм) -> остаток 880-940мм"
        ))
        
        # 2. Рез 500мм (460-530мм): остаток 670-740мм
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.46, 2.4), (0.74, 2.4)],  # минимальные границы
            waste=0.0,
            cut_info="Рез 500мм (460-530мм) -> остаток 670-740мм"
        ))
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.53, 2.4), (0.67, 2.4)],  # максимальные границы
            waste=0.0,
            cut_info="Рез 500мм (460-530мм) -> остаток 670-740мм"
        ))
        
        # 3. Рез 700мм (660-720мм): остаток 480-540мм
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.66, 2.4), (0.54, 2.4)],  # минимальные границы
            waste=0.0,
            cut_info="Рез 700мм (660-720мм) -> остаток 480-540мм"
        ))
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.72, 2.4), (0.48, 2.4)],  # максимальные границы
            waste=0.0,
            cut_info="Рез 700мм (660-720мм) -> остаток 480-540мм"
        ))
        
        # 4. Рез 900мм (860-920мм): остаток 280-340мм
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.86, 2.4), (0.34, 2.4)],  # минимальные границы
            waste=0.0,
            cut_info="Рез 900мм (860-920мм) -> остаток 280-340мм"
        ))
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.92, 2.4), (0.28, 2.4)],  # максимальные границы
            waste=0.0,
            cut_info="Рез 900мм (860-920мм) -> остаток 280-340мм"
        ))
        
        # 5. Специальные комбинации для плит 0.8м
        # Плита 0.8м -> рез на 0.8м + 0.4м (в пределах допустимых границ)
        options.append(CutOption(
            original_plate=plate_1_2,
            pieces=[(0.8, 2.4), (0.4, 2.4)],  # 800мм + 400мм
            waste=0.0,
            cut_info="Рез 800мм -> остаток 400мм (для плит 0.8м)"
        ))
        
        return options
    
    def find_optimal_layout(self) -> Dict:
        """Находит оптимальную раскладку плит"""
        # Для плиты 1.2×2.4м (ширина точно равна ширине дорожки) - простое решение
        plate_length = 2.4  # длина плиты
        plates_per_road = math.ceil(self.road_length / plate_length)  # плит на дорожку
        total_plates = plates_per_road * self.num_roads
        
        total_length = total_plates * plate_length
        used_area = total_plates * self.road_width * plate_length
        road_area = self.road_width * self.total_length
        
        waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
        
        plate_layout = []
        for i in range(total_plates):
            plate_layout.append({
                'original': 'ПБ 1.2x2.4',
                'width': self.road_width,
                'length': plate_length,
                'cut': False
            })
        
        return {
            'total_length': total_length,
            'used_area': used_area,
            'road_area': road_area,
            'waste_percent': waste_percent,
            'num_cuts': 0,  # разрезов не требуется
            'plate_layout': plate_layout,
            'num_plates': total_plates
        }
    
    def _evaluate_combination(self, plates: List[PlateSize]) -> Dict:
        """Оценивает комбинацию плит с правильной обработкой плит меньше 1.2м"""
        total_length = 0
        total_area = 0
        used_area = 0
        num_cuts = 0
        plate_layout = []
        remaining_parts = []  # Остатки от разрезов
        
        for plate in plates:
            # Для плиты точно подходящей по ширине
            if abs(plate.width - self.road_width) < 0.01:
                # Плита подходит без разреза
                plate_layout.append({
                    'original': plate.name,
                    'width': plate.width,
                    'length': plate.length,
                    'cut': False
                })
                total_length += plate.length
                used_area += plate.width * plate.length
                
            # Для плиты меньше ширины дорожки
            elif plate.width < self.road_width:
                result = self._handle_smaller_plate(plate)
                if result:
                    plate_layout.extend(result['layout'])
                    total_length += result['total_length']
                    used_area += result['used_area']
                    num_cuts += result['cuts']  # Добавляем разрезы для получения плит меньшего размера
                    remaining_parts.extend(result['remaining_parts'])
                else:
                    return None
                    
            # Для плиты больше ширины дорожки
            elif plate.width > self.road_width:
                result = self._handle_larger_plate(plate)
                if result:
                    plate_layout.extend(result['layout'])
                    total_length += result['total_length']
                    used_area += result['used_area']
                    num_cuts += result['cuts']
                    remaining_parts.extend(result['remaining_parts'])
                else:
                    return None
            
            total_area += plate.width * plate.length
        
        # Вычисляем отходы
        road_area = self.road_width * self.total_length
        waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
        
        return {
            'total_length': total_length,
            'used_area': used_area,
            'road_area': road_area,
            'waste_percent': waste_percent,
            'num_cuts': num_cuts,
            'plate_layout': plate_layout,
            'num_plates': len(plates),
            'remaining_parts': remaining_parts
        }
    
    def _find_cut_option(self, plate: PlateSize) -> CutOption:
        """Находит подходящий вариант разреза для плиты"""
        for option in self.cut_options:
            if (option.original_plate.width == plate.width and 
                option.original_plate.length == plate.length):
                return option
        return None
    
    def _handle_smaller_plate(self, plate: PlateSize) -> Dict:
        """
        Обрабатывает плиту меньше ширины дорожки
        Проверяет возможность размещения нескольких плит поперёк или делает рез
        """
        # Проверяем, поместится ли несколько плит поперёк
        plates_across = math.ceil(self.road_width / plate.width)
        total_width = plates_across * plate.width
        
        if total_width <= self.road_width:
            # Обычная раскладка несколькими плитами поперёк
            return self._normal_smaller_layout(plate, plates_across)
        else:
            # Нужен рез плиты
            return self._cut_smaller_plate(plate)
    
    def _normal_smaller_layout(self, plate: PlateSize, plates_across: int) -> Dict:
        """
        Обычная раскладка плит меньшего размера поперёк
        ВАЖНО: Для получения плит меньшего размера нужны исходные плиты большего размера!
        """
        plates_needed = math.ceil(self.road_length / plate.length)
        # Количество исходных плит (большего размера) для получения нужных плит
        original_plates_needed = plates_needed
        total_plates = plates_needed * plates_across
        
        layout = []
        for i in range(plates_needed):
            for j in range(plates_across):
                layout.append({
                    'original': plate.name,
                    'width': plate.width,
                    'length': plate.length,
                    'cut': True,  # ИСПРАВЛЕНИЕ: нужен разрез исходной плиты
                    'position_across': j + 1,
                    'cut_info': f"Рез исходной плиты на {plates_across} части"
                })
        
        return {
            'layout': layout,
            'total_length': plates_needed * plate.length,
            'used_area': total_plates * plate.width * plate.length,
            'cuts': original_plates_needed,  # ИСПРАВЛЕНИЕ: количество разрезов = количество исходных плит
            'remaining_parts': []
        }
    
    def _cut_smaller_plate(self, plate: PlateSize) -> Dict:
        """
        Делает рез плиты меньшего размера с использованием точной таблицы резов
        Пример: плита 0.8м → рез на 0.8м + 0.4м
        """
        # Ищем подходящий вариант разреза из таблицы
        best_cut_option = self._find_best_cut_for_smaller_plate(plate)
        
        if not best_cut_option:
            return None
        
        plates_needed = math.ceil(self.road_length / plate.length)
        
        layout = []
        remaining_parts = []
        
        for i in range(plates_needed):
            # Используем первый кусок из разреза (подходящий по ширине)
            usable_piece = None
            remaining_piece = None
            
            for piece_width, piece_length in best_cut_option.pieces:
                if piece_width <= self.road_width:
                    usable_piece = (piece_width, piece_length)
                else:
                    remaining_piece = (piece_width, piece_length)
            
            if usable_piece:
                layout.append({
                    'original': plate.name,
                    'width': usable_piece[0],
                    'length': usable_piece[1],
                    'cut': True,
                    'cut_info': best_cut_option.cut_info
                })
                
                # Добавляем остаток
                if remaining_piece:
                    remaining_usage = self._check_remaining_part(remaining_piece[0], remaining_piece[1])
                    if remaining_usage['can_use']:
                        remaining_parts.append({
                            'width': remaining_piece[0],
                            'length': remaining_piece[1],
                            'usage': remaining_usage['usage_type'],
                            'original_plate': plate.name,
                            'cut_info': best_cut_option.cut_info
                        })
        
        return {
            'layout': layout,
            'total_length': plates_needed * plate.length,
            'used_area': plates_needed * usable_piece[0] * usable_piece[1] if usable_piece else 0,
            'cuts': plates_needed,
            'remaining_parts': remaining_parts
        }
    
    def _handle_larger_plate(self, plate: PlateSize) -> Dict:
        """Обрабатывает плиту больше ширины дорожки"""
        # Ищем подходящий вариант разреза
        cut_option = self._find_cut_option(plate)
        if cut_option:
            plates_needed = math.ceil(self.road_length / plate.length)
            
            layout = []
            remaining_parts = []
            
            for i in range(plates_needed):
                for piece_width, piece_length in cut_option.pieces:
                    if piece_width <= self.road_width:
                        layout.append({
                            'original': plate.name,
                            'width': piece_width,
                            'length': piece_length,
                            'cut': True,
                            'cut_info': f"Рез: {piece_width}м"
                        })
                    else:
                        # Остаток больше ширины дорожки
                        remaining_parts.append({
                            'width': piece_width,
                            'length': piece_length,
                            'usage': 'waste',
                            'original_plate': plate.name
                        })
            
            return {
                'layout': layout,
                'total_length': plates_needed * plate.length,
                'used_area': plates_needed * self.road_width * plate.length,
                'cuts': plates_needed,
                'remaining_parts': remaining_parts
            }
        else:
            return None
    
    def _find_best_cut_for_smaller_plate(self, plate: PlateSize) -> CutOption:
        """
        Находит лучший вариант разреза для плиты меньшего размера
        """
        best_option = None
        min_waste = float('inf')
        
        for cut_option in self.cut_options:
            # Проверяем, подходит ли этот вариант разреза
            for piece_width, piece_length in cut_option.pieces:
                if abs(piece_width - plate.width) < 0.01:  # Нашли подходящую ширину
                    if cut_option.waste < min_waste:
                        best_option = cut_option
                        min_waste = cut_option.waste
                    break
        
        return best_option
    
    def _check_remaining_part(self, width: float, length: float) -> Dict:
        """
        Проверяет, можно ли использовать остаток от разреза
        """
        # Здесь должна быть проверка в базе данных заказов
        # Пока возвращаем базовую логику
        
        if width >= 0.3:  # Минимальная ширина для использования
            return {
                'can_use': True,
                'usage_type': 'check_orders'  # Нужно проверить в заказах
            }
        else:
            return {
                'can_use': False,
                'usage_type': 'waste'
            }
    
    def add_cut_option_from_table(self, plate_width: float, plate_length: float, 
                                 pieces: List[Tuple[float, float]], waste: float = 0.0):
        """Добавляет вариант разреза из таблицы резов"""
        plate = PlateSize(plate_width, plate_length, f"{plate_width}x{plate_length}")
        cut_option = CutOption(plate, pieces, waste)
        self.cut_options.append(cut_option)
        
        # Также добавляем плиту в доступные размеры, если её там нет
        if not any(p.width == plate_width and p.length == plate_length 
                  for p in self.available_plates):
            self.available_plates.append(plate)
    
    def print_solution(self, solution: Dict):
        """Выводит решение на экран"""
        if not solution:
            print("Решение не найдено")
            return
        
        print("=== ОПТИМАЛЬНАЯ РАСКЛАДКА ПЛИТ ===")
        print(f"Общая длина дорожек: {self.total_length} м")
        print(f"Ширина дорожки: {self.road_width} м")
        print(f"Количество дорожек: {self.num_roads}")
        print()
        print(f"Использовано плит: {solution['num_plates']}")
        print(f"Количество разрезов: {solution['num_cuts']}")
        print(f"Процент отходов: {solution['waste_percent']:.2f}%")
        print()
        
        print("Детальная раскладка:")
        print("-" * 50)
        for i, plate in enumerate(solution['plate_layout'], 1):
            cut_info = " (разрез)" if plate['cut'] else ""
            position_info = f" [поз.{plate.get('position_across', 1)}]" if 'position_across' in plate else ""
            cut_details = f" - {plate.get('cut_info', '')}" if plate.get('cut_info') else ""
            print(f"{i:2d}. {plate['original']} -> {plate['width']:.2f}x{plate['length']:.1f}{cut_info}{position_info}{cut_details}")
        
        # Показываем остатки от разрезов
        if solution.get('remaining_parts'):
            print("\nОстатки от разрезов:")
            print("-" * 50)
            for i, part in enumerate(solution['remaining_parts'], 1):
                usage_info = f" ({part['usage']})" if part.get('usage') else ""
                print(f"{i:2d}. {part['width']:.1f}x{part['length']:.1f}м от {part['original_plate']}{usage_info}")

def main():
    """Основная функция"""
    optimizer = PlateLayoutOptimizer()
    
    print("=== ОПТИМИЗАЦИЯ РАСКЛАДКИ ПЛИТ ПБ ===")
    print(f"Ширина дорожки: {optimizer.road_width} м")
    print(f"Длина дорожки: {optimizer.road_length} м") 
    print(f"Количество дорожек: {optimizer.num_roads}")
    print(f"Общая длина: {optimizer.total_length} м")
    print()
    
    print("Доступные варианты разрезов плиты 1.2x2.4м:")
    for i, option in enumerate(optimizer.cut_options, 1):
        pieces_str = " + ".join([f"{w:.2f}x{l:.1f}" for w, l in option.pieces])
        print(f"{i}. {pieces_str} (отходы: {option.waste*100:.0f}%)")
        if option.cut_info:
            print(f"   {option.cut_info}")
    print()
    
    # Тестируем исправленный алгоритм с разными плитами
    print("=== ТЕСТ: Плита 0.6x2.4м (2 плиты поперёк) ===")
    test_plate_1 = PlateSize(0.6, 2.4, "ПБ 0.6x2.4")
    result_1 = optimizer._handle_smaller_plate(test_plate_1)
    
    if result_1:
        print(f"[УСПЕХ] Результат для плиты {test_plate_1.name}:")
        print(f"   Плит нужно: {len(result_1['layout'])}")
        print(f"   Разрезов: {result_1['cuts']} (для получения плит 0.6м из плит 1.2м)")
        print(f"   Остатков: {len(result_1['remaining_parts'])}")
    
    print("\n=== ТЕСТ: Плита 0.8x5.3м (нужен рез) ===")
    test_plate_2 = PlateSize(0.8, 5.3, "ПБ 0.8x5.3")
    result_2 = optimizer._handle_smaller_plate(test_plate_2)
    
    if result_2:
        print(f"[УСПЕХ] Результат для плиты {test_plate_2.name}:")
        print(f"   Плит нужно: {len(result_2['layout'])}")
        print(f"   Разрезов: {result_2['cuts']}")
        print(f"   Остатков: {len(result_2['remaining_parts'])}")
        
        if result_2['remaining_parts']:
            print("   Остатки:")
            for part in result_2['remaining_parts']:
                print(f"     - {part['width']:.1f}x{part['length']:.1f}м ({part['usage']})")
    else:
        print("[ОШИБКА] Не удалось обработать плиту")
    
    print("\n" + "="*60)
    print("Поиск оптимального решения...")
    solution = optimizer.find_optimal_layout()
    
    optimizer.print_solution(solution)
    
    print("\n" + "="*60)
    print("Таблица резов ПБ (стандартная ширина 1200 мм):")
    print("300мм -> остаток 900мм")
    print("500мм -> остаток 700мм") 
    print("700мм -> остаток 500мм")
    print("900мм -> остаток 300мм")
    print("1020-1080мм -> остаток 120-180мм (утилизация)")
    print()
    print("ИСПРАВЛЕНИЯ:")
    print("[+] Плиты < 1.2м: проверка возможности размещения поперёк")
    print("[+] Если не помещается: рез плиты с проверкой остатков")
    print("[+] Остатки: проверка в заказах или утилизация")

if __name__ == "__main__":
    main()
