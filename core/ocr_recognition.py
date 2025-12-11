#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для распознавания текста с изображений плит
Использует EasyOCR - простой и надёжный вариант для русского текста
"""
import os
import re
from typing import Optional
import numpy as np
from PIL import Image

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False
    print("[OCR] ⚠️ EasyOCR не установлен. Установите: pip install easyocr")

# Глобальный объект читателя (создаётся один раз)
_reader = None


def get_reader():
    """Создаёт или возвращает готовый объект EasyOCR читателя"""
    global _reader
    if _reader is None:
        print("[OCR] Инициализация EasyOCR (первый запуск загрузит модели ~100 МБ)...")
        try:
            # Создаём читатель для русского и английского языков
            # gpu=False - используем CPU (проще на Windows)
            _reader = easyocr.Reader(['ru', 'en'], gpu=False)
            print("[OCR] ✅ EasyOCR готов к работе!")
        except Exception as e:
            print(f"[OCR] ❌ Ошибка инициализации EasyOCR: {e}")
            raise
    return _reader


def process_table_row(blocks: list, avg_x: float) -> str:
    """
    Обрабатывает строку таблицы, определяя структуру колонок.
    Пытается найти наименование плиты и количество.
    
    Args:
        blocks: список блоков текста в одной строке
        avg_x: средняя X-координата для определения колонок
    
    Returns:
        Строка в формате "ПБ 44-12-8п 4"
    """
    if not blocks:
        return ""
    
    # Сортируем блоки слева направо
    blocks.sort(key=lambda b: b['x'])
    
    # Разделяем на левую и правую части (колонки)
    # Определяем границу между колонками (обычно ~70% от левого края)
    max_x = max(b['x'] for b in blocks)
    threshold_x = max_x * 0.6  # 60% - примерная граница между "Наименование" и "Кол-во"
    
    left_blocks = [b for b in blocks if b['x'] < threshold_x]
    right_blocks = [b for b in blocks if b['x'] >= threshold_x]
    
    # Собираем текст из левой части (наименование)
    left_text = ' '.join(b['text'] for b in left_blocks)
    
    # Из правой части извлекаем количество
    # ВАЖНО: берём ПОСЛЕДНЕЕ маленькое число (1-99), игнорируя размеры (>100)
    quantity_text = ""
    potential_quantities = []
    
    for b in right_blocks:
        import re
        numbers = re.findall(r'\d+', b['text'])
        for num in numbers:
            num_int = int(num)
            # Количество плит обычно 1-99, размеры >100
            if 1 <= num_int <= 99:
                potential_quantities.append(num)
    
    # Берём ПОСЛЕДНЕЕ подходящее число (оно обычно и есть количество)
    if potential_quantities:
        quantity_text = potential_quantities[-1]
    
    # Формируем итоговую строку
    if quantity_text:
        return f"{left_text} {quantity_text}"
    else:
        return left_text


def recognize_text_from_image(image_path: str) -> Optional[str]:
    """
    Распознаёт текст с изображения и возвращает его как строку.
    
    Args:
        image_path: Путь к файлу изображения
        
    Returns:
        Распознанный текст или None, если не удалось
    """
    if not EASYOCR_AVAILABLE:
        print("[OCR] ❌ EasyOCR не доступен")
        return None
    
    if not os.path.exists(image_path):
        print(f"[OCR] ❌ Файл не найден: {image_path}")
        return None
    
    print(f"[OCR] Распознаю файл: {image_path}")
    
    try:
        reader = get_reader()
        
        # Загружаем изображение через PIL (поддерживает русские пути!)
        # Затем конвертируем в numpy array для EasyOCR
        image = Image.open(image_path)
        # Конвертируем PIL Image в numpy array (RGB)
        image_array = np.array(image)
        
        # Распознаём текст
        # Результат: список кортежей [(bbox, text, confidence), ...]
        results = reader.readtext(image_array)
        
        print(f"[OCR] Результат распознавания: найдено {len(results)} блоков текста")
        
        if not results:
            print("[OCR] ⚠️ EasyOCR не нашёл текст на изображении")
            return None
        
        # Группируем блоки текста по строкам (по Y-координате)
        # bbox формат: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        blocks_with_coords = []
        for idx, (bbox, text, confidence) in enumerate(results):
            if confidence > 0.3:
                # Берём среднюю Y-координату блока
                y_coords = [point[1] for point in bbox]
                avg_y = sum(y_coords) / len(y_coords)
                # Берём левую X-координату (начало строки)
                x_left = bbox[0][0]
                
                blocks_with_coords.append({
                    'text': text.strip(),
                    'y': avg_y,
                    'x': x_left,
                    'confidence': confidence
                })
                print(f"[OCR] Блок {idx+1}: '{text}' (Y={avg_y:.1f}, X={x_left:.1f}, conf={confidence:.2f})")
            else:
                print(f"[OCR] Блок {idx+1}: '{text}' (пропущен, низкая уверенность {confidence:.2f})")
        
        if not blocks_with_coords:
            print("[OCR] ⚠️ Нет блоков с достаточной уверенностью")
            return None
        
        # Сортируем блоки только по Y (сверху вниз)
        blocks_with_coords.sort(key=lambda b: b['y'])
        
        # Анализируем структуру таблицы по X-координатам
        # Определяем, где находятся колонки (например, "Наименование" и "Кол-во")
        all_x_coords = [b['x'] for b in blocks_with_coords]
        avg_x = sum(all_x_coords) / len(all_x_coords) if all_x_coords else 0
        
        # Группируем блоки в строки по близким Y-координатам
        y_threshold = 15  # Порог для объединения блоков в одну строку (пикселей)
        grouped_lines_with_structure = []
        current_line_blocks = []
        current_y = blocks_with_coords[0]['y']
        
        for block in blocks_with_coords:
            # Если блок на той же строке (Y-координата близка)
            if abs(block['y'] - current_y) < y_threshold:
                current_line_blocks.append(block)
            else:
                # Обрабатываем завершённую строку
                if current_line_blocks:
                    processed_line = process_table_row(current_line_blocks, avg_x)
                    if processed_line:
                        grouped_lines_with_structure.append(processed_line)
                current_line_blocks = [block]
                current_y = block['y']
        
        # Добавляем последнюю строку
        if current_line_blocks:
            processed_line = process_table_row(current_line_blocks, avg_x)
            if processed_line:
                grouped_lines_with_structure.append(processed_line)
        
        lines = grouped_lines_with_structure
        
        print(f"[OCR] Объединено в {len(lines)} строк(и)")
        for i, line in enumerate(lines, 1):
            print(f"[OCR] Строка {i}: {line}")
        
        # Объединяем все строки
        full_text = '\n'.join(lines)
        
        print(f"[OCR] Итоговый текст ({len(full_text)} символов):")
        print(f"[OCR] ---\n{full_text}\n[OCR] ---")
        
        # Очищаем текст от похожих символов
        full_text = full_text.replace('×', 'x').replace('х', 'x')
        full_text = full_text.replace('—', '-').replace('–', '-')
        
        return full_text if full_text else None
        
    except Exception as e:
        print(f"[OCR] ❌ Ошибка при распознавании: {e}")
        import traceback
        traceback.print_exc()
        return None


def clean_recognized_text(text: str) -> str:
    """
    Очищает распознанный текст от типичных ошибок OCR.
    
    Например:
    - "ПБ 78-12-8п" может распознаться как "ПБ 78-12-8р" (п → р)
    - "1.2×3.39" может стать "1.2x3.39" или "1,2x3,39"
    """
    if not text:
        return ""
    
    print(f"[OCR] Очистка текста, исходная длина: {len(text)} символов")
    
    # Убираем заголовки таблиц и служебные символы
    # "Ng (,) :", "№", "Товары" и т.п.
    text = re.sub(r'^[Nn]g\s*\([^\)]*\)\s*:\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^№\s+', '', text, flags=re.MULTILINE)
    
    # Убираем строки с заголовками таблиц
    lines = text.split('\n')
    filtered_lines = []
    for line in lines:
        line_lower = line.lower()
        # Пропускаем строки-заголовки
        if any(word in line_lower for word in ['наименование', 'кол-во', 'товары', 'работы', 'услуги', 'цена', 'сумма', 'итого']):
            # Но оставляем, если в строке есть маркировка плиты (ПБ/ПК + цифры)
            if not re.search(r'п[бк]\s*\d+', line_lower):
                continue
        filtered_lines.append(line)
    text = '\n'.join(filtered_lines)
    
    # Убираем слова "Товары", "работы", "услуги" (заголовки таблиц)
    text = re.sub(r'\b(товары|работы|услуги|кол-во|ед\.?)\b', '', text, flags=re.IGNORECASE)
    
    # Убираем слово "железобетонная" (лишнее из таблиц)
    text = text.replace('железобетонная', '').replace('железобетонный', '')
    
    # "петли" → убираем (лишнее слово из таблиц)
    text = text.replace('петли', '').replace('петлн', '').replace('петл', '')
    
    # Убираем размеры в любых скобках - это размеры в мм
    # (7980x1190x220), [6000x665x220], [5080x395x220 и т.п.
    text = re.sub(r'[\(\[][^\)\]]*\d{3,}x\d+x\d+[^\)\]]*', '', text)
    # Убираем оставшиеся скобки
    text = text.replace('[', '').replace(']', '').replace('(', '').replace(')', '')
    
    # Убираем размеры без скобок: паттерны типа "7980x1190x220" или "980x600x220"
    # Критерий: первое число >= 100 (чтобы не удалить маркировку типа "80-12-8")
    text = re.sub(r'\s+\d{3,}x\d+x\d+', ' ', text)
    
    # Убираем паттерны с пробелами внутри: "7980x11 0x220"
    text = re.sub(r'\s+\d{3,}x\d+\s+\d+x\d+', ' ', text)
    
    # Исправляем OCR ошибки в числах:
    # "44,C" или "44,С" → "44,0" (C/С вместо 0)
    text = re.sub(r'(\d+),([CС])\b', r'\1,0', text)
    # "56,С" → "56,0"
    text = re.sub(r'(\d+),([CС])', r'\1,0', text)
    
    # Исправляем паттерн "44,0 0-12-8" → "44,0-12-8"
    # (убираем лишний "0-" когда идёт после числа)
    text = re.sub(r'(\d+[,\.]\d+)\s+0-(\d+-\d+)', r'\1-\2', text)
    
    # Убираем лишние символы типа "ё", "/", которые OCR добавляет случайно
    text = text.replace('ё', '').replace('/', '')
    
    # "ШТ" → "шт" (OCR часто распознаёт заглавными)
    text = re.sub(r'\bШТ\b', 'шт', text)
    text = re.sub(r'\bШ[ТУ]\b', 'шт', text)
    
    # Заменяем запятые на точки в числах перед дефисами (49,0-12-8 → 49.0-12-8)
    # Паттерн: цифры, запятая, цифры, дефис
    text = re.sub(r'(\d+),(\d+)(?=\s*-)', r'\1.\2', text)
    
    # Добавляем маркер нагрузки "-8п" если его нет
    # Паттерн: "ПБ 44.0-12-8  4" → "ПБ 44.0-12-8п 4"
    # Ищем: ПБ + число-число-число + пробелы + число(количество)
    # Если после третьего числа нет "п", добавляем
    text = re.sub(r'(ПБ\s+\d+[,\.]\d+-\d+-\d+)(\s+\d+)', r'\1п\2', text, flags=re.IGNORECASE)
    
    # Исправляем частые ошибки OCR
    # "8р" → "8п" (если перед этим есть дефис и число)
    text = re.sub(r'-\s*(\d+)\s*р\b', r'-\1п', text)
    
    # "ПБ" может стать "ПВ" или "ПГ" → исправляем
    text = re.sub(r'\bП[ВГ]\s+', 'ПБ ', text)
    
    # Убираем номера строк в начале (1, 2, 3... перед наименованием)
    text = re.sub(r'^\s*\d{1,2}\s+(?=Плита|П[БК])', '', text, flags=re.MULTILINE)
    
    # Разделяем строки по паттернам:
    # 1. "...8 шт 2 Плита..." → "...8 шт\n2 Плита..."
    text = re.sub(r'(шт)\s+(\d{1,2})\s+(Плита)', r'\1\n\2 \3', text, flags=re.IGNORECASE)
    
    # 2. "...8 шт 2 ПБ..." → "...8 шт\nПБ..."  (убираем номер строки)
    text = re.sub(r'(шт)\s+\d{1,2}\s+(ПБ\s+\d)', r'\1\n\2', text, flags=re.IGNORECASE)
    
    # 3. "ПБ 78,3-10,2-8п 4 ПБ 67..." → "ПБ 78,3-10,2-8п 4\nПБ 67..."
    # Разделяем по паттерну: количество + пробелы + "ПБ" + цифры
    text = re.sub(r'(\d{1,3})\s+(П[БК]\s+\d)', r'\1\n\2', text)
    
    # 4. Также разделяем по паттерну: буква/цифра + номер строки + "Плита железобетонная"
    text = re.sub(r'([а-яА-Я]|шт|\d)\s+(\d{1,2})\s+(Плита\s+железобетонная)', r'\1\n\2 \3', text, flags=re.IGNORECASE)
    
    # Убираем лишние пробелы и пустые строки
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    text = '\n'.join(lines)
    
    print(f"[OCR] Очищенный текст, длина: {len(text)} символов")
    print(f"[OCR] Строк после очистки: {len(lines)}")
    
    return text
