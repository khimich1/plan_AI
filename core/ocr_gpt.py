#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Умное распознавание текста с изображений:
- Сначала пробуем EasyOCR (бесплатно, быстро)
- Если не получилось — используем GPT-4o Vision (платно, но точно)

Автор: AI Assistant для новичка в Python
Дата: 2025-11-27
"""

import os
import base64
import json
import re
from typing import Optional, Dict, List

try:
    from openai import AsyncOpenAI
    GPT_AVAILABLE = True
except ImportError:
    GPT_AVAILABLE = False
    print("[GPT OCR] ⚠️ OpenAI не установлен. Установите: pip install openai")

# Импорт существующего EasyOCR модуля
try:
    from .ocr_recognition import (
        recognize_text_from_image, 
        clean_recognized_text,
        EASYOCR_AVAILABLE
    )
except ImportError:
    EASYOCR_AVAILABLE = False
    print("[GPT OCR] ⚠️ EasyOCR недоступен")


async def recognize_text_smart(
    image_path: str, 
    force_gpt: bool = False,
    show_cost: bool = True
) -> Optional[Dict]:
    """
    🧠 УМНОЕ РАСПОЗНАВАНИЕ: EasyOCR → GPT fallback
    
    Как работает:
    1. Сначала пробуем бесплатный EasyOCR
    2. Если он нашёл хотя бы 1 плиту — используем его результат
    3. Если не нашёл или ошибка — подключаем платный GPT-4o
    
    Аргументы:
        image_path: путь к файлу изображения (jpg, png)
        force_gpt: True = сразу использовать GPT (игнорируя EasyOCR)
        show_cost: показывать стоимость в консоли
        
    Возвращает:
        {
            'text': str,           # Текст для парсера (в формате "ПБ XX-XX-Xп qty")
            'plates': list,        # Список плит [{name, qty}] (только для GPT)
            'method': str,         # 'EasyOCR' или 'GPT-4o'
            'confidence': float,   # Уверенность 0.0-1.0
            'cost_usd': float      # Стоимость в $ (только для GPT)
        }
        или None если не удалось распознать
    """
    
    # ============ ЭТАП 1: Пробуем EasyOCR (бесплатно!) ============
    if not force_gpt and EASYOCR_AVAILABLE:
        try:
            print("[OCR] 🤖 Пробую EasyOCR (бесплатно)...")
            text = recognize_text_from_image(image_path)
            
            if text:
                cleaned = clean_recognized_text(text)
                
                # Проверяем качество: ищем плиты в формате "ПБ XX-XX-Xп"
                # Например: "ПБ 78-12-8п", "ПБ 66,2-12-8п", "ПБ 44-3,2-10п"
                plates_found = re.findall(
                    r'П[БК]\s*\d+[,\.]?\d*\s*-\s*\d+[,\.]?\d*\s*-\s*\d+[,\.]?\d*п',
                    cleaned,
                    re.IGNORECASE
                )
                
                if len(plates_found) >= 1:
                    print(f"[OCR] ✅ EasyOCR распознал {len(plates_found)} плит(ы)")
                    return {
                        'text': cleaned,
                        'plates': [],  # EasyOCR возвращает просто текст
                        'method': 'EasyOCR',
                        'confidence': min(len(plates_found) / 5, 1.0),  # Макс 1.0 при 5+ плитах
                        'cost_usd': 0.0
                    }
                
                print(f"[OCR] ⚠️ EasyOCR нашёл только {len(plates_found)} плит, пробую GPT...")
        
        except Exception as e:
            print(f"[OCR] ❌ Ошибка EasyOCR: {e}")
            print("[OCR] Переключаюсь на GPT...")
    
    # ============ ЭТАП 2: Используем GPT-4o Vision ============
    if not GPT_AVAILABLE:
        print("[OCR] ❌ GPT недоступен. Установите: pip install openai")
        return None
    
    try:
        print("[OCR] 🧠 Запускаю GPT-4o Vision (платно)...")
        plates, cost = await recognize_with_gpt_vision(image_path)
        
        if plates:
            # Конвертируем список плит в текст для парсера
            text_lines = [f"{p['name']} {p['qty']}" for p in plates]
            text = '\n'.join(text_lines)
            
            if show_cost:
                rub_cost = cost * 75  # Примерный курс ₽
                print(f"[OCR] 💰 Стоимость: ${cost:.4f} (~{rub_cost:.2f}₽)")
            
            print(f"[OCR] ✅ GPT распознал {len(plates)} плит(ы)")
            return {
                'text': text,
                'plates': plates,
                'method': 'GPT-4o',
                'confidence': 0.95,  # GPT обычно очень точен
                'cost_usd': cost
            }
    
    except Exception as e:
        print(f"[OCR] ❌ Ошибка GPT: {e}")
        import traceback
        traceback.print_exc()
    
    return None


async def recognize_with_gpt_vision(image_path: str) -> tuple[List[Dict], float]:
    """
    Распознавание через GPT-4o Vision API
    
    Простыми словами:
    1. Читаем картинку и превращаем в base64 (текстовый формат)
    2. Отправляем GPT с умным промптом: "Найди все плиты в таблице"
    3. GPT возвращает JSON со списком плит
    4. Считаем примерную стоимость запроса
    
    Возвращает:
        (список_плит, стоимость_в_долларах)
        
    Пример списка плит:
        [
            {"name": "ПБ 78-12-8п", "qty": 4},
            {"name": "ПБ 66,2-12-8п", "qty": 6}
        ]
    """
    
    # Проверяем наличие API ключа
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        raise ValueError(
            "❌ Не найден OPENAI_API_KEY!\n\n"
            "Как исправить:\n"
            "1. Создай файл .env в корне проекта\n"
            "2. Добавь строку: OPENAI_API_KEY=sk-твой-ключ\n"
            "3. Получить ключ можно на https://platform.openai.com/api-keys"
        )
    
    # Создаём клиент OpenAI
    client = AsyncOpenAI(api_key=api_key)
    
    # Читаем изображение и кодируем в base64
    with open(image_path, "rb") as f:
        image_data = f.read()
        image_base64 = base64.b64encode(image_data).decode()
    
    # Определяем размер для расчёта стоимости
    image_size_kb = len(image_data) / 1024
    print(f"[GPT] Размер изображения: {image_size_kb:.1f} КБ")
    
    # Отправляем запрос к GPT-4o
    response = await client.chat.completions.create(
        model="gpt-4o",  # Модель с поддержкой изображений
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text", 
                        "text": get_recognition_prompt()  # Умный промпт ниже
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{image_base64}",
                            "detail": "high"  # Высокое качество для таблиц
                        }
                    }
                ]
            }
        ],
        max_tokens=2000,  # Максимум токенов в ответе
        temperature=0.1   # Низкая = более точные ответы
    )
    
    # Парсим ответ GPT
    result_text = response.choices[0].message.content
    plates = parse_gpt_response(result_text)
    
    # Считаем стоимость
    # GPT-4o: $2.50 за 1M входящих токенов
    tokens_used = response.usage.total_tokens
    cost_usd = (tokens_used / 1_000_000) * 2.5
    
    print(f"[GPT] Использовано токенов: {tokens_used}")
    
    return plates, cost_usd


def get_recognition_prompt() -> str:
    """
    Промпт для GPT — это инструкция, как распознавать плиты.
    
    Чем точнее промпт, тем лучше результат!
    """
    return """Ты эксперт по распознаванию заказов железобетонных плит из таблиц.

📋 ФОРМАТ ПЛИТ:
- Наименование: "Плиты ПБ 78-12-8п" или "ПБ 78-12-8п"
- Структура: ПБ {длина_дм}-{ширина_дм}-{нагрузка}
- Пример: "ПБ 66,2-12-8п" означает:
  • Длина: 6.62 метра (66,2 дециметра)
  • Ширина: 1.2 метра (12 дециметров)
  • Нагрузка: 8п (800 кг/м²)

🎯 ЗАДАЧА:
Извлеки ВСЕ плиты из таблицы в JSON формате:

[
  {"name": "ПБ 78-12-8п", "qty": 4},
  {"name": "ПБ 66,2-12-8п", "qty": 6},
  {"name": "ПБ 44-3,2-8п", "qty": 5}
]

⚠️ ВАЖНЫЕ ПРАВИЛА:
1. Формат названия: СТРОГО "ПБ XX-XX-Xп" (без слова "Плиты")
2. Нагрузка: обязательно с буквой "п": 6п, 8п, 10п, 12п, 12,5п
3. Количество (qty): число из колонки "Кол-во", обычно от 1 до 99
4. ИГНОРИРУЙ размеры в миллиметрах (например 7980x1190x220) — это НЕ количество!
5. Сохраняй запятые в русском формате (66,2 а не 66.2)
6. Пропускай строки с заголовками ("Наименование", "Кол-во", "Итого")
7. Если в строке непонятно — лучше пропусти её

✅ Верни ТОЛЬКО JSON массив, без текста до и после!"""


def parse_gpt_response(response_text: str) -> List[Dict]:
    """
    Извлекает JSON из ответа GPT.
    
    GPT может вернуть:
    - Чистый JSON: [{"name": "...", "qty": 4}]
    - Обёрнутый в блок: ```json [...] ```
    - С комментариями: "Вот результат: [...]"
    
    Эта функция находит JSON в любом случае.
    """
    
    # Пытаемся найти JSON в блоке ```json ... ```
    json_match = re.search(
        r'```(?:json)?\s*(\[.*?\])\s*```',
        response_text,
        re.DOTALL  # Флаг для многострочного поиска
    )
    
    if json_match:
        response_text = json_match.group(1)
    
    try:
        # Парсим JSON
        plates = json.loads(response_text)
        
        # Валидация: проверяем структуру каждой плиты
        validated_plates = []
        for plate in plates:
            if not isinstance(plate, dict):
                print(f"[GPT] ⚠️ Пропущена плита (не объект): {plate}")
                continue
            
            if 'name' not in plate or 'qty' not in plate:
                print(f"[GPT] ⚠️ Пропущена плита (нет name/qty): {plate}")
                continue
            
            # Приводим qty к целому числу
            try:
                plate['qty'] = int(plate['qty'])
            except (ValueError, TypeError):
                print(f"[GPT] ⚠️ Пропущена плита (qty не число): {plate}")
                continue
            
            validated_plates.append(plate)
        
        return validated_plates
    
    except json.JSONDecodeError as e:
        print(f"[GPT] ❌ Ошибка парсинга JSON: {e}")
        print(f"[GPT] Ответ GPT (первые 200 символов):")
        print(response_text[:200])
        return []


# ============ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ============

def estimate_monthly_cost(photos_per_month: int) -> Dict[str, float]:
    """
    Оценка месячных затрат на OCR
    
    Аргументы:
        photos_per_month: сколько фото в месяц обрабатываете
        
    Возвращает:
        {'easyocr': 0, 'gpt_only': X, 'hybrid': Y}
    """
    
    # Средняя стоимость одного фото через GPT-4o
    avg_cost_per_photo = 0.002  # $0.002 = ~0.15₽
    
    # Гибридный подход: 80% решает EasyOCR, 20% — GPT
    hybrid_ratio = 0.2
    
    return {
        'easyocr_only': 0.0,  # Всегда бесплатно
        'gpt_only': photos_per_month * avg_cost_per_photo,
        'hybrid': photos_per_month * avg_cost_per_photo * hybrid_ratio,
        'photos': photos_per_month
    }


if __name__ == '__main__':
    # Пример расчёта стоимости
    print("💰 Оценка месячных затрат на OCR:")
    print("=" * 50)
    
    for count in [100, 500, 1000, 5000]:
        costs = estimate_monthly_cost(count)
        print(f"\n📊 {count} фото в месяц:")
        print(f"  • EasyOCR: $0 (всегда бесплатно)")
        print(f"  • GPT-4o: ${costs['gpt_only']:.2f} (~{costs['gpt_only']*75:.0f}₽)")
        print(f"  • Гибрид: ${costs['hybrid']:.2f} (~{costs['hybrid']*75:.0f}₽) ⭐ рекомендую")

