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
                
                # Проверяем качество: ищем плиты в стандартном или каталожном форматах.
                # Стандартный: "ПБ 78-12-8п", "ПБ 66,2-12-8п", "ПБ 44-3,2-10п"
                # Каталожный:  "ПБ 59.12-8Вр1400-25", "ПБ56.05-10"
                plates_found = re.findall(
                    r'П[БК]\s*\d+[,\.]?\d*\s*-\s*\d+[,\.]?\d*\s*-\s*\d+[,\.]?\d*п',
                    cleaned,
                    re.IGNORECASE
                )
                if not plates_found:
                    # Каталожный формат: ПБ L.W-load (без «п» в конце, L >= 2 символов)
                    plates_found = re.findall(
                        r'П[БК]\s*\d{2,}\.\d+\s*-\s*\d+',
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
    return """Ты OCR-система для посимвольного копирования таблиц железобетонных плит.

🎯 ТВОЯ ЗАДАЧА: Переписать таблицу БЕЗ ИНТЕРПРЕТАЦИИ

Работай как КСЕРОКС - копируй символы точно как видишь, не думай о смысле!

📋 Формат вывода - JSON массив:
[
  {"name": "ПБ 66,2-12-8п", "qty": 6},
  {"name": "ПБ 66,2-10,2-8п", "qty": 1},
  {"name": "ПБ 61,0-12-8п", "qty": 2},
  {"name": "ПБ 52,0-7,2-8п", "qty": 1}
]

🔥 КРИТИЧЕСКИ ВАЖНО - ПОСИМВОЛЬНОЕ КОПИРОВАНИЕ:

Представь, что переписываешь таблицу от руки в блокнот.
Ты НЕ математик, ты НЕ думаешь о числах - ты просто КОПИРУЕШЬ текст!

✅ ПРАВИЛЬНО (копируешь ВСЁ что видишь):
• Видишь "Плиты ПБ 66,2-12-8п" → name: "ПБ 66,2-12-8п"
• Видишь "Плиты ПБ 61,0-12-8п" → name: "ПБ 61,0-12-8п"
• Видишь "Плиты ПБ 52,0-7,2-8п" → name: "ПБ 52,0-7,2-8п"
• Видишь "Плиты ПБ 21,5-10,2-8п" → name: "ПБ 21,5-10,2-8п"
• Видишь "Плиты ПБ 38,3-12-8п" → name: "ПБ 38,3-12-8п"
• Видишь "Плиты ПБ 58,4-12-8п" → name: "ПБ 58,4-12-8п"
• Видишь "Плиты ПБ 56-10,8-8п" → name: "ПБ 56-10,8-8п"
• Видишь "ПБ 59.12-8Вр1400-25" → name: "ПБ 59.12-8Вр1400-25" (копируй ДОСЛОВНО с суффиксом!)
• Видишь "ПБ56.05-10" → name: "ПБ56.05-10" (каталожный формат — копируй как есть)

❌ НЕПРАВИЛЬНО (думаешь и упрощаешь):
• "Плиты ПБ 66,2-12-8п" → "ПБ 66-12-8п" (ПОТЕРЯНА ЗАПЯТАЯ И ЦИФРА!)
• "Плиты ПБ 61,0-12-8п" → "ПБ 61-12-8п" (ГДЕ ",0"?)
• "Плиты ПБ 52,0-7,2-8п" → "ПБ 52-7,2-8п" (ГДЕ ПЕРВАЯ ",0"?)
• "Плиты ПБ 21,5-10,2-8п" → "ПБ 22-10,2-8п" (ТЫ ОКРУГЛИЛ! НЕ НАДО!)

📐 ЗАПОМНИ РАЗ И НАВСЕГДА:
• "66,2" ≠ "66" (это РАЗНЫЕ числа!)
• "61,0" ≠ "61" (это РАЗНЫЕ числа!)
• "52,0" ≠ "52" (это РАЗНЫЕ числа!)
• "21,5" ≠ "21" и ≠ "22" (это РАЗНЫЕ числа!)

Да, математически 61,0 = 61, но в нашей системе это КРИТИЧНО РАЗНЫЕ коды плит!
Плита "ПБ 61,0-12-8п" и "ПБ 61-12-8п" - это СОВЕРШЕННО РАЗНЫЕ изделия!

⚙️ ПРАВИЛА:

1. **КАЖДАЯ строка таблицы = отдельный элемент JSON**
   (даже если "ПБ 28-12-8п" повторяется 5 раз - создай 5 элементов!)

2. **Формат name:** убери слово "Плиты" и лишние пробелы:
   "Плиты ПБ 66,2-12-8п" → "ПБ 66,2-12-8п"

3. **Формат qty:** число из правой колонки (обычно 1-99)

4. **Порядок:** сверху вниз как в таблице

5. **Пропускай только заголовки:** "Наименование", "Кол-во", "Итого"

❌ КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО:
• Упрощать: 66,2 → 66
• Округлять: 21,5 → 22
• Убирать нули: 52,0 → 52
• Убирать десятичные: 10,2 → 10
• Группировать повторяющиеся строки
• Думать о смысле чисел (ты OCR, не математик!)

💡 Если сомневаешься - копируй ДОСЛОВНО! Лучше лишняя запятая, чем потерянная!

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

