#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Распознавание текста с изображений через GPT-4o Vision.
"""

import os
import base64
import json
import re
from typing import Optional, Dict, List, Any, Literal

try:
    from openai import AsyncOpenAI
    GPT_AVAILABLE = True
except ImportError:
    GPT_AVAILABLE = False
    print("[GPT OCR] ⚠️ OpenAI не установлен. Установите: pip install openai")


async def recognize_text_smart(
    image_path: str, 
    force_gpt: bool = False,
    show_cost: bool = True,
    mode: Literal["full_gpt", "hybrid"] = "full_gpt",
) -> Optional[Dict]:
    """
    🧠 Распознавание через GPT-4o Vision.
    
    Аргументы:
        image_path: путь к файлу изображения (jpg, png)
        force_gpt: аргумент оставлен для обратной совместимости
        show_cost: показывать стоимость в консоли
        mode: аргумент оставлен для обратной совместимости
        
    Возвращает:
        {
            'text': str,           # Текст для парсера (в формате "ПБ XX-XX-Xп qty")
            'plates': list,        # Список плит [{name, qty}]
            'method': str,         # 'GPT-4o'
            'confidence': float,   # Уверенность 0.0-1.0
            'cost_usd': float      # Стоимость в $
        }
        или None если не удалось распознать
    """
    
    # Аргументы сохраняем ради совместимости со старым API модуля.
    _ = force_gpt
    if mode == "hybrid":
        print("[OCR] ℹ️ Режим hybrid отключен: используется только GPT-4o")

    # ============ Используем GPT-4o Vision ============
    if not GPT_AVAILABLE:
        print("[OCR] ❌ GPT недоступен. Установите: pip install openai")
        return None
    
    try:
        print("[OCR] 🧠 Запускаю GPT-4o Vision (платно)...")
        plates, cost = await recognize_with_gpt_vision(image_path)
        
        if plates:
            # Конвертируем структурированные элементы в текст для парсера.
            text_lines = []
            for p in plates:
                candidate = (p.get("normalized_candidate") or p.get("raw_name") or "").strip()
                qty = int(p.get("qty", 1))
                if candidate:
                    text_lines.append(f"{candidate} {qty}")
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
    return """Ты OCR-система для таблиц железобетонных плит.

🎯 ТВОЯ ЗАДАЧА: Переписать таблицу БЕЗ ИНТЕРПРЕТАЦИИ

Работай как КСЕРОКС - копируй символы точно как видишь, не думай о смысле!

📋 Формат вывода - ТОЛЬКО JSON массив объектов:
[
  {
    "raw_name": "ПБ.19,6-12-10",
    "normalized_candidate": "ПБ 19,6-12-10",
    "qty": 7,
    "confidence": 0.92,
    "issues": []
  }
]

🔥 КРИТИЧЕСКИ ВАЖНО - ПОСИМВОЛЬНОЕ КОПИРОВАНИЕ:

Представь, что переписываешь таблицу от руки в блокнот.
Ты НЕ математик, ты НЕ думаешь о числах - ты просто КОПИРУЕШЬ текст!

✅ ПРАВИЛЬНО:
• Видишь "Плиты ПБ 66,2-12-8п", qty=6 ->
  {"raw_name":"ПБ 66,2-12-8п","normalized_candidate":"ПБ 66,2-12-8п","qty":6,"confidence":0.98,"issues":[]}
• Видишь "ПБ.19,6-12-10", qty=7 ->
  {"raw_name":"ПБ.19,6-12-10","normalized_candidate":"ПБ 19,6-12-10","qty":7,"confidence":0.92,"issues":["prefix_separator_dot"]}

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

2. **Формат raw_name:** убери слово "Плиты" и лишние пробелы, но не изменяй цифры.
3. **Формат normalized_candidate:** мягко нормализуй только префикс/пробелы (`ПБ.19,6` -> `ПБ 19,6`), не округляй числа.

4. **Формат qty:** число из правой колонки (обычно 1-99)
5. **Формат confidence:** число 0..1.
6. **Формат issues:** список строк; пустой список, если проблем нет.

7. **Порядок:** сверху вниз как в таблице

8. **Пропускай только заголовки:** "Наименование", "Кол-во", "Итого"

❌ КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО:
• Упрощать: 66,2 → 66
• Округлять: 21,5 → 22
• Убирать нули: 52,0 → 52
• Убирать десятичные: 10,2 → 10
• Группировать повторяющиеся строки
• Думать о смысле чисел (ты OCR, не математик!)

💡 Если сомневаешься - копируй ДОСЛОВНО! Лучше лишняя запятая, чем потерянная!

✅ Верни ТОЛЬКО JSON массив, без текста до и после!"""


def parse_gpt_response(response_text: str) -> List[Dict[str, Any]]:
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
        
        # Валидация: принимаем новый и legacy формат.
        validated_plates = []
        for plate in plates:
            if not isinstance(plate, dict):
                print(f"[GPT] ⚠️ Пропущена плита (не объект): {plate}")
                continue

            raw_name = plate.get("raw_name") or plate.get("name")
            normalized_candidate = plate.get("normalized_candidate") or raw_name
            if not raw_name or "qty" not in plate:
                print(f"[GPT] ⚠️ Пропущена плита (нет raw_name/name или qty): {plate}")
                continue

            try:
                qty = int(plate["qty"])
            except (ValueError, TypeError):
                print(f"[GPT] ⚠️ Пропущена плита (qty не число): {plate}")
                continue

            confidence = plate.get("confidence", 0.95)
            try:
                confidence = max(0.0, min(1.0, float(confidence)))
            except (ValueError, TypeError):
                confidence = 0.95

            issues = plate.get("issues") if isinstance(plate.get("issues"), list) else []
            validated_plates.append(
                {
                    "raw_name": str(raw_name).strip(),
                    "normalized_candidate": str(normalized_candidate).strip(),
                    "qty": qty,
                    "confidence": confidence,
                    "issues": issues,
                }
            )
        
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
        {'gpt_only': X, 'hybrid': X, 'photos': N}
    """
    
    # Средняя стоимость одного фото через GPT-4o
    avg_cost_per_photo = 0.002  # $0.002 = ~0.15₽
    
    return {
        'gpt_only': photos_per_month * avg_cost_per_photo,
        'hybrid': photos_per_month * avg_cost_per_photo,
        'photos': photos_per_month
    }


if __name__ == '__main__':
    # Пример расчёта стоимости
    print("💰 Оценка месячных затрат на OCR:")
    print("=" * 50)
    
    for count in [100, 500, 1000, 5000]:
        costs = estimate_monthly_cost(count)
        print(f"\n📊 {count} фото в месяц:")
        print(f"  • GPT-4o: ${costs['gpt_only']:.2f} (~{costs['gpt_only']*75:.0f}₽)")
        print(f"  • Гибрид (как GPT-4o): ${costs['hybrid']:.2f} (~{costs['hybrid']*75:.0f}₽)")

