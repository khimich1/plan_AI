#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для сравнения EasyOCR и GPT-4o Vision

Использование:
    python test_ocr_comparison.py путь/к/фото.jpg

Что делает:
1. Проверяет доступность обоих методов
2. Распознаёт фото через EasyOCR
3. Распознаёт фото через GPT-4o
4. Сравнивает результаты
5. Показывает статистику
"""

import sys
import os
import asyncio
from pathlib import Path

# Добавляем корень проекта в путь
sys.path.insert(0, str(Path(__file__).parent))

from core.ocr_gpt import (
    recognize_text_smart,
    GPT_AVAILABLE,
    EASYOCR_AVAILABLE,
    estimate_monthly_cost
)


async def test_photo(image_path: str):
    """Тестирует распознавание фото обоими методами"""
    
    print("=" * 70)
    print("🧪 ТЕСТ СРАВНЕНИЯ OCR МЕТОДОВ")
    print("=" * 70)
    print()
    
    # Проверка существования файла
    if not os.path.exists(image_path):
        print(f"❌ Файл не найден: {image_path}")
        print("\nИспользование:")
        print("  python test_ocr_comparison.py путь/к/фото.jpg")
        return
    
    print(f"📸 Файл: {image_path}")
    file_size = os.path.getsize(image_path) / 1024
    print(f"📊 Размер: {file_size:.1f} КБ")
    print()
    
    # Проверка доступности методов
    print("🔍 Проверка доступных методов:")
    print(f"  🤖 EasyOCR: {'✅ Установлен' if EASYOCR_AVAILABLE else '❌ Не установлен (pip install easyocr)'}")
    print(f"  🧠 GPT-4o:  {'✅ Установлен' if GPT_AVAILABLE else '❌ Не установлен (pip install openai)'}")
    
    if os.getenv('OPENAI_API_KEY'):
        print(f"  🔑 API ключ: ✅ Найден (sk-...{os.getenv('OPENAI_API_KEY')[-6:]})")
    else:
        print(f"  🔑 API ключ: ⚠️ Не найден в переменных окружения")
    print()
    
    if not EASYOCR_AVAILABLE and not GPT_AVAILABLE:
        print("❌ Ни один метод недоступен!")
        return
    
    # ===== ТЕСТ 1: EasyOCR (если доступен) =====
    if EASYOCR_AVAILABLE:
        print("─" * 70)
        print("🤖 ТЕСТ 1: EasyOCR (бесплатный)")
        print("─" * 70)
        
        import time
        start_time = time.time()
        
        try:
            result_easyocr = await recognize_text_smart(
                image_path, 
                force_gpt=False,
                show_cost=False
            )
            
            elapsed = time.time() - start_time
            
            if result_easyocr and result_easyocr['method'] == 'EasyOCR':
                print(f"✅ Успешно! (за {elapsed:.1f} сек)")
                print(f"📊 Уверенность: {result_easyocr['confidence']*100:.0f}%")
                print(f"💰 Стоимость: $0 (бесплатно)")
                print()
                print("📝 Распознанный текст:")
                print("-" * 70)
                print(result_easyocr['text'])
                print("-" * 70)
            else:
                print(f"⚠️ EasyOCR не справился (за {elapsed:.1f} сек)")
                print("   Вероятно, не нашёл плиты в тексте")
        
        except Exception as e:
            print(f"❌ Ошибка: {e}")
        
        print()
    
    # ===== ТЕСТ 2: GPT-4o (если доступен) =====
    if GPT_AVAILABLE and os.getenv('OPENAI_API_KEY'):
        print("─" * 70)
        print("🧠 ТЕСТ 2: GPT-4o Vision (платный)")
        print("─" * 70)
        
        import time
        start_time = time.time()
        
        try:
            result_gpt = await recognize_text_smart(
                image_path,
                force_gpt=True,  # Принудительно GPT
                show_cost=False
            )
            
            elapsed = time.time() - start_time
            
            if result_gpt and result_gpt['method'] == 'GPT-4o':
                print(f"✅ Успешно! (за {elapsed:.1f} сек)")
                print(f"📊 Уверенность: {result_gpt['confidence']*100:.0f}%")
                print(f"💰 Стоимость: ${result_gpt['cost_usd']:.4f} (~{result_gpt['cost_usd']*75:.2f}₽)")
                print()
                print("📝 Распознанный текст:")
                print("-" * 70)
                print(result_gpt['text'])
                print("-" * 70)
                
                # Показываем структурированные данные (если есть)
                if result_gpt.get('plates'):
                    print()
                    print("📋 Структурированные данные:")
                    for plate in result_gpt['plates']:
                        print(f"  • {plate['name']} — {plate['qty']} шт")
            else:
                print(f"⚠️ GPT не вернул результат (за {elapsed:.1f} сек)")
        
        except Exception as e:
            print(f"❌ Ошибка: {e}")
        
        print()
    
    # ===== ТЕСТ 3: Гибридный режим (умный выбор) =====
    print("─" * 70)
    print("⚡ ТЕСТ 3: Гибридный режим (EasyOCR → GPT fallback)")
    print("─" * 70)
    
    import time
    start_time = time.time()
    
    try:
        result_hybrid = await recognize_text_smart(
            image_path,
            force_gpt=False,
            show_cost=True
        )
        
        elapsed = time.time() - start_time
        
        if result_hybrid:
            print(f"✅ Использован метод: {result_hybrid['method']}")
            print(f"⏱️ Время: {elapsed:.1f} сек")
            print(f"📊 Уверенность: {result_hybrid['confidence']*100:.0f}%")
            print(f"💰 Стоимость: ${result_hybrid['cost_usd']:.4f}")
            print()
            print("📝 Итоговый текст:")
            print("-" * 70)
            print(result_hybrid['text'])
            print("-" * 70)
        else:
            print("❌ Оба метода не смогли распознать текст")
    
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
    
    print()
    
    # ===== КАЛЬКУЛЯТОР СТОИМОСТИ =====
    print("=" * 70)
    print("💰 КАЛЬКУЛЯТОР СТОИМОСТИ")
    print("=" * 70)
    print()
    
    for count in [10, 100, 500, 1000]:
        costs = estimate_monthly_cost(count)
        print(f"📊 {count} фото в месяц:")
        print(f"   • EasyOCR:  $0 (всегда бесплатно)")
        print(f"   • GPT-4o:   ${costs['gpt_only']:.2f} (~{costs['gpt_only']*75:.0f}₽)")
        print(f"   • Гибрид:   ${costs['hybrid']:.2f} (~{costs['hybrid']*75:.0f}₽) ⭐ экономия 80%")
        print()
    
    print("=" * 70)
    print("✅ Тест завершён!")
    print("=" * 70)


def main():
    """Главная функция"""
    
    # Проверка аргументов
    if len(sys.argv) < 2:
        print("❌ Не указан путь к изображению!")
        print()
        print("Использование:")
        print(f"  python {sys.argv[0]} путь/к/фото.jpg")
        print()
        print("Примеры:")
        print(f"  python {sys.argv[0]} банк\\ знаний/1.jpeg")
        print(f"  python {sys.argv[0]} банк\\ знаний/photo_2025-07-10_10-06-26.jpg")
        return
    
    image_path = sys.argv[1]
    
    # Запускаем асинхронный тест
    asyncio.run(test_photo(image_path))


if __name__ == '__main__':
    main()

