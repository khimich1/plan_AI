#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт запуска Telegram бота (DEPRECATED — см. bot/README.md).

Запускайте из корня проекта: python run_bot.py
Не использовать в production; активная поддержка прекращена 2026-06-19.
"""

import asyncio
import sys

from bot.bot_main import main

if __name__ == "__main__":
    print(
        "WARNING: Telegram-бот DEPRECATED — не для production. Подробнее: bot/README.md",
        file=sys.stderr,
    )
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nБот остановлен пользователем")
    except Exception as e:
        error_msg = str(e)
        print(f"Критическая ошибка: {error_msg}")
        
        # Более понятные сообщения для частых ошибок
        if "Token is invalid" in error_msg or "Unauthorized" in error_msg:
            print("\n" + "="*50)
            print("❌ ПРОБЛЕМА: Токен бота неверный или не установлен!")
            print("="*50)
            print("💡 Что делать:")
            print("   1. Откройте Telegram и найдите @BotFather")
            print("   2. Отправьте команду /newbot")
            print("   3. Следуйте инструкциям для создания бота")
            print("   4. Скопируйте полученный токен")
            print("   5. Откройте файл: bot/bot.env")
            print("   6. Замените 'your_bot_token_here' на ваш токен")
            print("="*50)

