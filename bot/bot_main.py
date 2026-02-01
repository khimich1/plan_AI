import asyncio
import logging
import os
import sys
from pathlib import Path

# Добавляем корень проекта в sys.path для импорта модулей из корня
BOT_DIR = Path(__file__).parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.logging_config import setup_logging
from core import kp_db

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode

from bot.handlers import register_all_handlers
from bot.bot_config import BOT_TOKEN, DB_PATH_STR

# Настройка логирования (консоль + logs/bot.log)
setup_logging(level=logging.INFO)
logger = logging.getLogger(__name__)

def init_database():
    """Проверяет наличие базы данных"""
    try:
        db_path = DB_PATH_STR
        
        if not os.path.exists(db_path):
            logger.warning(f"⚠️ База данных {db_path} не найдена!")
            logger.info("💡 Создайте базу pb.db или проверьте путь к файлу")
            logger.info("💡 Бот продолжит работу, но некоторые функции могут быть недоступны")
        else:
            logger.info(f"✅ База данных {db_path} найдена")
            
            # Автоматически восстанавливаем застрявшие плиты при старте
            try:
                recovered = kp_db.recover_stuck_plates(db_path)
                if recovered > 0:
                    logger.info(f"🔧 При старте восстановлено {recovered} застрявших плит")
            except Exception as e:
                logger.warning(f"⚠️ Не удалось проверить застрявшие плиты: {e}")
    except Exception as e:
        logger.error(f"❌ Ошибка проверки БД: {e}")
        # Не прерываем работу - база может быть создана позже

async def main():
    """Основная функция запуска бота"""
    if not BOT_TOKEN:
        logger.error("❌ BOT_TOKEN не найден! Проверьте файл bot/bot.env")
        logger.error("💡 Получите токен у @BotFather в Telegram")
        return
    
    # Проверяем, что токен не является заглушкой
    if BOT_TOKEN == "your_bot_token_here" or len(BOT_TOKEN) < 20:
        logger.error("❌ Токен не установлен или неверный!")
        logger.error("💡 Откройте файл bot/bot.env и замените 'your_bot_token_here' на настоящий токен")
        logger.error("💡 Получите токен у @BotFather в Telegram: /newbot")
        return
    
    # Инициализируем БД
    init_database()
    
    # Создаём бота и диспетчер
    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML)
    )
    dp = Dispatcher()
    
    # Регистрируем обработчики
    register_all_handlers(dp)
    
    logger.info("🚀 Бот запущен!")
    
    try:
        # Запускаем поллинг
        await dp.start_polling(bot)
    except Exception as e:
        logger.error(f"Ошибка при запуске бота: {e}")
    finally:
        await bot.session.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Бот остановлен пользователем")
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")

