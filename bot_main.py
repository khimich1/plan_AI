import asyncio
import logging
import os
from pathlib import Path
from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode

from bot_handlers import register_handlers
from bot_config import BOT_TOKEN

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def init_database():
    """Проверяет наличие базы данных"""
    try:
        db_path = "pb.db"
        
        if not os.path.exists(db_path):
            logger.warning(f"⚠️ База данных {db_path} не найдена!")
            logger.info("💡 Создайте базу pb.db или проверьте путь к файлу")
            logger.info("💡 Бот продолжит работу, но некоторые функции могут быть недоступны")
        else:
            logger.info(f"✅ База данных {db_path} найдена")
    except Exception as e:
        logger.error(f"❌ Ошибка проверки БД: {e}")
        # Не прерываем работу - база может быть создана позже

async def main():
    """Основная функция запуска бота"""
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN не найден! Проверьте файл .env")
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
    register_handlers(dp)
    
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
