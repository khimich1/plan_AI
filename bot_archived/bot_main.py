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
from core.config.settings import get_settings

# Настройка логирования (консоль + logs/bot.log)
setup_logging(level=logging.INFO)
logger = logging.getLogger(__name__)

def init_database():
    """Проверяет наличие базы данных и инициализирует схему plita.db."""
    try:
        db_path = DB_PATH_STR
        # Путь к plita.db (КП и плиты) — именно этот файл открывать в DB Browser
        plita_db_path = str(PROJECT_ROOT / "plita.db")
        kp_db.ensure_schema(plita_db_path)
        logger.info(f"📂 База КП (plita.db): {plita_db_path}")
        
        if not os.path.exists(db_path):
            logger.warning(f"⚠️ База данных {db_path} не найдена!")
            logger.info("💡 Создайте базу pb.db или проверьте путь к файлу")
            logger.info("💡 Бот продолжит работу, но некоторые функции могут быть недоступны")
        else:
            logger.info(f"✅ База данных {db_path} найдена")
    except Exception as e:
        logger.error(f"❌ Ошибка проверки БД: {e}")
        # Не прерываем работу - база может быть создана позже

def validate_bot_startup() -> bool:
    """Fail-fast: settings (allowlist) must be valid before polling."""
    try:
        settings = get_settings()
    except Exception as exc:
        logger.error("❌ Конфигурация бота не прошла проверку: %s", exc)
        return False
    allowlist_count = len(settings.bot_telegram_allowlist)
    logger.info(
        "🔐 Bot auth: enabled=%s allowed_users=%s app_env=%s",
        settings.bot_auth_enabled,
        allowlist_count,
        settings.app_env,
    )
    if not settings.bot_auth_enabled:
        if settings.app_env.lower() != "development":
            logger.error(
                "❌ BOT_AUTH_ENABLED=false допустим только при APP_ENV=development"
            )
            return False
        logger.warning(
            "⚠️ BOT_AUTH_ENABLED=false — dev-only open access без synthetic admin "
            "(роль production, если пользователь не в allowlist)"
        )
    return True


_DEPRECATION_WARNING = (
    "DEPRECATED: Telegram-бот заморожен (2026-06-19) и не предназначен для production. "
    "Используйте веб-интерфейс (FastAPI + React). Код сохранён только для совместимости — см. bot/README.md"
)


async def main():
    """Основная функция запуска бота"""
    logger.warning(_DEPRECATION_WARNING)
    if not validate_bot_startup():
        sys.exit(1)
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

