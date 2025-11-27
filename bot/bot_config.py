import os
from pathlib import Path
from dotenv import load_dotenv

# Загружаем переменные окружения из .env в корне проекта
BOT_DIR = Path(__file__).parent
PROJECT_ROOT = BOT_DIR.parent
load_dotenv(PROJECT_ROOT / ".env")

# Токен бота (получите у @BotFather)
BOT_TOKEN = os.getenv("BOT_TOKEN")

# Пути к данным (BASE_DIR указывает на корень проекта, на уровень выше bot/)
BASE_DIR = PROJECT_ROOT
OUTPUTS_DIR = BASE_DIR / "Визуализация_Раскладки"
PRICES_DIR = BASE_DIR / "банк знаний"
DB_PATH = BASE_DIR / "pb.db"

# Создаём папку результатов если её нет
OUTPUTS_DIR.mkdir(exist_ok=True)

# Для обратной совместимости (если где-то используются строки)
OUTPUTS_DIR_STR = str(OUTPUTS_DIR)
PRICES_DIR_STR = str(PRICES_DIR)
DB_PATH_STR = str(DB_PATH)

