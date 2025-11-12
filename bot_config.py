import os
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv("bot.env")

# Токен бота (получите у @BotFather)
BOT_TOKEN = os.getenv("BOT_TOKEN")

# Пути к данным (используем существующие папки)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUTS_DIR = os.path.join(BASE_DIR, "Визуализация_Раскладки")
PRICES_DIR = os.path.join(BASE_DIR, "банк знаний")
DB_PATH = os.path.join(BASE_DIR, "pb.db")

# Создаём папку результатов если её нет
os.makedirs(OUTPUTS_DIR, exist_ok=True)
