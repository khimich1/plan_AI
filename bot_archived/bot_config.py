from pathlib import Path
from app.core.settings import get_settings

BOT_DIR = Path(__file__).parent
PROJECT_ROOT = BOT_DIR.parent
settings = get_settings()

BOT_TOKEN = settings.bot_token
BASE_DIR = PROJECT_ROOT
OUTPUTS_DIR = settings.outputs_dir
PRICES_DIR = settings.prices_dir
DB_PATH = settings.pb_db_path

OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

# Для обратной совместимости (если где-то используются строки)
OUTPUTS_DIR_STR = str(OUTPUTS_DIR)
PRICES_DIR_STR = str(PRICES_DIR)
DB_PATH_STR = str(DB_PATH)

