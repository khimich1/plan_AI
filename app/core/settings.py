from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from dotenv import load_dotenv
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BOT_DIR = PROJECT_ROOT / "bot"

load_dotenv(PROJECT_ROOT / ".env", override=False)
load_dotenv(BOT_DIR / "bot.env", override=False)


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=(PROJECT_ROOT / ".env", BOT_DIR / "bot.env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_name: str = "Shishov Backend"
    app_env: str = "development"
    app_debug: bool = False
    app_secret_key: str = Field(
        default="change-this-secret-key-in-env",
        alias="APP_SECRET_KEY",
    )

    bot_token: str | None = Field(default=None, alias="BOT_TOKEN")
    openai_api_key: str | None = Field(default=None, alias="OPENAI_API_KEY")
    ocr_recognition_mode: str = Field(default="full_gpt", alias="OCR_RECOGNITION_MODE")
    weight_source: str = Field(default="formula", alias="WEIGHT_SOURCE")

    pb_db_path: Path = Field(default=PROJECT_ROOT / "pb.db")
    plita_db_path: Path = Field(default=PROJECT_ROOT / "plita.db")
    price_xlsx_path: Path = Field(
        default=PROJECT_ROOT / "банк знаний" / "Новые цены для прайса с 19.08.24.xlsx"
    )
    cuts_docx_path: Path = Field(
        default=PROJECT_ROOT / "банк знаний" / "Письмо Цены с 29.05.2024 цены на резы.docx"
    )
    outputs_dir: Path = Field(default=PROJECT_ROOT / "Визуализация_Раскладки")
    prices_dir: Path = Field(default=PROJECT_ROOT / "банк знаний")
    plans_dir: Path = Field(default=PROJECT_ROOT / "bot" / "data" / "plans")
    plans_metadata_path: Path = Field(default=PROJECT_ROOT / "bot" / "data" / "plans_metadata.json")
    work_calendar_path: Path = Field(default=PROJECT_ROOT / "bot" / "data" / "work_calendar.json")
    logs_dir: Path = Field(default=PROJECT_ROOT / "logs")
    drafts_dir: Path = Field(default=PROJECT_ROOT / ".app_data" / "drafts")

    database_url: str | None = Field(default=None, alias="DATABASE_URL")
    redis_url: str | None = Field(default=None, alias="REDIS_URL")

    bootstrap_admin_username: str | None = Field(default=None, alias="APP_ADMIN_USERNAME")
    bootstrap_admin_password: str | None = Field(default=None, alias="APP_ADMIN_PASSWORD")
    bootstrap_admin_role: str = Field(default="admin", alias="APP_ADMIN_ROLE")

    def ensure_directories(self) -> None:
        self.outputs_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.plans_dir.mkdir(parents=True, exist_ok=True)
        self.work_calendar_path.parent.mkdir(parents=True, exist_ok=True)
        self.drafts_dir.mkdir(parents=True, exist_ok=True)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    settings = Settings()
    settings.ensure_directories()
    return settings

