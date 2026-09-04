# -*- coding: utf-8 -*-
"""Application settings (env-backed). Lives in ``core`` so inner packages do not import ``app``."""

from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Literal

from dotenv import load_dotenv
from pydantic import Field, computed_field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


@dataclass(frozen=True)
class VehicleClassLimits:
    max_weight_kg: float
    body_length_m: float = 13.2
    max_tiers: int = 4

_logger = logging.getLogger(__name__)

APP_SECRET_KEY_MIN_LENGTH = 32
_FORBIDDEN_APP_SECRET_KEYS = frozenset(
    {
        "",
        "change-this-secret-key-in-env",
        "changeme",
        "secret",
    }
)

BOT_ALLOWED_ROLES = frozenset({"admin", "manager", "production"})

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BOT_DIR = PROJECT_ROOT / "bot"

load_dotenv(PROJECT_ROOT / ".env", override=False)
load_dotenv(BOT_DIR / "bot.env", override=False)


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=(PROJECT_ROOT / ".env", BOT_DIR / "bot.env"),
        env_file_encoding="utf-8",
        extra="ignore",
        populate_by_name=True,
    )

    app_name: str = "Shishov Backend"
    app_env: str = "development"
    app_debug: bool = False
    app_secret_key: str = Field(alias="APP_SECRET_KEY")
    cookie_secure: bool | None = Field(default=None, alias="COOKIE_SECURE")
    cookie_samesite: Literal["lax", "strict", "none"] = Field(
        default="lax",
        alias="COOKIE_SAMESITE",
    )
    session_ttl_seconds: int = Field(
        default=60 * 60 * 12,
        alias="SESSION_COOKIE_MAX_AGE",
        ge=60,
        le=60 * 60 * 24 * 30,
    )
    # Строка из .env: pydantic-settings иначе пытается json.loads для list[str] до field_validator.
    cors_allowed_origins_raw: str = Field(
        default="http://localhost:5173",
        alias="BACKEND_CORS_ALLOWED_ORIGINS",
    )

    bot_token: str | None = Field(default=None, alias="BOT_TOKEN")
    bot_auth_enabled: bool = Field(default=True, alias="BOT_AUTH_ENABLED")
    bot_telegram_allowlist_raw: str = Field(default="", alias="BOT_TELEGRAM_ALLOWLIST")
    bot_auth_fail_closed: bool | None = Field(default=None, alias="BOT_AUTH_FAIL_CLOSED")
    openai_api_key: str | None = Field(default=None, alias="OPENAI_API_KEY")
    # External OCR / vision recognition (off by default; enable explicitly for staging).
    ocr_external_enabled: bool = Field(default=False, alias="OCR_EXTERNAL_ENABLED")
    ocr_recognition_mode: str = Field(default="full_gpt", alias="OCR_RECOGNITION_MODE")
    ocr_provider: Literal["gigachat", "openai"] = Field(default="openai", alias="OCR_PROVIDER")
    ocr_verify_enabled: bool = Field(default=True, alias="OCR_VERIFY_ENABLED")
    ocr_verify_mode: Literal["auto", "always", "never"] = Field(
        default="auto",
        alias="OCR_VERIFY_MODE",
    )
    ocr_max_api_calls: int = Field(default=2, alias="OCR_MAX_API_CALLS", ge=1, le=2)
    ocr_verify_auto_max_rows: int = Field(default=10, alias="OCR_VERIFY_AUTO_MAX_ROWS", ge=1)
    ocr_verify_auto_min_confidence: float = Field(
        default=0.92,
        alias="OCR_VERIFY_AUTO_MIN_CONFIDENCE",
        ge=0.0,
        le=1.0,
    )
    ocr_verify_auto_max_bytes: int = Field(
        default=819_200,
        alias="OCR_VERIFY_AUTO_MAX_BYTES",
        ge=1024,
    )
    ocr_verify_auto_min_short_side: int = Field(
        default=1000,
        alias="OCR_VERIFY_AUTO_MIN_SHORT_SIDE",
        ge=0,
    )
    gigachat_credentials: str | None = Field(default=None, alias="GIGACHAT_CREDENTIALS")
    gigachat_model: str = Field(default="GigaChat-2-Max", alias="GIGACHAT_MODEL")
    gigachat_scope: str = Field(default="GIGACHAT_API_PERS", alias="GIGACHAT_SCOPE")
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
    plans_dir: Path = Field(default=PROJECT_ROOT / "data" / "plans")
    plans_metadata_path: Path = Field(default=PROJECT_ROOT / "data" / "plans_metadata.json")
    current_plan_path: Path = Field(default=PROJECT_ROOT / "data" / "current_plan.json")
    work_calendar_path: Path = Field(default=PROJECT_ROOT / "data" / "work_calendar.json")
    archived_data_dir: Path = Field(default=PROJECT_ROOT / "bot_archived" / "data")
    logs_dir: Path = Field(default=PROJECT_ROOT / "logs")
    drafts_dir: Path = Field(default=PROJECT_ROOT / ".app_data" / "drafts")
    frontend_dist_dir: Path = Field(default=PROJECT_ROOT / "frontend" / "dist")

    # single_instance: локальные каталоги, горизонтальное масштабирование без sticky/session affinity
    # не поддерживается. shared_volume: оператор монтирует один и тот же том на все реплики для
    # drafts_dir и outputs_dir (NFS/EFS/Azure Files и т.п.).
    app_storage_layout: Literal["single_instance", "shared_volume"] = Field(
        default="single_instance",
        alias="APP_STORAGE_LAYOUT",
    )
    draft_store_lock_timeout_seconds: float = Field(
        default=60.0,
        alias="DRAFT_STORE_LOCK_TIMEOUT_SECONDS",
        ge=1.0,
        le=600.0,
    )

    commercial_upload_max_bytes: int = Field(
        default=50 * 1024 * 1024,
        alias="COMMERCIAL_UPLOAD_MAX_BYTES",
        ge=1024,
    )
    commercial_ocr_uploads_per_hour: int = Field(
        default=0,
        alias="COMMERCIAL_OCR_UPLOADS_PER_HOUR",
        ge=0,
    )
    auth_login_attempts_per_minute: int = Field(
        default=5,
        alias="AUTH_LOGIN_ATTEMPTS_PER_MINUTE",
        ge=1,
    )
    auth_login_rate_limit_window_seconds: int = Field(
        default=60,
        alias="AUTH_LOGIN_RATE_LIMIT_WINDOW_SECONDS",
        ge=1,
        le=3600,
    )
    auth_password_change_attempts: int = Field(
        default=3,
        alias="AUTH_PASSWORD_CHANGE_ATTEMPTS",
        ge=1,
    )
    auth_password_change_window_seconds: int = Field(
        default=900,
        alias="AUTH_PASSWORD_CHANGE_WINDOW_SECONDS",
        ge=1,
        le=86400,
    )
    # Comma-separated IPs of reverse proxies allowed to set X-Forwarded-For (e.g. 127.0.0.1).
    # Empty default: do not trust XFF; use the direct TCP client address only.
    trusted_proxy_ips_raw: str = Field(default="", alias="TRUSTED_PROXY_IPS")

    database_url: str | None = Field(default=None, alias="DATABASE_URL")
    redis_url: str | None = Field(default=None, alias="REDIS_URL")

    # Логистика (SHIP-000): лимиты классов ТС — JSON {"t20": {"max_weight_kg": 19800, ...}}.
    vehicle_class_limits_kg_raw: str = Field(
        default=(
            '{"t20": {"max_weight_kg": 19800, "body_length_m": 13.2, "max_tiers": 4}, '
            '"t30plus": {"max_weight_kg": 30000, "body_length_m": 13.2, "max_tiers": 4}}'
        ),
        alias="VEHICLE_CLASS_LIMITS_KG",
    )
    # Событие shipment_completed в папку обмена 1С — выключено до интеграции G.
    shipment_events_enabled: bool = Field(default=False, alias="SHIPMENT_EVENTS_ENABLED")
    exchange_export_dir: Path = Field(
        default=PROJECT_ROOT / "data" / "exchange_export",
        alias="EXCHANGE_EXPORT_DIR",
    )

    # Раскладка: жадное чередование целых и групп с резом по мин. армированию; сплиттер — согласованный выбор целых.
    layout_greedy_reinf_merge: bool = Field(default=True, alias="LAYOUT_GREEDY_REINF_MERGE")
    layout_track_reinf_preference: bool = Field(default=True, alias="LAYOUT_TRACK_REINF_PREFERENCE")
    # Порядок армирования в sequence: asc (слабые первыми) или desc (сильные первыми).
    layout_reinforcement_order: Literal["asc", "desc"] = Field(
        default="asc",
        alias="LAYOUT_REINFORCEMENT_ORDER",
    )
    # Если True — при выборе целой для начала дорожки разрешить «фазу 2» без отсечения по армированию
    # предыдущей дорожки (иначе — строгое правило может привести к TrackLayoutInvariantError).
    layout_track_start_reinf_relaxation: bool = Field(
        default=True,
        alias="LAYOUT_TRACK_START_REINF_RELAXATION",
    )
    # Дозаполнение хвоста дорожки переносом solid-плит с последующих дорожек (до 101 м).
    track_top_up_from_following: bool = Field(
        default=True,
        alias="TRACK_TOP_UP_FROM_FOLLOWING",
    )

    @field_validator("bot_telegram_allowlist_raw", mode="before")
    @classmethod
    def parse_bot_telegram_allowlist_raw(cls, value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def _parse_bot_telegram_allowlist(raw: str) -> dict[int, str]:
        normalized = raw.strip()
        if not normalized:
            return {}
        if normalized.startswith("["):
            try:
                parsed = json.loads(normalized)
            except json.JSONDecodeError as exc:
                raise ValueError(
                    "BOT_TELEGRAM_ALLOWLIST JSON must be an array of "
                    '{"id": <telegram_id>, "role": "<admin|manager|production>"} objects.'
                ) from exc
            if not isinstance(parsed, list):
                raise ValueError("BOT_TELEGRAM_ALLOWLIST JSON must be an array.")
            result: dict[int, str] = {}
            for entry in parsed:
                if not isinstance(entry, dict):
                    raise ValueError("Each BOT_TELEGRAM_ALLOWLIST JSON entry must be an object.")
                telegram_id = entry.get("id", entry.get("telegram_id"))
                role = entry.get("role")
                if telegram_id is None or role is None:
                    raise ValueError(
                        "BOT_TELEGRAM_ALLOWLIST JSON entries require 'id' (or 'telegram_id') and 'role'."
                    )
                role_str = str(role).strip().lower()
                if role_str not in BOT_ALLOWED_ROLES:
                    raise ValueError(
                        f"Invalid bot role '{role}'; allowed: admin, manager, production."
                    )
                result[int(telegram_id)] = role_str
            return result
        result = {}
        for part in normalized.split(","):
            item = part.strip()
            if not item:
                continue
            if ":" not in item:
                raise ValueError(
                    "BOT_TELEGRAM_ALLOWLIST entries must use telegram_id:role format "
                    "(example: 123456789:admin,987654321:manager)."
                )
            id_part, role_part = item.split(":", 1)
            role_str = role_part.strip().lower()
            if role_str not in BOT_ALLOWED_ROLES:
                raise ValueError(
                    f"Invalid bot role '{role_part.strip()}'; allowed: admin, manager, production."
                )
            result[int(id_part.strip())] = role_str
        return result

    @computed_field
    @property
    def bot_telegram_allowlist(self) -> dict[int, str]:
        return self._parse_bot_telegram_allowlist(self.bot_telegram_allowlist_raw)

    @computed_field
    @property
    def bot_auth_fail_closed_enabled(self) -> bool:
        if self.bot_auth_fail_closed is not None:
            return self.bot_auth_fail_closed
        return self.app_env.lower() == "production"

    @model_validator(mode="after")
    def validate_bot_telegram_allowlist_format(self) -> Settings:
        try:
            self._parse_bot_telegram_allowlist(self.bot_telegram_allowlist_raw)
        except ValueError as exc:
            raise ValueError(str(exc)) from exc
        return self

    @model_validator(mode="after")
    def validate_bot_telegram_auth(self) -> Settings:
        if not self.bot_auth_enabled:
            if self.app_env.lower() != "development":
                raise ValueError(
                    "BOT_AUTH_ENABLED can only be false when APP_ENV=development. "
                    "Configure BOT_TELEGRAM_ALLOWLIST with telegram_id:role entries."
                )
            return self
        if self.bot_auth_fail_closed_enabled and not self.bot_telegram_allowlist:
            raise ValueError(
                "BOT_TELEGRAM_ALLOWLIST must contain at least one telegram_id:role entry "
                "when bot authentication is enabled (example: 123456789:admin). "
                "Roles: admin, manager, production."
            )
        return self

    @model_validator(mode="after")
    def migrate_ocr_verify_enabled(self) -> Settings:
        if os.getenv("OCR_VERIFY_ENABLED") is not None:
            _logger.warning(
                "OCR_VERIFY_ENABLED is deprecated; use OCR_VERIFY_MODE (auto|always|never) instead."
            )
            if os.getenv("OCR_VERIFY_MODE") is None:
                mode = "always" if self.ocr_verify_enabled else "never"
                object.__setattr__(self, "ocr_verify_mode", mode)
        return self

    @model_validator(mode="after")
    def validate_app_secret_key(self) -> Settings:
        key = self.app_secret_key.strip()
        if key in _FORBIDDEN_APP_SECRET_KEYS or len(key) < APP_SECRET_KEY_MIN_LENGTH:
            raise ValueError(
                "APP_SECRET_KEY must be set via environment, must not use a known default, "
                f"and must be at least {APP_SECRET_KEY_MIN_LENGTH} characters "
                "(generate with: python -c \"import secrets; print(secrets.token_urlsafe(48))\")."
            )
        object.__setattr__(self, "app_secret_key", key)
        return self

    @model_validator(mode="after")
    def reject_debug_in_production(self) -> Settings:
        if self.app_env.lower() == "production" and self.app_debug:
            raise ValueError("APP_DEBUG must be false when APP_ENV=production")
        return self

    @computed_field
    @property
    def cookie_secure_enabled(self) -> bool:
        if self.cookie_secure is not None:
            return self.cookie_secure
        return self.app_env.lower() == "production"

    @field_validator("cors_allowed_origins_raw", mode="before")
    @classmethod
    def parse_cors_allowed_origins_raw(cls, value: object) -> str:
        if value is None:
            return "http://localhost:5173"
        return str(value).strip() or "http://localhost:5173"

    @staticmethod
    def _split_cors_origins(raw: str) -> list[str]:
        normalized = raw.strip()
        if not normalized:
            return []
        if normalized.startswith("["):
            try:
                parsed = json.loads(normalized)
            except json.JSONDecodeError:
                parsed = None
            if isinstance(parsed, list):
                return [str(item).strip() for item in parsed if str(item).strip()]
        return [item.strip() for item in normalized.split(",") if item.strip()]

    @computed_field
    @property
    def cors_allowed_origins(self) -> list[str]:
        return self._split_cors_origins(self.cors_allowed_origins_raw)

    @staticmethod
    def _split_proxy_ips(raw: str) -> frozenset[str]:
        normalized = raw.strip()
        if not normalized:
            return frozenset()
        return frozenset(item.strip() for item in normalized.split(",") if item.strip())

    @computed_field
    @property
    def trusted_proxy_ips(self) -> frozenset[str]:
        return self._split_proxy_ips(self.trusted_proxy_ips_raw)

    @staticmethod
    def _default_vehicle_class_limits() -> dict[str, VehicleClassLimits]:
        return {
            "t20": VehicleClassLimits(max_weight_kg=19_800.0, body_length_m=13.2, max_tiers=4),
            "t30plus": VehicleClassLimits(
                max_weight_kg=30_000.0, body_length_m=13.2, max_tiers=4
            ),
        }

    @staticmethod
    def _parse_vehicle_class_limits(raw: str) -> dict[str, VehicleClassLimits]:
        defaults = Settings._default_vehicle_class_limits()
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            return defaults
        if not isinstance(parsed, dict) or not parsed:
            return defaults

        result: dict[str, VehicleClassLimits] = {}
        for key, value in parsed.items():
            cls = str(key)
            if isinstance(value, (int, float)):
                base = defaults.get(cls, defaults["t20"])
                result[cls] = VehicleClassLimits(
                    max_weight_kg=float(value),
                    body_length_m=base.body_length_m,
                    max_tiers=base.max_tiers,
                )
                continue
            if isinstance(value, dict):
                base = defaults.get(cls, defaults["t20"])
                result[cls] = VehicleClassLimits(
                    max_weight_kg=float(value.get("max_weight_kg", base.max_weight_kg)),
                    body_length_m=float(value.get("body_length_m", base.body_length_m)),
                    max_tiers=int(value.get("max_tiers", base.max_tiers)),
                )
        return result or defaults

    @computed_field
    @property
    def vehicle_class_limits(self) -> dict[str, VehicleClassLimits]:
        return self._parse_vehicle_class_limits(self.vehicle_class_limits_kg_raw)

    @computed_field
    @property
    def vehicle_class_limits_kg(self) -> dict[str, int]:
        return {k: int(v.max_weight_kg) for k, v in self.vehicle_class_limits.items()}

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
    _logger.info(
        "App storage: layout=%s drafts_dir=%s outputs_dir=%s",
        settings.app_storage_layout,
        settings.drafts_dir,
        settings.outputs_dir,
    )
    if (
        settings.app_env.lower() == "production"
        and settings.app_storage_layout == "single_instance"
    ):
        _logger.warning(
            "APP_STORAGE_LAYOUT=single_instance: при нескольких репликах без sticky-сессий "
            "черновики и файлы могут быть недоступны с другого узла. Для горизонтального "
            "масштабирования смонтируйте общий том для DRAFTS_DIR и OUTPUTS_DIR и установите "
            "APP_STORAGE_LAYOUT=shared_volume, либо ограничьтесь одним воркером."
        )
    return settings
