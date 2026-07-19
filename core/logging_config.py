from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging(
    level: int = logging.INFO,
    log_dir: Path | None = None,
    log_filename: str = "bot.log",
) -> None:
    """
    Единая настройка логов для всего проекта.

    Простыми словами:
    - Пишем логи в консоль и в файл logs/bot.log
    - Логи помогают понять, где и почему что-то сломалось
    """

    if log_dir is None:
        # корень проекта = папка на уровень выше core/
        log_dir = Path(__file__).resolve().parent.parent / "logs"

    log_dir.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()

    # Если логи уже настроены (например, при импорте в тестах) — не дублируем хендлеры.
    if root_logger.handlers:
        root_logger.setLevel(level)
        from core.config.logging import configure_optimization_logging_from_env

        configure_optimization_logging_from_env()
        return

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Windows console often uses cp1251/cp866, which fails on emoji.
    # Reconfigure stdout to utf-8 when possible to avoid logging errors.
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    file_handler = RotatingFileHandler(
        log_dir / log_filename,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    root_logger.setLevel(level)
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    from core.config.logging import configure_optimization_logging_from_env

    configure_optimization_logging_from_env()

