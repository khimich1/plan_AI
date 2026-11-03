"""Refuse production startup when the 1C reply directory is not mounted."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path


class ExchangeImportMisconfigured(RuntimeError):
    """Production would ignore invoice numbers that 1C already wrote."""


def fail_fast_if_exchange_import_not_mounted(
    app_env: str,
    directory: Path,
    *,
    environ: Mapping[str, str],
    mount_points: Iterable[Path],
) -> None:
    """In production, the reply directory must be its own mount point.

    A missing directory makes the reader return without recording a number.
    Development does not require the variable.
    """
    if app_env.strip().lower() != "production":
        return
    if "EXCHANGE_IMPORT_DIR" not in environ:
        raise ExchangeImportMisconfigured("EXCHANGE_IMPORT_DIR не задан")
    mounted = {path.resolve() for path in mount_points}
    if directory.resolve() not in mounted:
        raise ExchangeImportMisconfigured("каталог ответов не смонтирован отдельно")
