"""Production refuses to start when the 1C exchange directory is not a mount."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.exchange_export_guard import (
    ExchangeExportMisconfigured,
    fail_fast_if_exchange_dir_not_mounted,
)

_EXPORT = Path("/data/exchange/to_1c")


def test_development_without_mount_starts() -> None:
    fail_fast_if_exchange_dir_not_mounted(
        "development",
        _EXPORT,
        environ={},
        mount_points=(),
    )


def test_production_without_env_refuses() -> None:
    with pytest.raises(ExchangeExportMisconfigured, match="EXCHANGE_EXPORT_DIR не задан"):
        fail_fast_if_exchange_dir_not_mounted(
            "production",
            _EXPORT,
            environ={},
            mount_points=(_EXPORT,),
        )


def test_production_without_mount_refuses() -> None:
    with pytest.raises(
        ExchangeExportMisconfigured, match="каталог обмена не смонтирован отдельно"
    ):
        fail_fast_if_exchange_dir_not_mounted(
            "production",
            _EXPORT,
            environ={"EXCHANGE_EXPORT_DIR": str(_EXPORT)},
            mount_points=(Path("/data"),),
        )


def test_production_with_mount_starts() -> None:
    fail_fast_if_exchange_dir_not_mounted(
        "production",
        _EXPORT,
        environ={"EXCHANGE_EXPORT_DIR": str(_EXPORT)},
        mount_points=(Path("/data"), _EXPORT),
    )
