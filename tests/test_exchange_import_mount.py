"""Production refuses to start when the 1C reply directory is not a mount."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.exchange_import_guard import (
    ExchangeImportMisconfigured,
    fail_fast_if_exchange_import_not_mounted,
)

_IMPORT = Path("/exchange/from_1c")


def test_development_without_mount_starts() -> None:
    fail_fast_if_exchange_import_not_mounted(
        "development",
        _IMPORT,
        environ={},
        mount_points=(),
    )


def test_production_without_env_refuses() -> None:
    with pytest.raises(ExchangeImportMisconfigured, match="EXCHANGE_IMPORT_DIR не задан"):
        fail_fast_if_exchange_import_not_mounted(
            "production",
            _IMPORT,
            environ={},
            mount_points=(_IMPORT,),
        )


def test_production_without_mount_refuses() -> None:
    with pytest.raises(
        ExchangeImportMisconfigured, match="каталог ответов не смонтирован отдельно"
    ):
        fail_fast_if_exchange_import_not_mounted(
            "production",
            _IMPORT,
            environ={"EXCHANGE_IMPORT_DIR": str(_IMPORT)},
            mount_points=(Path("/data"),),
        )


def test_production_with_mount_starts() -> None:
    fail_fast_if_exchange_import_not_mounted(
        "production",
        _IMPORT,
        environ={"EXCHANGE_IMPORT_DIR": str(_IMPORT)},
        mount_points=(Path("/data"), _IMPORT),
    )
