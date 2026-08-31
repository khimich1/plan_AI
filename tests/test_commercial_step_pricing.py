"""STEP-004: step pricing in core.commercial_pricing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    ensure_order_priced,
    lookup_step_price,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.step_price_db import import_step_prices_from_xlsx


def _write_sample_step_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15],
        [1, "Лестничные ступени ЛС11", 1409.91],
        [2, "Лестничные ступени ЛС14-1лев", 1815.59],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def step_db(tmp_path: Path) -> str:
    xlsx_path = tmp_path / "steps.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_step_xlsx(xlsx_path)
    import_step_prices_from_xlsx(str(xlsx_path), str(db_path))
    return str(db_path)


def test_lookup_step_price_found(step_db: str) -> None:
    price = lookup_step_price("ЛС11", db_path=step_db)
    assert price == pytest.approx(1409.91, rel=1e-4)


def test_lookup_step_price_missing_raises(step_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="ЛС99"):
        lookup_step_price("ЛС99", db_path=step_db)


def _step_order_item(**overrides: object) -> dict:
    base = {
        "product_kind": "step",
        "name": "ЛС11",
        "mark": "ЛС11",
        "qty": 2,
    }
    base.update(overrides)
    return base


def test_collect_unpriced_step_positions(step_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_step_order_item(mark="ЛС99", name="ЛС99")],
        db_path=step_db,
    )
    assert unpriced == ["ЛС99"]


def test_ensure_order_priced_step_success(step_db: str) -> None:
    ensure_order_priced([_step_order_item()], db_path=step_db)


def test_ensure_order_priced_step_raises(step_db: str) -> None:
    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced(
            [_step_order_item(mark="ЛС99", name="ЛС99")],
            db_path=step_db,
        )
    assert exc_info.value.positions == ["ЛС99"]
