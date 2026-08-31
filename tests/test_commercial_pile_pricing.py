"""PILE-003: pile pricing in core.commercial_pricing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    ensure_order_priced,
    lookup_pile_price,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.pile_price_db import import_pile_prices_from_xlsx


def _write_sample_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [None, "35 СЕЧЕНИЕ", None, None, None, None, None],
        [69, "С120.35-12", 43760.31, 44108.15, 44371.09, 44634.03, 46159.37],
        [91, "С120.35-13и", 67512.27, 67860.11, 68123.05, 68385.98, 69911.33],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def pile_db(tmp_path: Path) -> str:
    xlsx_path = tmp_path / "piles.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_pile_xlsx(xlsx_path)
    import_pile_prices_from_xlsx(str(xlsx_path), str(db_path))
    return str(db_path)


def test_lookup_pile_price_found(pile_db: str) -> None:
    price = lookup_pile_price("С120.35-12", "B25", db_path=pile_db)
    assert price == pytest.approx(44634.03, rel=1e-4)


def test_lookup_pile_price_missing_raises(pile_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="С120.35-99"):
        lookup_pile_price("С120.35-99", "B25", db_path=pile_db)


def _pile_order_item(**overrides: object) -> dict:
    base = {
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 2,
    }
    base.update(overrides)
    return base


def test_collect_unpriced_pile_positions(pile_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_pile_order_item(mark="С120.35-99")],
        db_path=pile_db,
    )
    assert unpriced == ["С120.35-99 (B25)"]


def test_ensure_order_priced_pile_success(pile_db: str) -> None:
    ensure_order_priced([_pile_order_item()], db_path=pile_db)


def test_ensure_order_priced_pile_raises(pile_db: str) -> None:
    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced(
            [_pile_order_item(mark="С120.35-99", name="С120.35-99")],
            db_path=pile_db,
        )
    assert exc_info.value.positions == ["С120.35-99 (B25)"]
