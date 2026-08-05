"""MARCH-004: march pricing in core.commercial_pricing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    ensure_order_priced,
    lookup_march_price,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.march_price_db import import_march_prices_from_xlsx


def _write_sample_march_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Лестничные марши 1ЛМ 27-11-14-4", 13993.72, 14150.79, 14271.10, 14391.41, 14639.53],
        [6, "Лестничные марши ЛМ 2,8", 15819.88, 16000.91, 16139.57, 16278.24, 16564.20],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def march_db(tmp_path: Path) -> str:
    xlsx_path = tmp_path / "marches.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_march_xlsx(xlsx_path)
    import_march_prices_from_xlsx(str(xlsx_path), str(db_path))
    return str(db_path)


def test_lookup_march_price_found(march_db: str) -> None:
    price = lookup_march_price("1ЛМ 27-11-14-4", "B25", db_path=march_db)
    assert price == pytest.approx(14391.41, rel=1e-4)


def test_lookup_march_price_dot_form(march_db: str) -> None:
    price = lookup_march_price("ЛМ 2.8", "B25", db_path=march_db)
    assert price == pytest.approx(16278.24, rel=1e-4)


def test_lookup_march_price_missing_raises(march_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="1ЛМ 99"):
        lookup_march_price("1ЛМ 99-99-99-9", "B25", db_path=march_db)


def _march_order_item(**overrides: object) -> dict:
    base = {
        "product_kind": "march",
        "name": "1ЛМ 27-11-14-4",
        "mark": "1ЛМ 27-11-14-4",
        "concrete_grade": "B25",
        "qty": 2,
    }
    base.update(overrides)
    return base


def test_collect_unpriced_march_positions(march_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_march_order_item(mark="1ЛМ 99-99-99-9", name="1ЛМ 99-99-99-9")],
        db_path=march_db,
    )
    assert unpriced == ["1ЛМ 99-99-99-9 (B25)"]


def test_ensure_order_priced_march_success(march_db: str) -> None:
    ensure_order_priced([_march_order_item()], db_path=march_db)


def test_ensure_order_priced_march_raises(march_db: str) -> None:
    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced(
            [_march_order_item(mark="1ЛМ 99-99-99-9", name="1ЛМ 99-99-99-9")],
            db_path=march_db,
        )
    assert exc_info.value.positions == ["1ЛМ 99-99-99-9 (B25)"]
