"""FBS-004: FBS pricing in core.commercial_pricing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    ensure_order_priced,
    lookup_fbs_price,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.fbs_price_db import import_fbs_prices_from_xlsx


def _write_sample_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
        [2, "ФБС 12.4.6-Т", 2683.65, 2848.31, 2899.91, 2951.52],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def fbs_db(tmp_path: Path) -> str:
    xlsx_path = tmp_path / "fbs.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_xlsx(xlsx_path)
    import_fbs_prices_from_xlsx(str(xlsx_path), str(db_path))
    return str(db_path)


def test_lookup_fbs_price_found(fbs_db: str) -> None:
    price = lookup_fbs_price("ФБС 9.3.6-Т", "B25", db_path=fbs_db)
    assert price == pytest.approx(1788.33, rel=1e-4)


def test_lookup_fbs_normalized(fbs_db: str) -> None:
    assert lookup_fbs_price("фбс 9.3.6-т", "B7_5", db_path=fbs_db) == pytest.approx(
        1640.75, rel=1e-4
    )


def test_lookup_fbs_missing_raises(fbs_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="ФБС 99"):
        lookup_fbs_price("ФБС 99.9.9-Т", "B25", db_path=fbs_db)


def _item(**overrides: object) -> dict:
    base = {
        "product_kind": "fbs",
        "name": "ФБС 9.3.6-Т",
        "mark": "ФБС 9.3.6-Т",
        "concrete_grade": "B25",
        "qty": 2,
    }
    base.update(overrides)
    return base


def test_collect_unpriced_fbs_positions(fbs_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_item(mark="ФБС 99.9.9-Т", name="ФБС 99.9.9-Т")],
        db_path=fbs_db,
    )
    assert unpriced == ["ФБС 99.9.9-Т (B25)"]


def test_ensure_order_priced_fbs_success(fbs_db: str) -> None:
    ensure_order_priced([_item()], db_path=fbs_db)


def test_ensure_order_priced_fbs_raises(fbs_db: str) -> None:
    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced(
            [_item(mark="ФБС 99.9.9-Т", name="ФБС 99.9.9-Т")],
            db_path=fbs_db,
        )
    assert exc_info.value.positions == ["ФБС 99.9.9-Т (B25)"]


def test_manager_typed_mark_label(fbs_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_item(mark="фбс 9.3.6-т", name="фбс 9.3.6-т", concrete_grade="B30")],
        db_path=fbs_db,
    )
    assert unpriced == ["фбс 9.3.6-т (B30)"]
