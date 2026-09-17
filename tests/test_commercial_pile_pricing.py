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
        [40, "С110.30-9", 10000.0, 10100.0, 10200.0, 10300.0, 10400.0],
        [41, "С110.40-8", 11000.0, 11100.0, 11200.0, 11300.0, 11400.0],
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


def test_lookup_pile_price_latin_c_finds_cyrillic(pile_db: str) -> None:
    price = lookup_pile_price("C120.35-12", "B25", db_path=pile_db)
    assert price == pytest.approx(44634.03, rel=1e-4)


def test_lookup_pile_price_missing_raises(pile_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="С120.35-99"):
        lookup_pile_price("С120.35-99", "B25", db_path=pile_db)


def test_lookup_pile_price_strips_u_keeps_i(pile_db: str) -> None:
    assert lookup_pile_price("С120.35-12у", "B25", db_path=pile_db) == pytest.approx(
        44634.03, rel=1e-4
    )
    assert lookup_pile_price("С120.35-13и", "B25", db_path=pile_db) == pytest.approx(
        68385.98, rel=1e-4
    )
    with pytest.raises(PriceNotFoundError):
        lookup_pile_price("С120.35-13", "B25", db_path=pile_db)


def test_lookup_pile_price_fractional_load_fallback(pile_db: str) -> None:
    assert lookup_pile_price("C 110.30-9.1у", "B25", db_path=pile_db) == pytest.approx(
        10300.0, rel=1e-4
    )
    assert lookup_pile_price("C 110.40-8.1", "B25", db_path=pile_db) == pytest.approx(
        11300.0, rel=1e-4
    )
    with pytest.raises(PriceNotFoundError):
        lookup_pile_price("С110.30-6", "B25", db_path=pile_db)
    with pytest.raises(PriceNotFoundError):
        lookup_pile_price("С110.30-6у", "B25", db_path=pile_db)


def test_pile_preview_keeps_u_and_fraction_display(pile_db: str) -> None:
    from app.services.commercial_pile_service import CommercialPileService

    preview = CommercialPileService().generate_preview(
        "C 110.30-9.1у 2\nС120.35-13и 1",
        db_path=pile_db,
    )
    by_mark = {row["mark"]: row for row in preview.order_data}
    u_row = by_mark["С110.30-9.1у"]
    assert u_row["qty"] == 2
    assert u_row["unit_price"] == pytest.approx(10300.0, rel=1e-4)
    assert u_row["reinforced"] is True
    assert "9.1" in u_row["mark"]
    i_row = by_mark["С120.35-13и"]
    assert i_row["unit_price"] == pytest.approx(68385.98, rel=1e-4)
    assert i_row["reinforced"] is False


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
