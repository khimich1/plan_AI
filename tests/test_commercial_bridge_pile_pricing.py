"""BP-004: bridge pile pricing in core.commercial_pricing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    ensure_order_priced,
    lookup_bridge_pile_price,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.bridge_pile_price_db import import_bridge_pile_prices_from_xlsx


def _write_sample_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 25, 30],
        [1, "C8-35T1", 35695.27, 0],
        [2, "C8-35T4; C8-35В4", 49813.83, 0],
        [3, "C13-40T3", 0, 89879.61],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


@pytest.fixture()
def bridge_db(tmp_path: Path) -> str:
    xlsx_path = tmp_path / "bridge.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_xlsx(xlsx_path)
    import_bridge_pile_prices_from_xlsx(str(xlsx_path), str(db_path))
    return str(db_path)


def test_lookup_bridge_pile_price_found(bridge_db: str) -> None:
    price = lookup_bridge_pile_price("C8-35T1", "B25", db_path=bridge_db)
    assert price == pytest.approx(35695.27, rel=1e-4)


def test_lookup_bridge_pile_synonym(bridge_db: str) -> None:
    assert lookup_bridge_pile_price("C8-35В4", "B25", db_path=bridge_db) == pytest.approx(
        49813.83, rel=1e-4
    )
    assert lookup_bridge_pile_price("c8-35t4", "B25", db_path=bridge_db) == pytest.approx(
        49813.83, rel=1e-4
    )


def test_lookup_bridge_pile_missing_raises(bridge_db: str) -> None:
    with pytest.raises(PriceNotFoundError, match="C9-99"):
        lookup_bridge_pile_price("C9-99T1", "B25", db_path=bridge_db)


def _item(**overrides: object) -> dict:
    base = {
        "product_kind": "bridge_pile",
        "name": "C8-35T1",
        "mark": "C8-35T1",
        "concrete_grade": "B25",
        "qty": 2,
    }
    base.update(overrides)
    return base


def test_collect_unpriced_bridge_pile_positions(bridge_db: str) -> None:
    unpriced = collect_unpriced_positions(
        [_item(mark="C9-99T1", name="C9-99T1")],
        db_path=bridge_db,
    )
    assert unpriced == ["C9-99T1 (B25)"]


def test_ensure_order_priced_bridge_success(bridge_db: str) -> None:
    ensure_order_priced([_item()], db_path=bridge_db)


def test_ensure_order_priced_bridge_raises(bridge_db: str) -> None:
    with pytest.raises(UnpricedPlatesError) as exc_info:
        ensure_order_priced([_item(mark="C9-99T1", name="C9-99T1")], db_path=bridge_db)
    assert exc_info.value.positions == ["C9-99T1 (B25)"]


def test_manager_typed_mark_label(bridge_db: str) -> None:
    """Position label keeps manager spelling (В not forced to T)."""
    unpriced = collect_unpriced_positions(
        [_item(mark="C8-35В4", name="C8-35В4", concrete_grade="B30")],
        db_path=bridge_db,
    )
    assert unpriced == ["C8-35В4 (B30)"]
