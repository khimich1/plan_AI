"""CP-201+: schema + pricing for composite piles."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.commercial_pricing import (
    collect_unpriced_positions,
    lookup_composite_pile_price,
)
from core.composite_pile_kit import import_composite_pile_kit_from_xlsx
from core.composite_pile_price_db import import_composite_pile_prices_from_xlsx
from core.exceptions import PriceNotFoundError
from core.kp_db_schema import ensure_schema
from core.price_db import _connect

KIT = Path(__file__).parent / "fixtures" / "composite_pile_kit_sample.xlsx"
PRICE = Path(__file__).parent / "fixtures" / "composite_pile_price_sample.xlsx"


def test_kp_composite_piles_schema(tmp_path: Path) -> None:
    db = str(tmp_path / "kp.db")
    ensure_schema(db)
    conn = _connect(db)
    try:
        cur = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='kp_composite_piles'"
        )
        assert cur.fetchone() is not None
        # idempotent
        ensure_schema(db)
    finally:
        conn.close()


def test_lookup_composite_pile_price_strict(tmp_path: Path) -> None:
    db = str(tmp_path / "pb.db")
    import_composite_pile_kit_from_xlsx(str(KIT), db)
    import_composite_pile_prices_from_xlsx(str(PRICE), db)
    assert lookup_composite_pile_price("С60.30-ВС.1", "B25", db_path=db) == 130.0
    assert lookup_composite_pile_price("С60.30-ВС.1", "B15", db_path=db) == 100.0
    with pytest.raises(PriceNotFoundError):
        lookup_composite_pile_price("С999.30-ВС.1", "B25", db_path=db)


def test_unpriced_section_blocks_collect(tmp_path: Path) -> None:
    db = str(tmp_path / "pb.db")
    import_composite_pile_prices_from_xlsx(str(PRICE), db)
    order = [
        {
            "product_kind": "composite_pile",
            "mark": "С60.30-ВС.1",
            "concrete_grade": "B25",
            "qty": 1,
            "unit_price": None,
        }
    ]
    # B25 is priced — collect should clear after lookup succeeds in ensure path
    # With unit_price None, collect tries lookup — should find price → empty unpriced
    assert collect_unpriced_positions(order, db_path=db) == []

    missing = [
        {
            "product_kind": "composite_pile",
            "mark": "С999.30-ВС.1",
            "concrete_grade": "B25",
            "qty": 1,
            "unit_price": None,
        }
    ]
    assert collect_unpriced_positions(missing, db_path=db)
