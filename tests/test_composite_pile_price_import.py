"""CP-101/102: composite pile price import + lookup."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.composite_pile_price_db import (
    CompositePilePriceImportError,
    get_composite_pile_price,
    import_composite_pile_prices_from_xlsx,
    init_composite_pile_prices_schema,
    parse_composite_pile_price_rows_from_xlsx,
)

FIXTURE = Path(__file__).parent / "fixtures" / "composite_pile_price_sample.xlsx"
BAD = Path(__file__).parent / "fixtures" / "composite_pile_price_bad_name.xlsx"


def test_parse_canon_and_five_grades() -> None:
    rows = parse_composite_pile_price_rows_from_xlsx(str(FIXTURE))
    by = {(m, g): p for m, g, p in rows}
    assert by[("С60.30-ВС.1", "B15")] == 100.0
    assert by[("С60.30-ВС.1", "B20")] == 110.0
    assert by[("С60.30-ВС.1", "B22_5")] == 120.0
    assert by[("С60.30-ВС.1", "B25")] == 130.0
    assert by[("С60.30-ВС.1", "B30_granite")] == 140.0
    # empty / zero cells skipped
    assert ("С80.30-НС.1", "B22_5") not in by
    assert ("С80.30-НС.1", "B30_granite") not in by
    assert by[("С80.30-НС.1", "B25")] == 230.0


def test_bad_name_raises() -> None:
    with pytest.raises(CompositePilePriceImportError, match="кривое имя"):
        parse_composite_pile_price_rows_from_xlsx(str(BAD))


def test_import_and_lookup(tmp_path: Path) -> None:
    db = str(tmp_path / "pb.db")
    init_composite_pile_prices_schema(db)
    n = import_composite_pile_prices_from_xlsx(str(FIXTURE), db)
    assert n > 0
    assert get_composite_pile_price("С60.30-ВС.1", "B25", db) == 130.0
    assert get_composite_pile_price("Сваи С 60.30-ВС.1", "B25", db) == 130.0
    assert get_composite_pile_price("С60.30-ВС.1", "B15", db) == 100.0
    assert get_composite_pile_price("С999.30-ВС.1", "B25", db) is None
