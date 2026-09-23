"""CP-003/004: composite pile kit (assemblies) import + lookup."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.composite_pile_kit import (
    get_composite_pile_assembly,
    import_composite_pile_kit_from_xlsx,
    init_composite_pile_assemblies_schema,
    parse_composite_pile_kit_rows_from_xlsx,
)

FIXTURE = Path(__file__).parent / "fixtures" / "composite_pile_kit_sample.xlsx"
FULL_KIT = (
    Path(__file__).resolve().parents[1]
    / "банк знаний"
    / "Составные сваи — комплектация секций (серия 1.011.1-10 вып.8).xlsx"
)


def _length_dm_from_mark(mark: str) -> int:
    """Extract pile/section length in dm from canon mark (digits before first '.')."""
    body = mark.split("-", 1)[0]
    # С60.30 → 60; С140.30 → 140
    digits = "".join(ch if ch.isdigit() else " " for ch in body)
    parts = digits.split()
    assert parts, f"no length in mark {mark!r}"
    return int(parts[0])


def test_parse_fixture_has_three_distinct_joints() -> None:
    rows = parse_composite_pile_kit_rows_from_xlsx(str(FIXTURE))
    marks = {r.pile_mark for r in rows}
    assert "С140.30-С" in marks
    assert "С140.30-Св" in marks
    assert "С140.30-Св.ВП" in marks
    by_mark = {r.pile_mark: r for r in rows}
    assert by_mark["С140.30-С"].upper_mark == "С60.30-ВС.1"
    assert by_mark["С140.30-С"].lower_mark == "С80.30-НС.1"
    assert by_mark["С140.30-Св"].upper_mark == "С60.30-ВСв.1"
    assert by_mark["С140.30-Св.ВП"].upper_mark == "С60.30-ВСв.6"


def test_import_fixture_and_lookup(tmp_path: Path) -> None:
    db = str(tmp_path / "pb.db")
    init_composite_pile_assemblies_schema(db)
    n = import_composite_pile_kit_from_xlsx(str(FIXTURE), db)
    assert n == 4
    asm = get_composite_pile_assembly("С140.30-С", db)
    assert asm is not None
    assert asm == ("С60.30-ВС.1", "С80.30-НС.1")
    assert get_composite_pile_assembly("С999.30-С", db) is None


def test_fixture_length_invariant() -> None:
    rows = parse_composite_pile_kit_rows_from_xlsx(str(FIXTURE))
    for row in rows:
        pile_len = _length_dm_from_mark(row.pile_mark)
        upper_len = _length_dm_from_mark(row.upper_mark)
        lower_len = _length_dm_from_mark(row.lower_mark)
        assert upper_len + lower_len == pile_len, row.pile_mark


@pytest.mark.skipif(not FULL_KIT.is_file(), reason="full kit xlsx not in банк знаний")
def test_full_kit_123_assemblies_and_length_sum() -> None:
    rows = parse_composite_pile_kit_rows_from_xlsx(str(FULL_KIT))
    assert len(rows) == 123
    marks = [r.pile_mark for r in rows]
    assert len(set(marks)) == 123
    for row in rows:
        pile_len = _length_dm_from_mark(row.pile_mark)
        upper_len = _length_dm_from_mark(row.upper_mark)
        lower_len = _length_dm_from_mark(row.lower_mark)
        assert upper_len + lower_len == pile_len, (
            f"{row.pile_mark}: {upper_len}+{lower_len}!={pile_len}"
        )
