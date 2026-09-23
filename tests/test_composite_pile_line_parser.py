"""CP-005/006: composite pile line parser / expand + merge."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.composite_pile_kit import import_composite_pile_kit_from_xlsx
from core.composite_pile_line_parser import (
    expand_composite_pile_order_text,
    expand_order_line,
    merge_composite_pile_sections,
)

FIXTURE = Path(__file__).parent / "fixtures" / "composite_pile_kit_sample.xlsx"


@pytest.fixture()
def kit_db(tmp_path: Path) -> str:
    db = str(tmp_path / "pb.db")
    import_composite_pile_kit_from_xlsx(str(FIXTURE), db)
    return db


def test_three_joints_expand_to_different_pairs(kit_db: str) -> None:
    cup = expand_order_line("С140.30-С 1", db_path=kit_db)
    weld = expand_order_line("С140.30-Св 1", db_path=kit_db)
    vp = expand_order_line("С140.30-Св.ВП 1", db_path=kit_db)
    assert cup.parsed and weld.parsed and vp.parsed
    assert [(s.mark, s.qty) for s in cup.sections] == [
        ("С60.30-ВС.1", 1),
        ("С80.30-НС.1", 1),
    ]
    assert [(s.mark, s.qty) for s in weld.sections] == [
        ("С60.30-ВСв.1", 1),
        ("С80.30-НСв.1", 1),
    ]
    assert [(s.mark, s.qty) for s in vp.sections] == [
        ("С60.30-ВСв.6", 1),
        ("С80.30-НСв.6", 1),
    ]


def test_whole_pile_qty_copied_to_both_sections(kit_db: str) -> None:
    result = expand_order_line("С140.30-С 5", db_path=kit_db)
    assert result.parsed
    assert [(s.mark, s.concrete_grade, s.qty) for s in result.sections] == [
        ("С60.30-ВС.1", "B25", 5),
        ("С80.30-НС.1", "B25", 5),
    ]


def test_merge_whole_pile_plus_section_alias(kit_db: str) -> None:
    sections = expand_composite_pile_order_text(
        "С140.30-С 2\nС 60.30-ВС.1 3",
        db_path=kit_db,
    )
    by_key = {(s.mark, s.concrete_grade): s.qty for s in sections if s.parsed}
    assert by_key[("С60.30-ВС.1", "B25")] == 5
    assert by_key[("С80.30-НС.1", "B25")] == 2


def test_different_grades_stay_separate(kit_db: str) -> None:
    sections = expand_composite_pile_order_text(
        "С140.30-С B20 2\nС60.30-ВС.1 3",
        db_path=kit_db,
    )
    keys = {(s.mark, s.concrete_grade, s.qty) for s in sections if s.parsed}
    assert ("С60.30-ВС.1", "B20", 2) in keys
    assert ("С80.30-НС.1", "B20", 2) in keys
    assert ("С60.30-ВС.1", "B25", 3) in keys


def test_unknown_whole_pile_not_parsed(kit_db: str) -> None:
    result = expand_order_line("С999.30-С 1", db_path=kit_db)
    assert result.parsed is False
    assert result.reason_code == "pattern_not_matched"
    assert result.sections == []


def test_section_alias_single_canon_row(kit_db: str) -> None:
    result = expand_order_line("Сваи С 60.30-ВС.1 4", db_path=kit_db)
    assert result.parsed
    assert len(result.sections) == 1
    assert result.sections[0].mark == "С60.30-ВС.1"
    assert result.sections[0].qty == 4
    assert result.sections[0].concrete_grade == "B25"


def test_grade_latin_b_not_cyrillic(kit_db: str) -> None:
    result = expand_order_line("С60.30-ВС.1 B25 2", db_path=kit_db)
    assert result.parsed
    assert result.sections[0].concrete_grade == "B25"
    assert "В25" not in (result.sections[0].concrete_grade or "")


def test_merge_helper_sums_same_canon_grade() -> None:
    from core.composite_pile_line_parser import CompositePileSection

    merged = merge_composite_pile_sections(
        [
            CompositePileSection(mark="С60.30-ВС.1", concrete_grade="B25", qty=2),
            CompositePileSection(mark="С60.30-ВС.1", concrete_grade="B25", qty=3),
            CompositePileSection(mark="С60.30-ВС.1", concrete_grade="B20", qty=1),
        ]
    )
    assert len(merged) == 2
    by = {(s.mark, s.concrete_grade): s.qty for s in merged}
    assert by[("С60.30-ВС.1", "B25")] == 5
    assert by[("С60.30-ВС.1", "B20")] == 1
