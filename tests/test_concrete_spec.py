"""Справочник F/W заводской таблицы и прикрепление снимка к строке КП."""

from __future__ import annotations

import pytest

from core.concrete_spec import ConcreteSpec, apply_line_specs, suggest_concrete_spec


def test_b25_suggests_granite_f200_w8() -> None:
    spec = suggest_concrete_spec("B25")
    assert spec == ConcreteSpec(
        frost_resistance="F200",
        waterproofness="W8",
        concrete_aggregate="granite",
        ordinary_waterproofness="W6",
    )


def test_b22_5_granite_is_w6_and_ordinary_column_is_w4() -> None:
    dotted = suggest_concrete_spec("B22.5")
    underscored = suggest_concrete_spec("B22_5")
    expected = ConcreteSpec(
        frost_resistance="F200",
        waterproofness="W6",
        concrete_aggregate="granite",
        ordinary_waterproofness="W4",
    )
    assert dotted == expected
    assert underscored == expected


@pytest.mark.parametrize("code", ["B30", "B30_granite"])
def test_b30_is_granite_only_f300_w10(code: str) -> None:
    spec = suggest_concrete_spec(code)
    assert spec == ConcreteSpec(
        frost_resistance="F300",
        waterproofness="W10",
        concrete_aggregate="granite",
        ordinary_waterproofness=None,
    )


@pytest.mark.parametrize("code", ["М500", "M500"])
def test_m500_suggests_f300_w12(code: str) -> None:
    spec = suggest_concrete_spec(code)
    assert spec == ConcreteSpec(
        frost_resistance="F300",
        waterproofness="W12",
        concrete_aggregate="granite",
        ordinary_waterproofness=None,
    )


@pytest.mark.parametrize("code", ["", "   ", "B99", "не бетон"])
def test_unknown_or_blank_grade_returns_none(code: str) -> None:
    assert suggest_concrete_spec(code) is None


@pytest.mark.parametrize(
    ("code", "frost", "granite_w", "ordinary_w"),
    [
        ("B7.5", "F50", "W2", "W2"),
        ("B7_5", "F50", "W2", "W2"),
        ("B12.5", "F50", "W2", "W2"),
        ("B15", "F100", "W4", "W4"),
        ("B20", "F100", "W4", "W4"),
        ("М400", "F300", "W10", None),
        ("M400", "F300", "W10", None),
        ("B35", "F300", "W10", None),
        ("B40", "F300", "W12", None),
        ("B45", "F300", "W12", None),
    ],
)
def test_factory_table_row(
    code: str,
    frost: str,
    granite_w: str,
    ordinary_w: str | None,
) -> None:
    assert suggest_concrete_spec(code) == ConcreteSpec(
        frost_resistance=frost,
        waterproofness=granite_w,
        concrete_aggregate="granite",
        ordinary_waterproofness=ordinary_w,
    )


def test_new_b25_line_gets_granite_table_pair() -> None:
    new_line = {
        "line_id": "new-1",
        "concrete_grade": "B25",
        "unit_price": 44634.03,
        "mark": "С120.35-12",
    }
    [attached] = apply_line_specs([], [new_line])
    assert attached["frost_resistance"] == "F200"
    assert attached["waterproofness"] == "W8"
    assert attached["concrete_aggregate"] == "granite"
    assert attached["concrete_spec_source"] == "table"
    assert attached["concrete_grade"] == "B25"
    assert attached["unit_price"] == 44634.03


def test_table_ordinary_recalculates_to_b22_5_w4() -> None:
    previous = {
        "line_id": "row-1",
        "concrete_grade": "B25",
        "unit_price": 100.0,
        "frost_resistance": "F200",
        "waterproofness": "W6",
        "concrete_aggregate": "ordinary",
        "concrete_spec_source": "table",
    }
    new_line = {
        "line_id": "row-1",
        "concrete_grade": "B22.5",
        "unit_price": 90.0,
    }
    [attached] = apply_line_specs([previous], [new_line])
    assert attached["frost_resistance"] == "F200"
    assert attached["waterproofness"] == "W4"
    assert attached["concrete_aggregate"] == "ordinary"
    assert attached["concrete_spec_source"] == "table"
    assert attached["concrete_grade"] == "B22.5"
    assert attached["unit_price"] == 90.0


def test_ordinary_on_b30_falls_back_to_granite_table() -> None:
    previous = {
        "line_id": "row-1",
        "concrete_grade": "B25",
        "frost_resistance": "F200",
        "waterproofness": "W6",
        "concrete_aggregate": "ordinary",
        "concrete_spec_source": "table",
        "unit_price": 10,
    }
    new_line = {"line_id": "row-1", "concrete_grade": "B30", "unit_price": 12}
    [attached] = apply_line_specs([previous], [new_line])
    assert attached["frost_resistance"] == "F300"
    assert attached["waterproofness"] == "W10"
    assert attached["concrete_aggregate"] == "granite"
    assert attached["concrete_spec_source"] == "table"
    assert attached["concrete_grade"] == "B30"
    assert attached["unit_price"] == 12


def test_manual_pair_survives_grade_change() -> None:
    previous = {
        "line_id": "row-1",
        "concrete_grade": "B25",
        "frost_resistance": "F150",
        "waterproofness": "W4",
        "concrete_aggregate": "",
        "concrete_spec_source": "manual",
        "unit_price": 10,
    }
    new_line = {"line_id": "row-1", "concrete_grade": "B20", "unit_price": 8}
    [attached] = apply_line_specs([previous], [new_line])
    assert attached["frost_resistance"] == "F150"
    assert attached["waterproofness"] == "W4"
    assert attached["concrete_spec_source"] == "manual"
    assert attached["concrete_aggregate"] in ("", None)
    assert attached["concrete_grade"] == "B20"
    assert attached["unit_price"] == 8


def test_empty_snapshot_stays_empty_for_same_line_id() -> None:
    previous = {
        "line_id": "old-1",
        "concrete_grade": "B25",
        "unit_price": 10,
        "frost_resistance": None,
        "waterproofness": None,
        "concrete_aggregate": None,
        "concrete_spec_source": None,
    }
    new_line = {"line_id": "old-1", "concrete_grade": "B20", "unit_price": 11}
    [attached] = apply_line_specs([previous], [new_line])
    assert attached.get("frost_resistance") in (None, "")
    assert attached.get("waterproofness") in (None, "")
    assert attached.get("concrete_aggregate") in (None, "")
    assert attached.get("concrete_spec_source") in (None, "")
    assert attached["concrete_grade"] == "B20"
    assert attached["unit_price"] == 11


def test_apply_line_specs_does_not_match_by_index() -> None:
    previous = {
        "line_id": "kept",
        "concrete_grade": "B25",
        "frost_resistance": "F150",
        "waterproofness": "W4",
        "concrete_aggregate": "",
        "concrete_spec_source": "manual",
    }
    new_line = {"line_id": "other", "concrete_grade": "B25", "unit_price": 5}
    [attached] = apply_line_specs([previous], [new_line])
    assert attached["frost_resistance"] == "F200"
    assert attached["waterproofness"] == "W8"
    assert attached["concrete_spec_source"] == "table"
    assert attached["unit_price"] == 5
