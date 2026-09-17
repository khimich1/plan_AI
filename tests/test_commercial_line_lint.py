"""Per-physical-line lint for commercial source text (no preview / ILP / DraftStore)."""

from __future__ import annotations

import inspect

import pytest

from app.schemas.commercial import ProductType
from app.services.commercial_line_lint import LineLint, lint_source_lines


CASES: list[tuple[ProductType, str, str]] = [
    ("plates", "ПБ 78-12-8п 2", "xyz-not-a-plate"),
    ("piles", "С120.35-12 B25 5", "ПБ 78-12-8п 2"),
    ("steps", "ЛС11 10", "С120.35-12 5"),
    ("marches", "1ЛМ 27-11-14-4 2", "ПБ 78-12-8п 2"),
    ("bridge_piles", "С7-35Т5", "С120.35-12 2"),
    ("fbs", "ФБС 9.3.6-Т 2", "С120.35-12 2"),
]


def _by_index(lines: list[LineLint]) -> dict[int, LineLint]:
    return {line.index: line for line in lines}


@pytest.mark.parametrize(("product_type", "ok_line", "bad_line"), CASES)
def test_lint_ok_and_not_ok_for_each_product_type(
    product_type: ProductType, ok_line: str, bad_line: str
) -> None:
    lines = lint_source_lines(f"{ok_line}\n{bad_line}", product_type)
    assert len(lines) == 2
    assert lines[0].index == 0
    assert lines[0].text == ok_line
    assert lines[0].empty is False
    assert lines[0].ok is True
    assert lines[0].reason_text is None
    assert lines[1].index == 1
    assert lines[1].text == bad_line
    assert lines[1].empty is False
    assert lines[1].ok is False
    assert lines[1].reason_text


def test_plates_slash_format_is_not_ok() -> None:
    lines = lint_source_lines("ПБ 40,3/2,6-8п", "plates")
    assert len(lines) == 1
    assert lines[0].ok is False
    assert lines[0].empty is False
    assert lines[0].text == "ПБ 40,3/2,6-8п"


def test_empty_and_whitespace_lines_are_ok() -> None:
    text = "ПБ 78-12-8п 2\n\n  \nплохо"
    lines = lint_source_lines(text, "plates")
    by_index = _by_index(lines)
    assert by_index[1].text == ""
    assert by_index[1].empty is True
    assert by_index[1].ok is True
    assert by_index[1].reason_text is None
    assert by_index[2].text == "  "
    assert by_index[2].empty is True
    assert by_index[2].ok is True
    assert by_index[2].reason_text is None
    assert by_index[3].ok is False
    assert by_index[3].empty is False


@pytest.mark.parametrize(
    ("product_type", "raw"),
    [
        ("bridge_piles", "C14-35T7 80 шт"),
        ("bridge_piles", "C14-35T7 80шт"),
        ("bridge_piles", "C 14.35-T7 80 шт"),
        ("plates", "ПБ 78-12-8п 5шт"),
        ("piles", "С120.35-12 5шт"),
        ("piles", "С110.30-9у"),
        ("piles", "C 110.30-9у"),
        ("piles", "С120.35-13и"),
        ("piles", "С110.30-6у"),
        ("fbs", "ФБС 9.3.6-Т 2шт"),
        ("steps", "ЛС11 2шт"),
        ("marches", "1ЛМ 27-11-14-4 2шт"),
    ],
)
def test_lint_prepare_makes_dirty_qty_ok(product_type: ProductType, raw: str) -> None:
    lines = lint_source_lines(raw, product_type)
    assert len(lines) == 1
    assert lines[0].ok is True
    assert lines[0].text == raw


def test_lint_bridge_gost_without_class_not_ok() -> None:
    lines = lint_source_lines("C 14.35 80", "bridge_piles")
    assert lines[0].ok is False
    assert lines[0].text == "C 14.35 80"


def test_lint_bridge_mark_on_plates_not_ok() -> None:
    raw = "C14-35T7 80шт"
    lines = lint_source_lines(raw, "plates")
    assert lines[0].ok is False
    assert lines[0].text == raw


def test_lint_slash_plate_still_not_ok() -> None:
    lines = lint_source_lines("ПБ 40,3/2,6-8п", "plates")
    assert lines[0].ok is False
    assert lines[0].text == "ПБ 40,3/2,6-8п"


def test_lint_module_does_not_call_preview_optimize_or_draft_store() -> None:
    import ast

    import app.services.commercial_line_lint as mod

    tree = ast.parse(inspect.getsource(mod))
    imported_modules: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported_modules.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported_modules.append(node.module)
    joined = " ".join(imported_modules)
    assert "draft_store" not in joined
    assert "commercial_service" not in joined
    assert "optimization_service" not in joined
    assert "product_draft" not in joined
