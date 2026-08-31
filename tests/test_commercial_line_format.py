"""TDD (MNA-002): format_line_name — mark + optional grade in parentheses."""

from core.commercial_line_format import format_line_name


def test_format_line_name_pile_mark_with_grade() -> None:
    """Свая: марка + класс бетона в скобках — как в спеке КП."""
    item = {
        "product_type": "piles",
        "mark": "С30.15-3",
        "concrete_grade": "B25",
        "qty": 12,
    }
    assert format_line_name(item) == "С30.15-3 (B25)"


def test_format_line_name_bridge_pile_mark_with_grade() -> None:
    """Мостовая свая: тот же shape mark + concrete_grade."""
    item = {
        "product_type": "bridge_piles",
        "mark": "СМ150.30-6",
        "concrete_grade": "B30",
        "qty": 4,
    }
    assert format_line_name(item) == "СМ150.30-6 (B30)"


def test_format_line_name_uses_name_when_mark_absent() -> None:
    """Fallback mark ← name, как в commercial_offer (mark or name)."""
    item = {
        "product_type": "piles",
        "name": "С30.15-3",
        "concrete_grade": "B25",
    }
    assert format_line_name(item) == "С30.15-3 (B25)"


def test_format_line_name_no_grade_returns_mark_only() -> None:
    """Нет grade → только марка, без скобок и без дефолта B25 в имени."""
    item = {"product_type": "piles", "mark": "С30.15-3"}
    assert format_line_name(item) == "С30.15-3"


def test_format_line_name_empty_grade_returns_mark_only() -> None:
    item = {
        "product_type": "piles",
        "mark": "С30.15-3",
        "concrete_grade": "",
    }
    assert format_line_name(item) == "С30.15-3"


def test_format_line_name_whitespace_grade_returns_mark_only() -> None:
    item = {
        "product_type": "piles",
        "mark": "С30.15-3",
        "concrete_grade": "   ",
    }
    assert format_line_name(item) == "С30.15-3"


def test_format_line_name_plate_without_grade_is_name_only() -> None:
    """Плита без grade → наименование как сегодня (только name/mark)."""
    item = {
        "product_type": "plates",
        "name": "ПБ 56-6-8п",
        "length_m": 5.6,
        "width_m": 0.6,
        "qty": 1,
        "load_class": 800,
    }
    assert format_line_name(item) == "ПБ 56-6-8п"


def test_format_line_name_plate_mark_without_grade() -> None:
    item = {"product_type": "plates", "mark": "ПБ 68-12-8п", "qty": 2}
    assert format_line_name(item) == "ПБ 68-12-8п"


def test_format_line_name_step_without_grade_is_mark_only() -> None:
    """Ступени: grade в прайсе нет — только марка."""
    item = {"product_type": "steps", "mark": "ЛС-12", "qty": 10}
    assert format_line_name(item) == "ЛС-12"


def test_format_line_name_empty_mark_does_not_crash() -> None:
    """Пустая марка — безопасный fallback, без ведущего пробела."""
    item = {"product_type": "piles", "mark": "", "concrete_grade": "B25"}
    assert format_line_name(item) == "(B25)"


def test_format_line_name_whitespace_mark_with_grade() -> None:
    """Whitespace-only mark strips to empty → same as empty mark."""
    item = {"product_type": "piles", "mark": "   ", "concrete_grade": "B25"}
    assert format_line_name(item) == "(B25)"


def test_format_line_name_missing_mark_and_name_does_not_crash() -> None:
    item = {"product_type": "piles", "concrete_grade": "B25"}
    assert format_line_name(item) == "(B25)"


def test_format_line_name_none_mark_does_not_crash() -> None:
    item = {"mark": None, "concrete_grade": "B25"}
    assert format_line_name(item) == "(B25)"


def test_format_line_name_fbs_and_march_with_grade() -> None:
    """FBS / марши: тот же concrete_grade, что в pricing lookup."""
    fbs = {
        "product_type": "fbs",
        "mark": "ФБС 24.4.6",
        "concrete_grade": "B15",
    }
    march = {
        "product_type": "marches",
        "mark": "ЛМ 27.12.14-4",
        "concrete_grade": "B25",
    }
    assert format_line_name(fbs) == "ФБС 24.4.6 (B15)"
    assert format_line_name(march) == "ЛМ 27.12.14-4 (B25)"
