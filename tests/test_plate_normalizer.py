#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Регрессионные тесты для plate_text_normalizer и интеграции с парсером.

Покрывает:
  - Каталожные марки (ПБ L.W-loadВр…)
  - OCR-ошибки в префиксе (ПВ→ПБ)
  - Каноническое написание — не должно меняться
  - Формат WxL — не должен меняться
  - Смешанный ввод
  - Крайние случаи и негативные кейсы
  - Сквозные тесты через set_plate_lists_from_text
"""
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        import codecs
        sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer, "strict")

from core.plate_text_normalizer import (
    canonicalize_plate_line,
    normalize_order_text,
    parse_catalog_mark,
    _format_width_dm,
    _format_load,
)
from core.plate_line_parser import parse_line
import core.config_and_data as cfg


# ── Вспомогательные функции ────────────────────────────────────────────────

def _assert_eq(actual, expected, label: str):
    if actual == expected:
        print(f"  OK  {label}: {actual!r}")
    else:
        print(f"  FAIL {label}:")
        print(f"       expected: {expected!r}")
        print(f"       got:      {actual!r}")
        raise AssertionError(f"{label}: expected {expected!r}, got {actual!r}")


def _assert_close(actual: float, expected: float, tol: float, label: str):
    if abs(actual - expected) <= tol:
        print(f"  OK  {label}: {actual}")
    else:
        print(f"  FAIL {label}: expected ≈{expected}, got {actual}")
        raise AssertionError(f"{label}: expected ≈{expected}, got {actual}")


# ── Блок 1: форматирующие утилиты ──────────────────────────────────────────

def test_format_width_dm():
    print("\n=== test_format_width_dm ===")
    _assert_eq(_format_width_dm(12), "12",  "12дм")
    _assert_eq(_format_width_dm(10), "10",  "10дм")
    _assert_eq(_format_width_dm(5),  "5,0", "5дм")
    _assert_eq(_format_width_dm(2),  "2,0", "2дм")
    _assert_eq(_format_width_dm(9),  "9,0", "9дм")


def test_format_load():
    print("\n=== test_format_load ===")
    _assert_eq(_format_load(8.0),  "8п",    "8")
    _assert_eq(_format_load(10.0), "10п",   "10")
    _assert_eq(_format_load(12.5), "12,5п", "12.5")
    _assert_eq(_format_load(6.0),  "6п",    "6")


# ── Блок 2: parse_catalog_mark ─────────────────────────────────────────────

def test_parse_catalog_mark_basic():
    print("\n=== test_parse_catalog_mark ===")
    # «ПБ 59.12-8Вр1400-25» → L=59, W=12, load=8, qty=1
    r = parse_catalog_mark("ПБ 59.12-8Вр1400-25")
    assert r is not None, "Должно распознать ПБ 59.12-8Вр1400-25"
    prefix, L, W, load, qty = r
    _assert_eq(prefix, "ПБ",  "prefix")
    _assert_eq(L, 59,          "L_dm")
    _assert_eq(W, 12,          "W_dm")
    _assert_close(load, 8.0, 1e-6, "load")
    _assert_eq(qty, 1,         "qty")


def test_parse_catalog_mark_no_suffix():
    print("\n=== test_parse_catalog_mark_no_suffix ===")
    # «ПБ56.05-10» → L=56, W=5, load=10, qty=1
    r = parse_catalog_mark("ПБ56.05-10")
    assert r is not None, "Должно распознать ПБ56.05-10"
    prefix, L, W, load, qty = r
    _assert_eq(prefix, "ПБ", "prefix")
    _assert_eq(L, 56,         "L_dm")
    _assert_eq(W, 5,          "W_dm")
    _assert_close(load, 10.0, 1e-6, "load")
    _assert_eq(qty, 1,        "qty")


def test_parse_catalog_mark_with_qty():
    print("\n=== test_parse_catalog_mark_with_qty ===")
    r = parse_catalog_mark("ПБ 59.12-8Вр1400-25 5")
    assert r is not None
    _, _, _, _, qty = r
    _assert_eq(qty, 5, "qty=5")

    r2 = parse_catalog_mark("пб56.05-10 3 шт")
    assert r2 is not None
    _, _, _, _, qty2 = r2
    _assert_eq(qty2, 3, "qty=3 шт")


def test_parse_catalog_mark_ocr_prefix():
    print("\n=== test_parse_catalog_mark_ocr_prefix ===")
    # ПВ → ПБ (OCR-ошибка)
    from core.plate_text_normalizer import normalize_plate_prefixes
    cleaned = normalize_plate_prefixes("ПВ 59.12-8Вр1400-25")
    r = parse_catalog_mark(cleaned)
    assert r is not None, "ПВ должно быть исправлено до ПБ"
    prefix, _, _, _, _ = r
    _assert_eq(prefix, "ПБ", "prefix после исправления ПВ→ПБ")


def test_parse_catalog_mark_pk():
    print("\n=== test_parse_catalog_mark_pk (ПК) ===")
    r = parse_catalog_mark("ПК 80.12-8Вр1400-30")
    assert r is not None
    prefix, L, W, _, _ = r
    _assert_eq(prefix, "ПК", "ПК сохранён")
    _assert_eq(L, 80,         "L=80")
    _assert_eq(W, 12,         "W=12")


def test_parse_catalog_mark_guards():
    print("\n=== test_parse_catalog_mark_guards (негативные) ===")
    # L_dm < 20: «ПБ 7.3-12-8п» — должно вернуть None
    r = parse_catalog_mark("ПБ 7.3-12-8п")
    assert r is None, "ПБ 7.3-12-8п не должно распознаться как каталожная (L=7 < 20)"
    print("  OK  ПБ 7.3-12-8п → None (L_dm=7 < 20)")

    # Обычная марка без точки: «ПБ 78-12-8п» — None
    r2 = parse_catalog_mark("ПБ 78-12-8п")
    assert r2 is None, "ПБ 78-12-8п не должно распознаться как каталожная"
    print("  OK  ПБ 78-12-8п → None (нет точки L.W)")

    # W_dm > 15: нереальная ширина → None
    r3 = parse_catalog_mark("ПБ 56.20-8")
    assert r3 is None, "ПБ 56.20-8 не должно распознаться (W=20 > 15)"
    print("  OK  ПБ 56.20-8 → None (W_dm=20 > 15)")


# ── Блок 3: canonicalize_plate_line ────────────────────────────────────────

def test_canonicalize_catalog():
    print("\n=== test_canonicalize_catalog ===")
    canon, warn = canonicalize_plate_line("ПБ 59.12-8Вр1400-25")
    _assert_eq(canon, "ПБ 59-12-8п", "каталог → канон")
    assert warn is not None, "Предупреждение должно быть"


def test_canonicalize_catalog_no_suffix():
    print("\n=== test_canonicalize_catalog_no_suffix ===")
    canon, warn = canonicalize_plate_line("ПБ56.05-10")
    _assert_eq(canon, "ПБ 56-5,0-10п", "ПБ56.05-10 → ПБ 56-5,0-10п")
    assert warn is not None


def test_canonicalize_canonical_unchanged():
    print("\n=== test_canonicalize_canonical_unchanged ===")
    # Обычные марки должны проходить без изменений (кроме базовой чистки)
    for line in [
        "ПБ 66,2-12-8п 6",
        "ПБ 78-12-8п 3",
        "ПБ 44-3,2-10п 5",
        "Плиты ПБ 73-12-8п — 93 шт",
    ]:
        canon, warn = canonicalize_plate_line(line)
        assert warn is None, f"Обычная марка {line!r} не должна менять: warn={warn}"
        print(f"  OK  {line!r} → без изменений")


def test_canonicalize_wxl_unchanged():
    print("\n=== test_canonicalize_wxl_unchanged ===")
    canon, warn = canonicalize_plate_line("1.2x3.39 — 2 шт")
    assert warn is None, "WxL не должен трогаться"
    # Тире должно нормализоваться
    assert "—" not in canon, "Тире должно быть заменено на дефис"
    print(f"  OK  WxL: {canon!r}")


def test_canonicalize_with_qty():
    print("\n=== test_canonicalize_with_qty ===")
    canon, _ = canonicalize_plate_line("ПБ 59.12-8Вр1400-25 5")
    _assert_eq(canon, "ПБ 59-12-8п 5", "с кол-вом 5")

    canon2, _ = canonicalize_plate_line("пб56.05-10 3 шт")
    _assert_eq(canon2, "ПБ 56-5,0-10п 3", "с кол-вом 3 шт")


def test_canonicalize_ocr_prefix_fix():
    print("\n=== test_canonicalize_ocr_prefix_fix ===")
    canon, warn = canonicalize_plate_line("ПВ 59.12-8Вр1400-25")
    _assert_eq(canon, "ПБ 59-12-8п", "ПВ→ПБ при каталожной марке")


# ── Блок 4: normalize_order_text ───────────────────────────────────────────

def test_normalize_order_text_mixed():
    print("\n=== test_normalize_order_text_mixed ===")
    text = (
        "ПБ 59.12-8Вр1400-25 5\n"
        "ПБ56.05-10 3\n"
        "ПБ 66,2-12-8п 6\n"
        "1.2x3.39 — 2 шт"
    )
    result = normalize_order_text(text)
    lines = result.normalized_lines
    assert len(lines) == 4, f"Должно быть 4 строки, получено {len(lines)}"
    _assert_eq(lines[0], "ПБ 59-12-8п 5",  "каталог с qty")
    _assert_eq(lines[1], "ПБ 56-5,0-10п 3", "каталог без суффикса с qty")
    # Обычные строки — без изменений
    assert "66,2" in lines[2], "Обычная марка не изменилась"
    assert "1.2x3.39" in lines[3] or "1.2x3.39" in lines[3].replace("—", "-"), "WxL не изменился"
    assert len(result.warnings) == 2, "2 предупреждения для 2 каталожных марок"
    print(f"  OK  {len(result.warnings)} предупреждения, строки: {lines}")


def test_normalize_empty_and_single():
    print("\n=== test_normalize_empty ===")
    r = normalize_order_text("")
    assert r.normalized_text == ""
    print("  OK  пустая строка")

    r2 = normalize_order_text("ПБ 59.12-8Вр1400-25")
    _assert_eq(r2.normalized_lines[0], "ПБ 59-12-8п", "одна строка")


# ── Блок 5: сквозные тесты через set_plate_lists_from_text ─────────────────

def test_end_to_end_catalog_mark():
    print("\n=== test_end_to_end_catalog_mark ===")
    text = "ПБ 59.12-8Вр1400-25"
    unparsed_lines, _, _ = cfg.set_plate_lists_from_text(text)
    assert not unparsed_lines, f"Нераспознанных строк не должно быть: {unparsed_lines}"

    # Ожидаем: длина 5.9м (59дм), ширина 1.2м (12дм), нагрузка 8
    plates = cfg.PLATES_1_2
    assert len(plates) == 1, f"Должна быть 1 плита 1.2м, получено {len(plates)}"
    _assert_close(plates[0], 5.9, 0.01, "длина плиты (59дм→5.9м)")

    # PLATE_LOAD_DETAILS должен содержать запись
    found = [
        (k, v) for k, v in cfg.PLATE_LOAD_DETAILS.items()
        if abs(k[0] - 5.9) < 0.01 and abs(k[1] - 1.2) < 0.01
    ]
    assert found, "PLATE_LOAD_DETAILS должен содержать плиту 5.9×1.2"
    _, qty = found[0]
    _assert_eq(qty, 1, "qty=1")
    load = found[0][0][2]
    _assert_close(float(load), 8.0, 0.01, "load=8")
    print(f"  OK  PLATE_LOAD_DETAILS: {found[0]}")


def test_end_to_end_catalog_no_suffix():
    print("\n=== test_end_to_end_catalog_no_suffix ===")
    text = "ПБ56.05-10 3"
    unparsed_lines, _, _ = cfg.set_plate_lists_from_text(text)
    assert not unparsed_lines, f"Нераспознанных строк: {unparsed_lines}"

    # Ожидаем: длина 5.6м (56дм), ширина 0.5м (5дм), нагрузка 10, qty=3
    found = [
        (k, v) for k, v in cfg.PLATE_LOAD_DETAILS.items()
        if abs(k[0] - 5.6) < 0.01 and abs(k[1] - 0.5) < 0.01
    ]
    assert found, "PLATE_LOAD_DETAILS должен содержать плиту 5.6×0.5"
    _, qty = found[0]
    _assert_eq(qty, 3, "qty=3")
    load = found[0][0][2]
    _assert_close(float(load), 10.0, 0.01, "load=10")
    print(f"  OK  PLATE_LOAD_DETAILS: {found[0]}")


def test_end_to_end_canonical_unchanged():
    print("\n=== test_end_to_end_canonical_unchanged ===")
    # Обычный ввод должен работать точно так же, как раньше
    text = "ПБ 66,2-12-8п 6"
    unparsed_lines, _, _ = cfg.set_plate_lists_from_text(text)
    assert not unparsed_lines, f"Нераспознанных строк: {unparsed_lines}"

    found = [
        (k, v) for k, v in cfg.PLATE_LOAD_DETAILS.items()
        if abs(k[0] - 6.62) < 0.01 and abs(k[1] - 1.2) < 0.01
    ]
    assert found, "PLATE_LOAD_DETAILS должен содержать плиту 6.62×1.2"
    _, qty = found[0]
    _assert_eq(qty, 6, "qty=6")
    print(f"  OK  {found[0]}")


def test_end_to_end_mixed_order():
    print("\n=== test_end_to_end_mixed_order ===")
    text = (
        "ПБ 59.12-8Вр1400-25 5\n"
        "ПБ56.05-10 3\n"
        "ПБ 66,2-12-8п 6"
    )
    unparsed_lines, _, _ = cfg.set_plate_lists_from_text(text)
    assert not unparsed_lines, f"Нераспознанных строк: {unparsed_lines}"

    total = sum(cfg.PLATE_LOAD_DETAILS.values())
    _assert_eq(total, 14, "итого 5+3+6=14 плит")
    print(f"  OK  всего плит: {total}")


def test_end_to_end_wrong_not_changed():
    """
    Строка «ПБ 7.3-12-8п» — обычный формат с дробной длиной.
    Нормализатор не должен её трогать (L=7 < 20).
    Парсер должен обработать её в штатном режиме.
    """
    print("\n=== test_end_to_end_regular_decimal_length ===")
    # ПБ 7.3-12-8п: длина 7.3дм = 0.73м через length_dm_to_m("7.3") — нереальная
    # Нормализатор не трогает; парсер разбирает как есть.
    canon, warn = canonicalize_plate_line("ПБ 7.3-12-8п 2")
    assert warn is None, f"ПБ 7.3-12-8п не должна трогаться нормализатором, warn={warn}"
    print("  OK  ПБ 7.3-12-8п → нормализатор не трогает")


# ── NM-101: каталог vs десятичная длина ────────────────────────────────────

_CATALOG_CASES = [
    ("ПБ 59.12-8Вр1400-25", "ПБ 59-12-8п", 1),
    ("ПБ56.05-10", "ПБ 56-5,0-10п", 1),
    ("ПБ56.05-10 3 шт", "ПБ 56-5,0-10п 3", 3),
    ("ПБ 59.12-8п", "ПБ 59-12-8п", 1),
]

_NOT_CATALOG_CASES = [
    "ПБ 70.5-12-8п",
    "ПБ 70.5-12-8п 1",
    "ПБ 47.5-10.7-8  4",
    "ПБ 70.5-10.7-8п доб. 70.5-1.25-8",
    "ПБ 47.5-5.3-8п 1",
    "ПБ 47.5-12-8п 2",
    "ПБ 60.5-12-8п 2",
    "ПБ 60.5-10.7-8п 1",
]


def _decimal_length_token(line: str) -> str:
    for token in ("70.5", "47.5", "60.5"):
        if token in line:
            return token
    raise AssertionError(f"нет десятичной длины 70.5/47.5/60.5 в {line!r}")


@pytest.mark.parametrize("raw, expected_canon, expected_qty", _CATALOG_CASES)
def test_parse_catalog_mark_accepts_factory_catalog(raw, expected_canon, expected_qty):
    parsed = parse_catalog_mark(raw)
    assert parsed is not None, f"каталог должен распознаться: {raw!r}"
    _, _, _, _, qty = parsed
    assert qty == expected_qty

    canon, warn = canonicalize_plate_line(raw)
    assert canon == expected_canon
    assert warn is not None
    assert "→" in warn


@pytest.mark.parametrize("raw", _NOT_CATALOG_CASES)
def test_parse_catalog_mark_rejects_decimal_length(raw):
    assert parse_catalog_mark(raw) is None

    canon, warn = canonicalize_plate_line(raw)
    assert warn is None
    token = _decimal_length_token(raw)
    comma = token.replace(".", ",")
    assert token in canon or comma in canon, (
        f"канон должен сохранить десятичную длину {token} или {comma}, получено {canon!r}"
    )
    assert f"{token.split('.')[0]}-{token.split('.')[1]},0" not in canon


# ── NM-102: эталонная ведомость ────────────────────────────────────────────

_REFERENCE_LEDGER = (
    "ПБ 58-12-8п 4\n"
    "ПБ 58-10.7-8 2\n"
    "ПБ 47.5-5.3-8п 1\n"
    "ПБ 47.5-5.3-8п 1\n"
    "ПБ 47.5-5.3-8п 1\n"
    "ПБ 47.5-5.3-8п 1\n"
    "ПБ 70.5-10.7-8п 3\n"
    "ПБ 70.5-12-8п 1\n"
    "ПБ 43.0-12-8п 5\n"
    "ПБ 43.0-10.7-8 4\n"
    "ПБ 47.5-10.7-8  4\n"
    "ПБ 47.5-12-8п 2\n"
    "ПБ 60.5-12-8п 2\n"
    "ПБ 60.5-10.7-8п 1\n"
)

_REFERENCE_EXPECTED = [
    (5.8, 1.2, 8.0, 4),
    (5.8, 1.07, 8.0, 2),
    (4.75, 0.53, 8.0, 1),
    (4.75, 0.53, 8.0, 1),
    (4.75, 0.53, 8.0, 1),
    (4.75, 0.53, 8.0, 1),
    (7.05, 1.07, 8.0, 3),
    (7.05, 1.2, 8.0, 1),
    (4.3, 1.2, 8.0, 5),
    (4.3, 1.07, 8.0, 4),
    (4.75, 1.07, 8.0, 4),
    (4.75, 1.2, 8.0, 2),
    (6.05, 1.2, 8.0, 2),
    (6.05, 1.07, 8.0, 1),
]

_FALSE_CATALOG_ARROWS = (
    ("47.5", "47-5,0"),
    ("70.5", "70-5,0"),
    ("60.5", "60-5,0"),
)


def test_reference_ledger_keeps_sizes_and_qty_32():
    result = normalize_order_text(_REFERENCE_LEDGER)
    assert len(result.normalized_lines) == 14

    parsed_rows = [parse_line(line) for line in result.normalized_lines]
    assert all(row.parsed for row in parsed_rows)
    assert sum(row.qty for row in parsed_rows) == 32

    for idx, (row, expected) in enumerate(zip(parsed_rows, _REFERENCE_EXPECTED)):
        length_m, width_m, load, qty = expected
        assert row.length_m == pytest.approx(length_m, abs=1e-3), f"строка {idx + 1} длина"
        assert row.width_m == pytest.approx(width_m, abs=1e-3), f"строка {idx + 1} ширина"
        assert row.load_code == pytest.approx(load, abs=1e-6), f"строка {idx + 1} нагрузка"
        assert row.qty == qty, f"строка {idx + 1} qty"

    joined_warnings = " ".join(result.warnings)
    for src, dst in _FALSE_CATALOG_ARROWS:
        assert not (src in joined_warnings and dst in joined_warnings), (
            f"ложный каталожный warning {src!r} → {dst!r}: {result.warnings!r}"
        )


def test_comma_and_dot_decimal_length_parse_the_same():
    comma_line, _ = canonicalize_plate_line("ПБ 47,5-12-8п 2")
    dot_line, warn_dot = canonicalize_plate_line("ПБ 47.5-12-8п 2")
    assert warn_dot is None

    comma = parse_line(comma_line)
    dot = parse_line(dot_line)
    assert comma.parsed and dot.parsed
    assert comma.length_m == pytest.approx(dot.length_m)
    assert comma.width_m == pytest.approx(dot.width_m)
    assert comma.load_code == pytest.approx(dot.load_code)
    assert comma.qty == dot.qty == 2
    assert comma.length_m == pytest.approx(4.75, abs=1e-3)
    assert comma.width_m == pytest.approx(1.2, abs=1e-3)
    assert comma.load_code == pytest.approx(8.0, abs=1e-6)


# ── NM-301: сторож канона ──────────────────────────────────────────────────

def test_catalog_tuple_matches_canonical_accepts_legit_rewrite():
    from core.plate_text_normalizer import _catalog_tuple_matches_canonical

    assert _catalog_tuple_matches_canonical(59, 12, 8.0, 1, "ПБ 59-12-8п")
    assert _catalog_tuple_matches_canonical(56, 5, 10.0, 3, "ПБ 56-5,0-10п 3")


def test_catalog_tuple_matches_canonical_rejects_mismatch():
    from core.plate_text_normalizer import _catalog_tuple_matches_canonical

    # Кортеж каталога 70×5дм / 12п не совпадает с обычной маркой 70.5-12-8п
    assert not _catalog_tuple_matches_canonical(70, 5, 12.0, 1, "ПБ 70.5-12-8п 1")
    assert not _catalog_tuple_matches_canonical(59, 12, 8.0, 1, "не плита")


# ── Запуск ──────────────────────────────────────────────────────────────────

def run_all():
    tests = [
        test_format_width_dm,
        test_format_load,
        test_parse_catalog_mark_basic,
        test_parse_catalog_mark_no_suffix,
        test_parse_catalog_mark_with_qty,
        test_parse_catalog_mark_ocr_prefix,
        test_parse_catalog_mark_pk,
        test_parse_catalog_mark_guards,
        test_canonicalize_catalog,
        test_canonicalize_catalog_no_suffix,
        test_canonicalize_canonical_unchanged,
        test_canonicalize_wxl_unchanged,
        test_canonicalize_with_qty,
        test_canonicalize_ocr_prefix_fix,
        test_normalize_order_text_mixed,
        test_normalize_empty_and_single,
        test_end_to_end_catalog_mark,
        test_end_to_end_catalog_no_suffix,
        test_end_to_end_canonical_unchanged,
        test_end_to_end_mixed_order,
        test_end_to_end_wrong_not_changed,
    ]

    passed = failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except Exception as e:
            print(f"\n  *** FAIL: {t.__name__}: {e}")
            failed += 1

    print("\n" + "=" * 60)
    if failed == 0:
        print(f"SUCCESS! Все {passed} тестов прошли.")
    else:
        print(f"ИТОГ: {passed} прошли, {failed} упали.")
    return failed


if __name__ == "__main__":
    exit_code = run_all()
    sys.exit(exit_code)
