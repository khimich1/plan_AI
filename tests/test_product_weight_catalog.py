"""Контракт справочника весов ФБС/ЛС/ЛМ: parse, validate, resolve, upsert."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from core.product_weight_catalog import (
    parse_product_weights_from_xlsx,
    resolve_product_weight_kg,
    upsert_product_weights,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
REAL_DIR = REPO_ROOT / "банк знаний" / "Новая папка"
REAL_FBS = REAL_DIR / "Блоки.xlsx"
REAL_STEPS = REAL_DIR / "ЛС.xlsx"
REAL_MARCHES = REAL_DIR / "ЛМ и ЛП.xlsx"


def _write_xlsx(path: Path, rows: list[tuple]) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист_1"
    ws.append(["Наименование", "Вес (числитель)", "Объем (числитель)", "Код"])
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    return str(path)


def test_parse_real_fbs_xlsx_keeps_14_and_skips_cross_blocks() -> None:
    if not REAL_FBS.is_file():
        pytest.skip(f"нет файла {REAL_FBS}")
    result = parse_product_weights_from_xlsx(str(REAL_FBS), "fbs")
    assert len(result.records) == 14
    assert all(row.product_type == "fbs" for row in result.records)
    assert any("24.6.6" in row.display_name for row in result.records)
    assert len(result.skipped) == 2
    skipped_text = " ".join(
        f"{item.display_name} {item.reason}" for item in result.skipped
    )
    assert "Кросс" in skipped_text or "КБ" in skipped_text
    assert result.quarantine == []


def test_parse_real_steps_xlsx_quarantines_density_artifact() -> None:
    if not REAL_STEPS.is_file():
        pytest.skip(f"нет файла {REAL_STEPS}")
    result = parse_product_weights_from_xlsx(str(REAL_STEPS), "steps")
    assert len(result.records) == 18
    assert len(result.quarantine) == 1
    assert len(result.records) + len(result.quarantine) == 19
    quarantined = result.quarantine[0]
    assert "ЛС-15-1" in quarantined.display_name
    assert "закл" in quarantined.display_name.lower()
    assert "плотность" in quarantined.reason.lower()


def test_parse_real_marches_xlsx_quarantines_density_and_skips_lp() -> None:
    if not REAL_MARCHES.is_file():
        pytest.skip(f"нет файла {REAL_MARCHES}")
    result = parse_product_weights_from_xlsx(str(REAL_MARCHES), "marches")
    # В файле 9 ЛМ с валидной плотностью (оценка «7» в плане — покрытие прайса 7/7).
    assert len(result.records) == 9
    assert all(row.product_type == "marches" for row in result.records)
    assert len(result.quarantine) == 3
    q_names = " ".join(item.display_name for item in result.quarantine)
    assert "ЛП 28-15-5ш" in q_names
    assert "1ЛП 28-15-5ш-1" in q_names
    assert "ЛМ 30.13.15-5ш" in q_names
    assert all("плотность" in item.reason.lower() for item in result.quarantine)
    skipped_text = " ".join(
        f"{item.display_name} {item.reason}" for item in result.skipped
    )
    assert "ЛП" in skipped_text or "площадк" in skipped_text.lower()


def test_parse_quarantines_density_out_of_range(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "density.xlsx",
        [
            ("Блоки ФБС 9.3.6-Т", 350, 0.146, "00-00077237"),
            ("Блоки ФБС 9.9.9-Т", 168, 0.67, "00-00003520"),
        ],
    )
    result = parse_product_weights_from_xlsx(xlsx, "fbs")
    assert len(result.records) == 1
    assert result.records[0].weight_kg == pytest.approx(350.0)
    assert len(result.quarantine) == 1
    assert "плотность" in result.quarantine[0].reason.lower()


def test_parse_rejects_1c_code_in_volume_column(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "code.xlsx",
        [
            ("Блоки ФБС 9.3.6-Т", 350, 0.146, "00-00077237"),
            ("Блоки ФБС 24.6.6-Т", 1960, "00-00076381", "00-00077099"),
        ],
    )
    result = parse_product_weights_from_xlsx(xlsx, "fbs")
    assert len(result.records) == 1
    assert len(result.quarantine) == 1
    reason = result.quarantine[0].reason.lower()
    assert "код" in reason
    assert "объём" in reason or "объем" in reason


def test_parse_rejects_duplicate_marks_with_different_weights(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "dupes.xlsx",
        [
            ("Блоки ФБС 9.3.6-Т", 350, 0.146, "00-1"),
            ("Блоки ФБС 9.3.6-Т", 400, 0.146, "00-2"),
            ("Блоки ФБС 12.4.6-Т", 640, 0.265, "00-3"),
        ],
    )
    result = parse_product_weights_from_xlsx(xlsx, "fbs")
    assert len(result.records) == 1
    assert result.records[0].display_name.endswith("12.4.6-Т")
    assert len(result.quarantine) == 2
    assert all("дубл" in item.reason.lower() for item in result.quarantine)


def test_parse_skips_out_of_scope_lp_and_cross_blocks(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "scope.xlsx",
        [
            ("Кросс-Блок 50 (КБ-50)", 31.2, 0.013, "00-kb"),
            ("Лестничные площадки 2ЛП25.12-4-к", 1160, 0.464, "00-lp"),
            ("Блоки ФБС 12.4.6-Т", 640, 0.265, "00-fbs"),
        ],
    )
    result = parse_product_weights_from_xlsx(xlsx, "fbs")
    assert len(result.records) == 1
    assert result.records[0].product_type == "fbs"
    skipped_names = " ".join(item.display_name for item in result.skipped)
    assert "Кросс-Блок" in skipped_names
    assert "2ЛП25.12-4-к" in skipped_names


def test_resolve_exact_fbs_normalizes_spaces_and_t(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "fbs.xlsx",
        [("Блоки ФБС 24.6.6-Т", 1960, 0.815, "00-00077099")],
    )
    db_path = str(tmp_path / "plita.db")
    parsed = parse_product_weights_from_xlsx(xlsx, "fbs")
    inserted, updated = upsert_product_weights(db_path, parsed.records)
    assert (inserted, updated) == (1, 0)
    assert resolve_product_weight_kg("ФБС 24.6.6-Т", "fbs", db_path) == pytest.approx(
        1960.0
    )
    assert resolve_product_weight_kg("ФБС24.6.6-T", "fbs", db_path) == pytest.approx(
        1960.0
    )
    assert resolve_product_weight_kg("ФБС24", "fbs", db_path) is None


def test_resolve_ls_family_fallback_and_miss(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "ls.xlsx",
        [
            ("Лестничные ступени ЛС-14", 150, 0.06, "00-00003161"),
            ("Лестничные ступени ЛС-12", 133, 0.047, "00-00003506"),
        ],
    )
    db_path = str(tmp_path / "plita.db")
    parsed = parse_product_weights_from_xlsx(xlsx, "steps")
    upsert_product_weights(db_path, parsed.records)
    assert resolve_product_weight_kg("ЛС14-Б", "steps", db_path) == pytest.approx(150.0)
    assert resolve_product_weight_kg("ЛС-14", "steps", db_path) == pytest.approx(150.0)
    assert resolve_product_weight_kg("ЛС99", "steps", db_path) is None


def test_upsert_idempotent_updates_without_duplicates(tmp_path: Path) -> None:
    xlsx = _write_xlsx(
        tmp_path / "fbs.xlsx",
        [
            ("Блоки ФБС 9.3.6-Т", 350, 0.146, "00-1"),
            ("Блоки ФБС 12.4.6-Т", 640, 0.265, "00-2"),
        ],
    )
    db_path = str(tmp_path / "plita.db")
    records = parse_product_weights_from_xlsx(xlsx, "fbs").records
    assert (upsert_product_weights(db_path, records)) == (2, 0)
    assert (upsert_product_weights(db_path, records)) == (0, 2)
    with sqlite3.connect(db_path) as conn:
        total = conn.execute("SELECT COUNT(*) FROM product_weight_catalog").fetchone()[0]
    assert total == 2
    assert resolve_product_weight_kg("ФБС 9.3.6-Т", "fbs", db_path) == pytest.approx(
        350.0
    )
