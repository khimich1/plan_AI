"""TDD RED for Task T3: импорт транзакций ГСМ (парсер + сервис + endpoint).

Оригиналы ``ГСМ/**`` только на чтение. Синтетические .xls — в
``tests/fixtures/gsm/`` (BIFF2 + копия маленького реального файла).

Ожидаемый API для worker — см. блок в конце файла.
"""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

# ---------------------------------------------------------------------------
# Expected public API (worker) — ImportError until GREEN
# ---------------------------------------------------------------------------
from core.gsm.transactions import (  # noqa: E402
    ParsedTxFile,
    ParsedTxRow,
    classify_service,
    parse_transactions_xls,
)
from app.services.gsm_transaction_service import GsmTransactionService  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIXTURES = PROJECT_ROOT / "tests" / "fixtures" / "gsm"
REAL_TX_DIR = PROJECT_ROOT / "ГСМ" / "транзакции"

SAMPLE_OK = FIXTURES / "sample_ok.xls"
SAMPLE_MISMATCH = FIXTURES / "sample_mismatch.xls"
SAMPLE_TINY = FIXTURES / "sample_tiny.xls"
SAMPLE_UNKNOWN = FIXTURES / "sample_unknown_card.xls"
REAL_SMALL = FIXTURES / "real_card_4268_small.xls"

IMPORT_API = "/api/v1/gsm/transactions/import"
XLS_MEDIA = "application/vnd.ms-excel"

TEST_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


# =============================================================================
# Helpers / fixtures
# =============================================================================


def _fresh_db(tmp_path: Path, name: str = "gsm_tx.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _count(db_path: str, table: str, where: str = "1=1", params: tuple = ()) -> int:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            f"SELECT COUNT(*) FROM {table} WHERE {where}", params
        ).fetchone()
    return int(row[0])


def _seed_known_card(
    repo: GsmRepository,
    *,
    card_number: str = "3005454268",
) -> tuple[int, int]:
    driver_id = repo.create_driver(
        full_name="Тестов Тест",
        license_number="00 00 000000",
    )
    vehicle_id = repo.create_vehicle(
        name="Test Car",
        plate_number="А000АА00",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number=card_number,
        vehicle_id=vehicle_id,
        assigned_at="2025-01-01",
    )
    return vehicle_id, card_id


def _file_tuple(path: Path) -> tuple[str, bytes, str]:
    return (path.name, path.read_bytes(), XLS_MEDIA)


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    return _fresh_db(tmp_path)


@pytest.fixture()
def repo(db_path: str) -> GsmRepository:
    return GsmRepository(db_path=db_path)


@pytest.fixture()
def service(repo: GsmRepository) -> GsmTransactionService:
    return GsmTransactionService(repo=repo)


@pytest.fixture()
def api_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> CsrfAwareTestClient:
    db = tmp_path / "plita.db"
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db))
    monkeypatch.setenv("PB_DB_PATH", str(db))
    get_settings.cache_clear()
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(str(db))
    patch_auth_users(monkeypatch, TEST_USERS)
    return CsrfAwareTestClient(create_app())


# =============================================================================
# 1. Pure parser — core/gsm/transactions.py
# =============================================================================


@pytest.mark.parametrize(
    ("service", "expected"),
    [
        ("АИ-95", "fuel"),
        ("АИ-95 Фирменный", "fuel"),
        ("аи-92", "fuel"),
        ("Мойка", "wash"),
        ("мойка", "wash"),
        ("Кофе", "other"),
        ("Сопутствующие товары", "other"),
    ],
)
def test_classify_service_fuel_wash_other(service: str, expected: str) -> None:
    assert classify_service(service) == expected


def test_parse_sample_ok_header_skip_totals_classify_and_reconcile() -> None:
    """Шапка «Дата трн.», строки «Итоги:» отброшены, типы fuel/wash/other, сверка ок."""
    assert SAMPLE_OK.is_file()
    parsed = parse_transactions_xls(SAMPLE_OK)

    assert parsed.filename == SAMPLE_OK.name
    assert len(parsed.rows) == 4
    assert [r.service_type for r in parsed.rows] == ["fuel", "wash", "fuel", "other"]

    fuel0 = parsed.rows[0]
    assert fuel0.card_number == "3005454268"
    assert fuel0.fuel_grade == "АИ-95"
    assert fuel0.qty_liters == pytest.approx(40.0)
    assert fuel0.amount == pytest.approx(2000.0)
    assert isinstance(fuel0.ts, datetime)

    wash = parsed.rows[1]
    assert wash.service_type == "wash"
    assert wash.fuel_grade is None
    assert wash.qty_liters is None
    assert wash.amount == pytest.approx(400.0)

    other = parsed.rows[3]
    assert other.service_type == "other"
    assert other.fuel_grade is None

    # Суммы по строкам vs футер «Итоги:» — совпадают → без warning
    assert parsed.sum_liters == pytest.approx(66.0)
    assert parsed.sum_amount == pytest.approx(4000.0)
    assert parsed.footer_liters == pytest.approx(66.0)
    assert parsed.footer_amount == pytest.approx(4000.0)
    assert list(parsed.warnings) == []


def test_parse_mismatch_footer_yields_warning_not_failure() -> None:
    """Расхождение с «Итоги:» → warnings, парсер не падает и строки возвращает."""
    parsed = parse_transactions_xls(SAMPLE_MISMATCH)

    assert len(parsed.rows) == 4
    assert parsed.sum_liters == pytest.approx(66.0)
    assert parsed.sum_amount == pytest.approx(4000.0)
    assert parsed.footer_liters == pytest.approx(999.0)
    assert parsed.footer_amount == pytest.approx(1.0)
    assert len(parsed.warnings) >= 1
    blob = " ".join(parsed.warnings).lower()
    assert "итог" in blob or "≠" in " ".join(parsed.warnings) or "!=" in blob


def test_parse_real_small_fixture_skips_totals_row() -> None:
    """Копия реального файла: 12 транзакций, шапка на строке с «Дата трн.»."""
    assert REAL_SMALL.is_file()
    parsed = parse_transactions_xls(REAL_SMALL)

    assert len(parsed.rows) == 12
    assert all(r.service_type == "fuel" for r in parsed.rows)
    assert all(r.card_number == "3005454268" for r in parsed.rows)
    assert parsed.footer_liters is not None
    assert parsed.footer_amount is not None
    assert abs(parsed.sum_liters - parsed.footer_liters) <= 0.01
    assert abs(parsed.sum_amount - parsed.footer_amount) <= 0.01
    assert list(parsed.warnings) == []


# =============================================================================
# 2. Dedup — повторный импорт → 0 новых строк (UNIQUE)
# =============================================================================


def test_reimport_same_file_inserts_zero_new_rows(
    service: GsmTransactionService, repo: GsmRepository, db_path: str
) -> None:
    _seed_known_card(repo, card_number="3005454999")

    first = service.import_files(
        [(SAMPLE_TINY.name, SAMPLE_TINY.read_bytes())],
        uploaded_by="accountant_user",
    )
    assert first.rows_inserted == 1
    assert _count(db_path, "gsm_transaction") == 1

    second = service.import_files(
        [(SAMPLE_TINY.name, SAMPLE_TINY.read_bytes())],
        uploaded_by="accountant_user",
    )
    assert second.rows_inserted == 0
    assert second.rows_duplicate == 1
    assert second.files[0].rows_inserted == 0
    assert second.files[0].rows_duplicate == 1
    assert _count(db_path, "gsm_transaction") == 1


# =============================================================================
# 3. Unknown card → accepted, unmatched_card, card auto-created (vehicle_id NULL)
# =============================================================================


def test_unknown_card_accepted_with_unmatched_flag_and_null_vehicle(
    service: GsmTransactionService, db_path: str
) -> None:
    """Неизвестная карта не стопает импорт: карта создаётся, vehicle_id может быть NULL.

    Worker: при необходимости ослабить ``gsm_fuel_card.vehicle_id`` до NULLABLE.
    """
    report = service.import_files(
        [(SAMPLE_UNKNOWN.name, SAMPLE_UNKNOWN.read_bytes())],
        uploaded_by="accountant_user",
    )

    assert report.rows_inserted == 1
    assert report.files[0].rows_inserted == 1
    assert "9990001111" in report.files[0].unmatched_cards

    assert _count(db_path, "gsm_fuel_card", "card_number = ?", ("9990001111",)) == 1
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT vehicle_id, archived_at FROM gsm_fuel_card WHERE card_number = ?",
            ("9990001111",),
        ).fetchone()
    assert row is not None
    assert row[0] is None  # vehicle_id NULL until linked later
    assert row[1] is None  # not archived

    assert _count(db_path, "gsm_transaction") == 1


# =============================================================================
# 3b. Wash qty safety net — numeric payload still persists as NULL
# =============================================================================


def test_import_wash_with_numeric_qty_persists_null_liters(
    service: GsmTransactionService,
    repo: GsmRepository,
    db_path: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Safety net: wash with qty=1.0 from parser still stores qty_liters NULL."""
    card_number = "3005454268"
    _seed_known_card(repo, card_number=card_number)

    wash_row = ParsedTxRow(
        ts=datetime(2026, 8, 4, 10, 0, 0),
        card_number=card_number,
        service="Мойка",
        service_type="wash",
        fuel_grade=None,
        qty_liters=1.0,
        amount=400.0,
        unit="шт",
        brand="Роснефть",
        city="Ярославль",
        raw_address="ул. Тестовая, 1",
    )
    parsed = ParsedTxFile(
        filename="wash_with_qty.xls",
        rows=(wash_row,),
        sum_liters=1.0,
        sum_amount=400.0,
        footer_liters=1.0,
        footer_amount=400.0,
        warnings=(),
    )
    monkeypatch.setattr(
        "app.services.gsm_transaction_service.parse_transactions_content",
        lambda _content, filename: parsed,
    )

    report = service.import_files(
        [("wash_with_qty.xls", b"dummy-xls")],
        uploaded_by="accountant_user",
    )

    assert report.rows_inserted == 1
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT service_type, qty_liters FROM gsm_transaction"
        ).fetchone()
    assert row is not None
    assert row[0] == "wash"
    assert row[1] is None


# =============================================================================
# 4. Endpoint POST /api/v1/gsm/transactions/import (REQUIRE_ACCOUNTING)
# =============================================================================


def test_import_endpoint_multi_file_report_under_accounting(
    api_client: CsrfAwareTestClient, tmp_path: Path
) -> None:
    """Мульти-файл: ответ по каждому файлу (rows, liters, amount vs итог)."""
    # Seed known card into the API DB so sample_ok does not flood unmatched.
    db = str(tmp_path / "plita.db")
    repo = GsmRepository(db_path=db)
    _seed_known_card(repo, card_number="3005454268")

    response = api_client.post(
        IMPORT_API,
        files=[
            ("files", _file_tuple(SAMPLE_OK)),
            ("files", _file_tuple(SAMPLE_MISMATCH)),
        ],
        cookies=session_cookie(5, "accountant", "accountant_user"),
    )
    assert response.status_code == 200, response.text
    body = response.json()

    assert "files" in body
    assert len(body["files"]) == 2
    by_name = {f["filename"]: f for f in body["files"]}

    ok = by_name[SAMPLE_OK.name]
    assert ok["rows_total"] == 4
    assert ok["rows_inserted"] == 4
    assert ok["sum_liters"] == pytest.approx(66.0)
    assert ok["sum_amount"] == pytest.approx(4000.0)
    assert ok["footer_liters"] == pytest.approx(66.0)
    assert ok["footer_amount"] == pytest.approx(4000.0)
    assert ok.get("warnings") in ([], None) or ok["warnings"] == []

    bad = by_name[SAMPLE_MISMATCH.name]
    assert bad["rows_total"] == 4
    # Дедуп против уже вставленных из sample_ok (тот же UNIQUE-ключ) → 0 новых
    assert bad["rows_inserted"] == 0
    assert bad["rows_duplicate"] == 4
    assert bad["footer_liters"] == pytest.approx(999.0)
    assert bad["footer_amount"] == pytest.approx(1.0)
    assert len(bad["warnings"]) >= 1

    assert body["rows_inserted"] == 4
    assert body["rows_duplicate"] == 4


def test_import_endpoint_forbids_manager(api_client: CsrfAwareTestClient) -> None:
    response = api_client.post(
        IMPORT_API,
        files=[("files", _file_tuple(SAMPLE_TINY))],
        cookies=session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 403


def test_import_endpoint_allows_admin(api_client: CsrfAwareTestClient) -> None:
    response = api_client.post(
        IMPORT_API,
        files=[("files", _file_tuple(SAMPLE_TINY))],
        cookies=session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["rows_inserted"] == 1
    assert body["files"][0]["unmatched_cards"] == ["3005454999"]


# =============================================================================
# 5. Acceptance smoke — 9 files → 509 txs (integration / slow)
# =============================================================================


requires_real_tx = pytest.mark.skipif(
    not REAL_TX_DIR.is_dir() or len(list(REAL_TX_DIR.glob("*.xls"))) < 9,
    reason="реальный каталог ГСМ/транзакции отсутствует или неполный",
)


@requires_real_tx
@pytest.mark.integration
def test_import_nine_real_files_yields_509_transactions(
    tmp_path: Path,
) -> None:
    """SC-1: 9 файлов → 509 транзакций; повторный импорт без дублей."""
    db_path = _fresh_db(tmp_path, "gsm_tx_real.db")
    repo = GsmRepository(db_path=db_path)
    service = GsmTransactionService(repo=repo)

    paths = sorted(REAL_TX_DIR.glob("*.xls"))
    assert len(paths) == 9

    payload = [(p.name, p.read_bytes()) for p in paths]
    report = service.import_files(payload, uploaded_by="accountant_user")

    assert report.rows_inserted == 509
    assert _count(db_path, "gsm_transaction") == 509
    assert len(report.files) == 9
    # Сверка по файлам: нет unexpected fail — mismatch только warning
    for f in report.files:
        assert f.rows_total >= 1
        if f.footer_liters is not None:
            # либо сошлось, либо есть warning
            matched = abs(f.sum_liters - f.footer_liters) <= 0.01
            assert matched or len(f.warnings) >= 1

    again = service.import_files(payload, uploaded_by="accountant_user")
    assert again.rows_inserted == 0
    assert again.rows_duplicate == 509
    assert _count(db_path, "gsm_transaction") == 509


# =============================================================================
# Expected API for worker (Task T3)
# =============================================================================
#
# core/gsm/transactions.py  (pure; no app.* / no DB)
# -----------------------------------------------
# classify_service(service: str) -> Literal["fuel", "wash", "other"]
#     "Мойка" (case-insensitive) → wash
#     услуга, начинающаяся с "АИ-" (case-insensitive) → fuel
#     иначе → other
#
# @dataclass(frozen=True, slots=True)
# class ParsedTxRow:
#     ts: datetime
#     card_number: str          # нормализованный, без ".0"
#     service: str              # сырой текст «Услуга»
#     service_type: Literal["fuel","wash","other"]
#     fuel_grade: str | None    # = service для fuel, иначе None
#     qty_liters: float | None
#     amount: float
#     unit: str
#     brand: str
#     city: str
#     raw_address: str          # «Адрес ТО»
#
# @dataclass(frozen=True, slots=True)
# class ParsedTxFile:
#     filename: str
#     rows: tuple[ParsedTxRow, ...]
#     sum_liters: float         # Σ qty по строкам (None→0)
#     sum_amount: float
#     footer_liters: float | None
#     footer_amount: float | None
#     warnings: tuple[str, ...] # расхождение с «Итоги:» → warning, НЕ exception
#
# parse_transactions_xls(path: Path | str) -> ParsedTxFile
#     Шапка: первая строка (в первых 5), где col0 == "Дата трн."
#     Строки, начинающиеся с "Итоги" — футер, не транзакции
#     Сверка: abs(sum - footer) > 0.01 → warning в warnings
#
# app/services/gsm_transaction_service.py
# ---------------------------------------
# class GsmTransactionService:
#     def __init__(self, *, repo: GsmRepository) -> None: ...
#     def import_files(
#         self,
#         files: Sequence[tuple[str, bytes]],  # (filename, content)
#         *,
#         uploaded_by: str | None = None,
#     ) -> TransactionImportReport: ...
#
# TransactionImportReport:
#     files: list[FileImportReport]
#     rows_inserted: int
#     rows_duplicate: int
#
# FileImportReport:
#     filename: str
#     rows_total: int
#     rows_inserted: int
#     rows_duplicate: int
#     sum_liters: float
#     sum_amount: float
#     footer_liters: float | None
#     footer_amount: float | None
#     warnings: list[str]
#     unmatched_cards: list[str]   # номера карт, авто-созданных без vehicle
#
# Поведение:
#   - неизвестная карта → CREATE gsm_fuel_card(card_number, vehicle_id=NULL, ...)
#     + номер в unmatched_cards; транзакция всё равно INSERT
#   - UNIQUE(card_id, ts, qty_liters, amount) → duplicate skip (не 500)
#   - схема: vehicle_id на gsm_fuel_card допускается NULL (amend IF NOT EXISTS / migrate)
#
# app/api/v1/endpoints/gsm.py (+ router include)
# ---------------------------------------------
# POST /api/v1/gsm/transactions/import
#     Depends(REQUIRE_ACCOUNTING)
#     multipart field name: "files" (несколько UploadFile)
#     response_model ≈ TransactionImportReport (JSON как выше)
#
# app/schemas/gsm.py — Pydantic v2, ConfigDict(extra="forbid") на request при наличии
# app/dependencies/services.py — get_gsm_transaction_service()
