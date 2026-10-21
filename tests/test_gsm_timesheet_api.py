"""POST/GET /api/v1/gsm/timesheets — Т-13 import and attendance grid."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.gsm_t13 import build_t13_xlsx
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

IMPORT = "/api/v1/gsm/timesheets/import"
GRID = "/api/v1/gsm/timesheets"

DRIVERS = (
    ("Кулигин Никита Валерьевич", "143", "44 21 846315"),
    ("Лоншакова Мария", "470", "44 21 000470"),
    ("Скрябин Иван Петрович", "29", "44 21 000029"),
    ("Полякова Анна", "1", "44 21 000001"),
    ("Шишов Алексей", "438", "44 21 000438"),
)

TEST_USERS = [
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


def _auth(username: str = "accountant_user") -> dict[str, str]:
    if username == "manager_a":
        return session_cookie(3, "manager", username)
    return session_cookie(5, "accountant", username)


def _repo() -> GsmRepository:
    return GsmRepository(db_path=str(get_settings().plita_db_path))


def _seed_drivers() -> None:
    repo = _repo()
    for full_name, personnel, license_number in DRIVERS:
        repo.create_driver(
            full_name=full_name,
            license_number=license_number,
            personnel_number=personnel,
        )


def _july_people() -> list[dict]:
    skryabin_ot = {day: "ОТ" for day in range(1, 10)}
    people = []
    for full_name, personnel, _license in DRIVERS:
        raw = f"СЭ00-{int(personnel):05d}"
        codes = skryabin_ot if personnel == "29" else {}
        people.append(
            {
                "full_name": full_name,
                "personnel": raw,
                "default_code": "Я",
                "codes": codes,
            }
        )
    people.append(
        {
            "full_name": "Офисная Сотрудница",
            "personnel": "СЭ00-00999",
            "default_code": "Я",
        }
    )
    people.append(
        {
            "full_name": "Без Табельного",
            "personnel": "",
            "default_code": "Я",
        }
    )
    return people


def _upload(client: CsrfAwareTestClient, filename: str, content: bytes, month: str | None = None):
    data = {"month": month} if month else None
    return client.post(
        IMPORT,
        files={
            "file": (
                filename,
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
        data=data,
        cookies=_auth(),
    )


def test_july_import_matches_five_drivers_and_stores_days(
    api_client: CsrfAwareTestClient,
) -> None:
    _seed_drivers()
    content = build_t13_xlsx(_july_people())
    resp = _upload(api_client, "табель от 31.07.2026 (1).xlsx", content)
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["period"] == "2026-07"
    assert body["matched"] == 5
    assert body["rows_total"] == 7
    unmatched_names = {row["full_name"] for row in body["unmatched"]}
    assert "Офисная Сотрудница" in unmatched_names
    assert "Без Табельного" in unmatched_names

    grid = api_client.get(f"{GRID}?period=2026-07", cookies=_auth())
    assert grid.status_code == 200, grid.text
    cells = grid.json()["cells"]
    skryabin = _repo().list_drivers()
    skryabin_id = next(row["id"] for row in skryabin if row["personnel_number"] == "29")
    ot_days = [
        cell["date"]
        for cell in cells
        if cell["driver_id"] == skryabin_id and cell["code"] == "ОТ"
    ]
    assert ot_days[0] == "2026-07-01"
    assert "2026-07-09" in ot_days
    assert "2026-07-10" not in ot_days
    assert len(grid.json()["batches"]) == 1


def test_import_without_period_in_name_requires_month(
    api_client: CsrfAwareTestClient,
) -> None:
    _seed_drivers()
    content = build_t13_xlsx(_july_people())
    missing = _upload(api_client, "табель.xlsx", content)
    assert missing.status_code == 400
    assert missing.json()["detail"]["code"] == "gsm_timesheet_month_required"

    chosen = _upload(api_client, "табель.xlsx", content, month="2026-07")
    assert chosen.status_code == 200, chosen.text
    assert chosen.json()["period"] == "2026-07"


def test_import_forbidden_for_manager(api_client: CsrfAwareTestClient) -> None:
    content = build_t13_xlsx(_july_people())
    resp = api_client.post(
        IMPORT,
        files={"file": ("табель от 31.07.2026.xlsx", content, "application/octet-stream")},
        cookies=_auth("manager_a"),
    )
    assert resp.status_code == 403


def test_repeat_import_replaces_days_and_keeps_both_batches(
    api_client: CsrfAwareTestClient,
) -> None:
    _seed_drivers()
    first_people = _july_people()
    first = _upload(api_client, "табель от 31.07.2026.xlsx", build_t13_xlsx(first_people))
    assert first.status_code == 200, first.text

    revised = []
    for person in first_people:
        item = dict(person)
        if person["personnel"] == "СЭ00-00143":
            item["codes"] = {12: "В"}
        revised.append(item)
    second = _upload(api_client, "табель от 31.07.2026.xlsx", build_t13_xlsx(revised))
    assert second.status_code == 200, second.text

    grid = api_client.get(f"{GRID}?period=2026-07", cookies=_auth())
    assert grid.status_code == 200
    kuligin_id = next(
        row["id"] for row in _repo().list_drivers() if row["personnel_number"] == "143"
    )
    july_12 = next(
        cell
        for cell in grid.json()["cells"]
        if cell["driver_id"] == kuligin_id and cell["date"] == "2026-07-12"
    )
    assert july_12["code"] == "В"
    assert len(grid.json()["batches"]) == 2
    assert _repo().get_timesheet_code(kuligin_id, date(2026, 7, 12)) == "В"
