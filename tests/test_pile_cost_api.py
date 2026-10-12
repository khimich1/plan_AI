"""PC-006: HTTP ручек, превью и записи цен цельных свай.

Превью не пишет pile_prices. PUT не принимает объём и килограммы арматуры.
Файл с якорем «цельн» по-прежнему 422 и ничего не пишет.
"""

from __future__ import annotations

import json
import sqlite3
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.services.pile_cost_service import MSG_NEED_SHA, MSG_STALE
from app.services.price_desk_service import MSG_PILE_CALCULATED, file_sha256
from core.pile_price_db import GRADE_CODES, init_pile_prices_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

PILE = "/api/v1/prices/pile-cost"
PREVIEW = "/api/v1/prices/pile-cost/preview"
APPLY = "/api/v1/prices/pile-cost/apply"
DESK_PREVIEW = "/api/v1/prices/preview"
DESK_APPLY = "/api/v1/prices/apply"
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_S30 = "С30.20-3"
_S40 = "С40.30-1"
_EXCLUDED = "С130.35-13А**"
_S30_B15 = {
    "materials_rub": 1431.4179,
    "expenses_rub": 968.3648,
    "cost": 2399.7827,
    "buyer_price": 3513.2818728,
}
_S30_BUYER = {
    "B15": 3513.2818728,
    "B20": 3538.3051464,
    "B22_5": 3559.5848256,
    "B25": 3580.8645048,
    "B30_granite": 3655.4166552,
}
_S40_B15_BUYER = 9033.108758784
_TOLERANCE = 0.01

_NORM_PATH = Path(__file__).resolve().parents[1] / "core" / "pile_cost_norm.json"
_NORM_ROWS = json.loads(_NORM_PATH.read_text(encoding="utf-8"))["rows"]
_NORM_ARTICLES = {
    str(line["article"]) for row in _NORM_ROWS for line in row["lines"]
}
_NORM_MARKS = {str(row["mark"]) for row in _NORM_ROWS}

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
        "id": 6,
        "username": "economist_user",
        "role": "economist",
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
def desk_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    pb_db = tmp_path / "pb.db"
    plita_db = tmp_path / "plita.db"
    pb_db.touch()
    plita_db.touch()
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PB_DB_PATH", str(pb_db))
    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return pb_db


@pytest.fixture()
def client(desk_env: Path) -> CsrfAwareTestClient:
    del desk_env
    return CsrfAwareTestClient(create_app())


def _economist() -> dict[str, str]:
    return session_cookie(6, "economist", "economist_user")


def _manager() -> dict[str, str]:
    return session_cookie(3, "manager", "manager_a")


def _pile_rows(db: Path) -> list[tuple] | None:
    conn = sqlite3.connect(db)
    try:
        exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'pile_prices'"
        ).fetchone()
        if exists is None:
            return None
        return conn.execute(
            """
            SELECT mark, concrete_grade, price, price_list_date, imported_at
            FROM pile_prices
            ORDER BY mark, concrete_grade
            """
        ).fetchall()
    finally:
        conn.close()


def _columns(db: Path, table: str) -> list[str]:
    conn = sqlite3.connect(db)
    try:
        exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table,),
        ).fetchone()
        if exists is None:
            return []
        return [str(row[1]) for row in conn.execute(f"PRAGMA table_info({table})")]
    finally:
        conn.close()


def _seed_pile(db: Path, rows: list[tuple]) -> None:
    init_pile_prices_schema(str(db))
    conn = sqlite3.connect(db)
    try:
        conn.executemany(
            """
            INSERT INTO pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def _inputs_payload(state: dict) -> dict:
    return {
        "prices": [dict(item) for item in state["inputs"]["prices"]],
        "recipes": [dict(item) for item in state["inputs"]["recipes"]],
        "overhead": dict(state["inputs"]["overhead"]),
    }


def _with_article(payload: dict, article: str, price: float) -> dict:
    prices = []
    replaced = False
    for item in payload["prices"]:
        copied = dict(item)
        if copied["article"] == article:
            copied["price"] = price
            replaced = True
        prices.append(copied)
    if not replaced:
        prices.append({"article": article, "price": price})
    return {**payload, "prices": prices}


def _with_norm_articles(payload: dict) -> dict:
    known = {item["article"] for item in payload["prices"]}
    prices = [dict(item) for item in payload["prices"]]
    for article in sorted(_NORM_ARTICLES - known):
        prices.append({"article": article, "price": 1.0})
    return {**payload, "prices": prices}


def _upload(client: CsrfAwareTestClient, url: str, path: Path, *, cookies, data: dict | None = None):
    return client.post(
        url,
        files={"file": (path.name, path.read_bytes(), XLSX_MEDIA)},
        data=data,
        cookies=cookies,
    )


def _write_fbs_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _write_plates_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", "Цены по прайсу, который прислал Никита"],
        [1, "ПБ 17-12-6", 6049],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _grade(body: dict, code: str) -> dict:
    found = [item for item in body["grades"] if item["concrete_grade"] == code]
    assert len(found) == 1
    return found[0]


def test_get_returns_handles_and_brief_mark_list(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    response = client.get(PILE, cookies=_economist())
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["inputs_sha256"]) == 64
    cement = next(item for item in body["inputs"]["prices"] if item["article"] == "cement")
    assert cement["price"] == 7.54
    assert [item["concrete_grade"] for item in body["inputs"]["recipes"]] == list(GRADE_CODES)
    assert body["breakdown"] is None
    assert len(body["marks"]) == len(_NORM_MARKS)
    assert len(body["marks"]) < len(_NORM_ROWS)
    assert _EXCLUDED not in {item["mark"] for item in body["marks"]}
    summary = next(item for item in body["marks"] if item["mark"] == _S30)
    assert set(summary) == {"mark", "section_cm", "grades"}
    assert summary["section_cm"] == 20
    assert [item["concrete_grade"] for item in summary["grades"]] == list(GRADE_CODES)
    assert _grade(summary, "B15")["buyer_price"] == pytest.approx(_S30_B15["buyer_price"], abs=_TOLERANCE)
    unpriced = next(item for item in body["marks"] if item["mark"] == "С30.15")
    assert _grade(unpriced, "B15")["buyer_price"] is None
    assert _pile_rows(desk_env) is None


def test_get_mark_breakdown_matches_control_values(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    response = client.get(PILE, params={"mark": _S30}, cookies=_economist())
    assert response.status_code == 200, response.text
    card = response.json()["breakdown"]
    assert card["mark"] == _S30
    assert card["section_cm"] == 20
    assert card["volume_m3"] == pytest.approx(0.13)
    b15 = _grade(card, "B15")
    assert b15["surcharge_applied"] is False
    assert b15["materials_rub"] == pytest.approx(_S30_B15["materials_rub"], abs=_TOLERANCE)
    assert b15["expenses_rub"] == pytest.approx(_S30_B15["expenses_rub"], abs=_TOLERANCE)
    assert b15["cost"] == pytest.approx(_S30_B15["cost"], abs=_TOLERANCE)
    assert b15["buyer_price"] == pytest.approx(_S30_B15["buyer_price"], abs=_TOLERANCE)
    kg = {line["article"]: line["kg"] for line in b15["lines"]}
    assert kg == {"wire_5": 3.1, "ai_10": 3.15, "a500c_10": 8.0}
    wide = client.get(PILE, params={"mark": _S40}, cookies=_economist())
    assert wide.status_code == 200, wide.text
    wide_b15 = _grade(wide.json()["breakdown"], "B15")
    assert wide_b15["surcharge_applied"] is True
    assert wide_b15["buyer_price"] == pytest.approx(_S40_B15_BUYER, abs=_TOLERANCE)
    assert _pile_rows(desk_env) is None


def test_get_unknown_mark_404(client: CsrfAwareTestClient) -> None:
    response = client.get(PILE, params={"mark": "нет такой"}, cookies=_economist())
    assert response.status_code == 404
    assert response.json()["detail"] == "марка не найдена"


def test_get_unpriced_mark_422_writes_nothing(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    response = client.get(PILE, params={"mark": "С30.15"}, cookies=_economist())
    assert response.status_code == 422
    assert "wire_4_8" in response.json()["detail"]
    assert _pile_rows(desk_env) is None
    assert "volume_m3" not in _columns(desk_env, "pile_cost_recipe")
    assert "kg" not in _columns(desk_env, "pile_cost_article_price")


def test_put_saves_cement_and_does_not_touch_pile_prices(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(desk_env, [(_EXCLUDED, "B15", 60980.85, "2026-10-01", "2026-10-01T08:00:00")])
    before = _pile_rows(desk_env)
    state = client.get(PILE, cookies=_economist())
    assert state.status_code == 200
    old_sha = state.json()["inputs_sha256"]
    payload = _with_article(_inputs_payload(state.json()), "cement", 8.1)
    saved = client.put(PILE, json=payload, cookies=_economist())
    assert saved.status_code == 200, saved.text
    body = saved.json()
    assert body["inputs_sha256"] != old_sha
    assert len(body["inputs_sha256"]) == 64
    cement = next(item for item in body["inputs"]["prices"] if item["article"] == "cement")
    assert cement["price"] == 8.1
    assert "volume_m3" not in body["inputs"]
    assert _pile_rows(desk_env) == before
    again = client.get(PILE, cookies=_economist())
    assert again.status_code == 200
    assert again.json()["inputs_sha256"] == body["inputs_sha256"]


def test_put_rejects_non_positive_price_and_negative_recipe(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    state = client.get(PILE, cookies=_economist()).json()
    zero = _with_article(_inputs_payload(state), "cement", 0)
    rejected = client.put(PILE, json=zero, cookies=_economist())
    assert rejected.status_code == 422
    assert "больше нуля" in rejected.json()["detail"]
    negative = _inputs_payload(state)
    negative["recipes"][0]["cement_kg_m3"] = -1
    rejected_recipe = client.put(PILE, json=negative, cookies=_economist())
    assert rejected_recipe.status_code == 422
    assert "отрицательн" in rejected_recipe.json()["detail"]
    kept = client.get(PILE, cookies=_economist())
    cement = next(item for item in kept.json()["inputs"]["prices"] if item["article"] == "cement")
    assert cement["price"] == 7.54
    assert _pile_rows(desk_env) is None


def test_put_rejects_volume_and_rebar_kg(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    state = client.get(PILE, cookies=_economist()).json()
    payload = _inputs_payload(state)
    payload["volume_m3"] = 0.13
    rejected = client.put(PILE, json=payload, cookies=_economist())
    assert rejected.status_code == 422
    priced = _inputs_payload(state)
    priced["prices"][0]["kg"] = 8
    rejected_kg = client.put(PILE, json=priced, cookies=_economist())
    assert rejected_kg.status_code == 422
    assert _pile_rows(desk_env) is None
    assert "volume_m3" not in _columns(desk_env, "pile_cost_recipe")
    assert "kg" not in _columns(desk_env, "pile_cost_article_price")
    cement = next(
        item
        for item in client.get(PILE, cookies=_economist()).json()["inputs"]["prices"]
        if item["article"] == "cement"
    )
    assert cement["price"] == 7.54


def test_preview_does_not_update_pile_prices(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(desk_env, [(_EXCLUDED, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00")])
    state = client.get(PILE, cookies=_economist())
    saved = client.put(
        PILE,
        json=_with_norm_articles(_inputs_payload(state.json())),
        cookies=_economist(),
    )
    assert saved.status_code == 200, saved.text
    before = _pile_rows(desk_env)
    preview = client.post(PREVIEW, cookies=_economist())
    assert preview.status_code == 200, preview.text
    body = preview.json()
    assert body["inputs_sha256"] == saved.json()["inputs_sha256"]
    assert body["missing"] == 1
    assert _pile_rows(desk_env) == before


def test_preview_without_metal_prices_422_writes_nothing(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    client.get(PILE, cookies=_economist())
    _seed_pile(desk_env, [(_EXCLUDED, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00")])
    before = _pile_rows(desk_env)
    preview = client.post(PREVIEW, cookies=_economist())
    assert preview.status_code == 422
    assert "нет цены статьи" in preview.json()["detail"]
    assert _pile_rows(desk_env) == before


def test_apply_without_sha256_409_writes_nothing(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(desk_env, [(_EXCLUDED, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00")])
    before = _pile_rows(desk_env)
    missing = client.post(APPLY, cookies=_economist())
    assert missing.status_code == 409
    assert MSG_NEED_SHA in missing.json()["detail"]
    empty = client.post(APPLY, json={}, cookies=_economist())
    assert empty.status_code == 409
    assert MSG_NEED_SHA in empty.json()["detail"]
    blank = client.post(APPLY, json={"inputs_sha256": "  "}, cookies=_economist())
    assert blank.status_code == 409
    assert _pile_rows(desk_env) == before


def test_apply_wrong_sha256_409_keeps_prices(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(desk_env, [(_S30, "B15", 1.0, "2026-10-06", "2026-10-06T12:00:00")])
    state = client.get(PILE, cookies=_economist())
    assert client.put(
        PILE,
        json=_with_norm_articles(_inputs_payload(state.json())),
        cookies=_economist(),
    ).status_code == 200
    before = _pile_rows(desk_env)
    response = client.post(APPLY, json={"inputs_sha256": "ab" * 32}, cookies=_economist())
    assert response.status_code == 409
    assert MSG_STALE in response.json()["detail"]
    assert _pile_rows(desk_env) == before


def test_apply_writes_five_s30_prices_and_keeps_missing(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(
        desk_env,
        [
            (_EXCLUDED, "B15", 60980.85, "2026-10-01", "2026-10-01T08:00:00"),
            (_EXCLUDED, "B20", 61000.0, "2026-10-01", "2026-10-01T08:00:00"),
        ],
    )
    state = client.get(PILE, cookies=_economist())
    saved = client.put(
        PILE,
        json=_with_norm_articles(_inputs_payload(state.json())),
        cookies=_economist(),
    )
    assert saved.status_code == 200, saved.text
    digest = saved.json()["inputs_sha256"]
    preview = client.post(PREVIEW, cookies=_economist())
    assert preview.status_code == 200, preview.text
    assert preview.json()["inputs_sha256"] == digest
    before_count = len(_pile_rows(desk_env) or [])
    applied = client.post(APPLY, json={"inputs_sha256": digest.upper()}, cookies=_economist())
    assert applied.status_code == 200, applied.text
    assert applied.json()["inputs_sha256"] == digest
    assert applied.json()["missing"] == 2
    rows = _pile_rows(desk_env)
    assert rows is not None
    assert len(rows) > before_count
    by_key = {
        (mark, grade): (price, list_date, imported_at)
        for mark, grade, price, list_date, imported_at in rows
    }
    for grade, expected in _S30_BUYER.items():
        price, list_date, imported_at = by_key[(_S30, grade)]
        assert price == pytest.approx(expected, abs=_TOLERANCE)
        assert price != round(price)
        assert list_date == date.today().isoformat()
        assert imported_at
    assert by_key[(_EXCLUDED, "B15")] == (60980.85, "2026-10-01", "2026-10-01T08:00:00")
    assert by_key[(_EXCLUDED, "B20")] == (61000.0, "2026-10-01", "2026-10-01T08:00:00")
    assert "volume_m3" not in _columns(desk_env, "pile_prices")
    assert "kg" not in _columns(desk_env, "pile_prices")


def test_apply_rejects_volume_and_does_not_write(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    _seed_pile(desk_env, [(_EXCLUDED, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00")])
    before = _pile_rows(desk_env)
    state = client.get(PILE, cookies=_economist())
    digest = state.json()["inputs_sha256"]
    response = client.post(
        APPLY,
        json={"inputs_sha256": digest, "volume_m3": 0.13, "kg": 8},
        cookies=_economist(),
    )
    assert response.status_code == 422
    assert _pile_rows(desk_env) == before


def test_manager_forbidden_on_pile_cost(
    client: CsrfAwareTestClient, desk_env: Path
) -> None:
    payload = {
        "prices": [{"article": "cement", "price": 7.54}],
        "recipes": [
            {
                "concrete_grade": "B15",
                "cement_kg_m3": 320,
                "sand_kg_m3": 990,
                "crushed_kg_m3": 950,
            }
        ],
        "overhead": {"general": 1, "payroll": 1, "forming": 1, "other": 1},
    }
    assert client.get(PILE, cookies=_manager()).status_code == 403
    assert client.put(PILE, json=payload, cookies=_manager()).status_code == 403
    assert client.post(PREVIEW, cookies=_manager()).status_code == 403
    assert client.post(APPLY, json={"inputs_sha256": "ab" * 32}, cookies=_manager()).status_code == 403
    assert _pile_rows(desk_env) is None


def test_solid_pile_file_422_writes_nothing(
    client: CsrfAwareTestClient, desk_env: Path, tmp_path: Path
) -> None:
    xlsx = tmp_path / "Прайс на цельные сваи от 07.09.2026.xlsx"
    _write_fbs_xlsx(xlsx)
    data = xlsx.read_bytes()
    before = _pile_rows(desk_env)
    preview = _upload(client, DESK_PREVIEW, xlsx, cookies=_economist())
    assert preview.status_code == 422
    assert MSG_PILE_CALCULATED in preview.json()["detail"]
    applied = _upload(
        client,
        DESK_APPLY,
        xlsx,
        cookies=_economist(),
        data={"file_sha256": file_sha256(data)},
    )
    assert applied.status_code == 422
    assert MSG_PILE_CALCULATED in applied.json()["detail"]
    assert _pile_rows(desk_env) == before


def test_plates_and_fbs_preview_still_200(
    client: CsrfAwareTestClient, desk_env: Path, tmp_path: Path
) -> None:
    plates = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_plates_xlsx(plates)
    fbs = tmp_path / "Прайс на ФБС от 07.09.2026.xlsx"
    _write_fbs_xlsx(fbs)
    plates_response = _upload(client, DESK_PREVIEW, plates, cookies=_economist())
    assert plates_response.status_code == 200, plates_response.text
    assert plates_response.json()["product_kind"] == "plates"
    fbs_response = _upload(client, DESK_PREVIEW, fbs, cookies=_economist())
    assert fbs_response.status_code == 200, fbs_response.text
    assert fbs_response.json()["product_kind"] == "fbs"
    assert _pile_rows(desk_env) is None
