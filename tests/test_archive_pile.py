"""Archive API: product_type filter and pile KP details (AC-10, AC-14, AC-15)."""

from __future__ import annotations

import asyncio
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.services.archive_service import ArchiveService
from core.kp_persistence_service import KpPersistenceService
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"

ADMIN_USER = {
    "id": 1,
    "username": "admin",
    "role": "admin",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
}

PLATE_ORDER = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
    }
]

PILE_ORDER = [
    {
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 44634.03,
    }
]


def _admin_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": 1, "username": "admin", "role": "admin"},
            ttl_seconds=300,
        )
    }


@pytest.fixture()
def archive_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[TestClient, str]:
    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, [ADMIN_USER])
    return TestClient(create_app()), db_path


def _save_plate_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        PLATE_ORDER,
        customer_name="Plate client",
        status="в архиве",
        db_path=db_path,
    )


def _save_pile_kp(db_path: str) -> int:
    return KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        PILE_ORDER,
        customer_name="Pile client",
        status="в архиве",
        product_type="piles",
        db_path=db_path,
    )


def test_archive_list_includes_product_type(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    _save_plate_kp(db_path)
    _save_pile_kp(db_path)

    response = client.get(
        "/api/v1/commercial/archive?section=archived",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    by_type = {item["product_type"] for item in response.json()}
    assert by_type == {"plates", "piles"}


def test_archive_list_filter_product_type(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    plate_id = _save_plate_kp(db_path)
    pile_id = _save_pile_kp(db_path)

    piles_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=piles",
        cookies=_admin_cookie(),
    )
    assert piles_resp.status_code == 200
    pile_items = piles_resp.json()
    assert len(pile_items) == 1
    assert pile_items[0]["kp_id"] == pile_id
    assert pile_items[0]["product_type"] == "piles"

    plates_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=plates",
        cookies=_admin_cookie(),
    )
    assert plates_resp.status_code == 200
    plate_items = plates_resp.json()
    assert len(plate_items) == 1
    assert plate_items[0]["kp_id"] == plate_id
    assert plate_items[0]["product_type"] == "plates"

    all_resp = client.get(
        "/api/v1/commercial/archive?section=archived&product_type=all",
        cookies=_admin_cookie(),
    )
    assert all_resp.status_code == 200
    assert len(all_resp.json()) == 2


def test_archive_detail_includes_piles(archive_client: tuple[TestClient, str]) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    response = client.get(
        f"/api/v1/commercial/archive/{pile_id}",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["product_type"] == "piles"
    assert payload["plates"] == []
    assert len(payload["piles"]) == 1
    pile = payload["piles"][0]
    assert pile["mark"] == "С120.35-12"
    assert pile["concrete_grade"] == "B25"
    assert pile["qty"] == 3
    assert pile["unit_price"] == pytest.approx(44634.03)


def test_archive_service_maps_pile_details(tmp_path: Path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    pile_id = _save_pile_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path,
    )
    details = service.get_details(pile_id, user={"id": 1, "role": "admin"})

    assert details.product_type == "piles"
    assert len(details.piles) == 1
    assert details.piles[0].mark == "С120.35-12"
    assert details.plates == []


def test_archive_generate_pdf_for_saved_pile_kp(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "outputs",
    )
    (tmp_path / "outputs").mkdir(exist_ok=True)

    path = asyncio.run(
        service.generate_document(pile_id, "pdf", user={"id": 1, "role": "admin"})
    )

    assert path.exists()
    assert path.name == f"КП_{pile_id}.pdf"
    assert path.stat().st_size > 100


def test_archive_download_pdf_http_for_pile_kp(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = archive_client
    pile_id = _save_pile_kp(db_path)

    outputs_dir = tmp_path / "archive_outputs"
    outputs_dir.mkdir()
    monkeypatch.setenv("OUTPUTS_DIR", str(outputs_dir))
    get_settings.cache_clear()

    response = client.get(
        f"/api/v1/commercial/archive/{pile_id}/files/pdf",
        cookies=_admin_cookie(),
    )

    assert response.status_code == 200, response.text
    assert response.headers["content-type"] == "application/pdf"
    assert len(response.content) > 100


# --- MNA-402: archive regen agrees with export (mixed layout + append_batches) ---


MIXED_ORDER = [
    {
        "line_id": "ln_plate",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "mark": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 10,
        "unit_price": 100.0,
        "weight": 2000.0,
        # Explicit grade so resolve_concrete_grade short-circuits (no pb.db).
        "concrete_grade": "М500",
    },
    {
        "line_id": "ln_pile",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 44634.03,
    },
]


def _save_mixed_kp(db_path: str, *, logistics_cost: float = 0.0) -> int:
    return KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        MIXED_ORDER,
        customer_name="Mixed client",
        status="в архиве",
        logistics_cost=logistics_cost,
        db_path=db_path,
    )


def _xlsx_headers_and_names(path: Path) -> tuple[list[str], list[str]]:
    import pandas as pd

    df = pd.read_excel(path, sheet_name="КП", header=None)
    header_idx = None
    headers: list[str] = []
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if vals and vals[0] == "№":
            header_idx = int(i)
            raw = [v if pd.notna(v) else "" for v in row.tolist()]
            while raw and raw[-1] == "":
                raw.pop()
            headers = [str(c).strip() if c != "" else "" for c in raw]
            break
    assert header_idx is not None, f"№ header not found in {path}"
    name_col = headers.index("Наименование")
    names: list[str] = []
    for _, row in df.iloc[header_idx + 1 :].iterrows():
        cells = list(row.tolist())
        if all(pd.isna(v) or str(v).strip() == "" for v in cells):
            continue
        if name_col >= len(cells) or pd.isna(cells[name_col]):
            names.append("")
        else:
            names.append(str(cells[name_col]).strip())
    return headers, names


def test_archive_generate_xlsx_mixed_uses_unified_layout(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
) -> None:
    """MNA-402: archive regen of mixed KP uses unified columns (Тип)."""
    _client, db_path = archive_client
    kp_id = _save_mixed_kp(db_path)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    out = tmp_path / "outputs"
    out.mkdir()
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=out,
    )
    path = asyncio.run(
        service.generate_document(kp_id, "xlsx", user={"id": 1, "role": "admin"})
    )
    headers, _names = _xlsx_headers_and_names(path)
    assert headers == ["№", "Тип", "Наименование", "Кол-во", "Цена", "Сумма"]


def test_archive_generate_xlsx_mixed_pb_only_delivery(
    archive_client: tuple[TestClient, str],
    tmp_path: Path,
) -> None:
    """MNA-402: archive regen delivery row only when PB logistics > 0."""
    _client, db_path = archive_client
    kp_id = _save_mixed_kp(db_path, logistics_cost=5000.0)

    from app.repositories.kp_archive_repository import KpArchiveRepository

    out = tmp_path / "outputs"
    out.mkdir()
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=out,
    )
    path = asyncio.run(
        service.generate_document(kp_id, "xlsx", user={"id": 1, "role": "admin"})
    )
    headers, names = _xlsx_headers_and_names(path)
    assert headers[1] == "Тип"
    assert any("доставк" in n.lower() for n in names), names


def test_archive_generate_pdf_passes_append_batches_for_same_type_multi(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-402: archive regen must pass append_batches so same-type multi is unified.

    Order lines omit append_batch_id (as after DB round-trip); batches live on KP raw
    (or equivalent) and must be forwarded into generate_commercial_offer_pdf.
    """
    batches = [
        {"batch_id": "batch-a", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "batch-b", "product_type": "plates", "line_ids": ["p2"]},
    ]
    order_data = [
        {
            "line_id": "p1",
            "product_type": "plates",
            "name": "ПБ-A",
            "mark": "ПБ-A",
            "length_m": 6.0,
            "width_m": 1.2,
            "qty": 1,
            "unit_price": 1000.0,
            "weight": 500.0,
            "load_class": 800,
        },
        {
            "line_id": "p2",
            "product_type": "plates",
            "name": "ПБ-B",
            "mark": "ПБ-B",
            "length_m": 4.8,
            "width_m": 1.2,
            "qty": 1,
            "unit_price": 900.0,
            "weight": 400.0,
            "load_class": 800,
        },
    ]

    repository = MagicMock()
    repository.get_by_id.return_value = {
        "kp_id": 402,
        "creation_date": "12.08.2026",
        "customer_name": "Same-type multi",
        "manager_name": "Иван",
        "discount_percent": 0.0,
        "logistics_cost": 1000.0,
        "delivery_conditions": None,
        "payment_conditions": None,
        "owner_user_id": 1,
        "product_type": "plates",
        "plates": order_data,
        "append_batches": batches,
    }
    monkeypatch.setattr(
        "app.services.archive_service.order_data_from_kp_info",
        lambda _raw: order_data,
    )

    class FakeBuffer:
        def getvalue(self) -> bytes:
            return b"%PDF-FAKE"

    fake_pdf = MagicMock(return_value=FakeBuffer())
    monkeypatch.setattr(
        "app.services.archive_service.generate_commercial_offer_pdf",
        fake_pdf,
    )

    service = ArchiveService(repository=repository, outputs_dir=tmp_path)
    path = asyncio.run(
        service.generate_document(402, "pdf", user={"id": 1, "role": "admin"})
    )
    assert path.exists()
    fake_pdf.assert_called_once()
    call_kwargs = fake_pdf.call_args.kwargs
    assert "append_batches" in call_kwargs, (
        "ArchiveService.generate_document must pass append_batches into PDF generator"
    )
    assert call_kwargs["append_batches"] == batches


def test_archive_generate_xlsx_passes_append_batches_for_same_type_multi(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-402: archive xlsx regen forwards append_batches for same-type multi."""
    batches = [
        {"batch_id": "batch-a", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "batch-b", "product_type": "plates", "line_ids": ["p2"]},
    ]
    order_data = [
        {
            "line_id": "p1",
            "product_type": "plates",
            "name": "ПБ-A",
            "length_m": 6.0,
            "width_m": 1.2,
            "qty": 1,
            "unit_price": 1000.0,
            "weight": 500.0,
            "load_class": 800,
        },
        {
            "line_id": "p2",
            "product_type": "plates",
            "name": "ПБ-B",
            "length_m": 4.8,
            "width_m": 1.2,
            "qty": 1,
            "unit_price": 900.0,
            "weight": 400.0,
            "load_class": 800,
        },
    ]

    repository = MagicMock()
    repository.get_by_id.return_value = {
        "kp_id": 403,
        "creation_date": "12.08.2026",
        "customer_name": "Same-type multi",
        "manager_name": "Иван",
        "discount_percent": 0.0,
        "logistics_cost": 0.0,
        "delivery_conditions": None,
        "payment_conditions": None,
        "owner_user_id": 1,
        "product_type": "plates",
        "plates": order_data,
        "append_batches": batches,
    }
    monkeypatch.setattr(
        "app.services.archive_service.order_data_from_kp_info",
        lambda _raw: order_data,
    )

    class FakeBuffer:
        def getvalue(self) -> bytes:
            return b"PK\x03\x04fake-xlsx"

    fake_xlsx = MagicMock(return_value=FakeBuffer())
    monkeypatch.setattr(
        "app.services.archive_service.generate_commercial_offer_xlsx",
        fake_xlsx,
    )

    service = ArchiveService(repository=repository, outputs_dir=tmp_path)
    path = asyncio.run(
        service.generate_document(403, "xlsx", user={"id": 1, "role": "admin"})
    )
    assert path.exists()
    fake_xlsx.assert_called_once()
    call_kwargs = fake_xlsx.call_args.kwargs
    assert "append_batches" in call_kwargs, (
        "ArchiveService.generate_document must pass append_batches into XLSX generator"
    )
    assert call_kwargs["append_batches"] == batches
