from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_service import CommercialService
from app.services.commercial_service import CommercialPreviewResult
from app.services.commercial_workflow_service import CommercialWorkflowService


def _sample_draft(draft_id: str = "draft-123") -> dict:
    return {
        "draft_id": draft_id,
        "order": {"plates_1_2": [], "plate_load_details": []},
        "optimization": {"result": {"total_plates": 2}, "total_plates": 2, "total_cost": 1500.0},
        "order_data": [
            {
                "name": "ПБ 78-12-8п",
                "qty": 2,
                "length_m": 7.8,
                "width_m": 1.2,
                "unit_price": 1000.0,
            }
        ],
        "metadata": {
            "source_type": "text",
            "client_name": "ООО Тест",
            "manager_name": "Иван Иванов",
            "discount_percent": 5.0,
            "warnings": [],
            "unparsed_lines": [],
        },
        "files": [],
        "saved_offer": None,
        "totals": {"subtotal": 1639.34, "vat_amount": 360.66, "total_with_vat": 2000.0},
    }


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    get_settings.cache_clear()
    monkeypatch.setattr(
        AuthRepository,
        "list_users",
        lambda self: [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    app = create_app()
    return TestClient(app)


@pytest.fixture()
def auth_cookie() -> dict[str, str]:
    return {
        "app_session": create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300),
    }


def test_create_draft_from_form_text_only(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def fake_create(self, **kwargs):
        assert kwargs["text"] == "ПБ 78-12-8п 2"
        assert kwargs["image_bytes"] is None
        return _sample_draft()

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    response = client.post(
        "/api/v1/commercial/from-form",
        data={
            "text": "ПБ 78-12-8п 2",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "5",
            "delivery_conditions": "Самовывоз",
            "payment_conditions": "100% предоплата",
        },
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["draft_id"] == "draft-123"
    assert payload["metadata"]["client_name"] == "ООО Тест"


def test_create_draft_from_form_image_only(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    async def fake_create(self, **kwargs):
        assert kwargs["text"] == ""
        assert kwargs["image_bytes"] == b"fake-image"
        assert kwargs["image_filename"] == "scan.png"
        return _sample_draft("draft-image")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    response = client.post(
        "/api/v1/commercial/from-form",
        data={
            "text": "",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "0",
            "delivery_conditions": "",
            "payment_conditions": "",
        },
        files={"image": ("scan.png", b"fake-image", "image/png")},
        cookies=auth_cookie,
    )

    assert response.status_code == 200
    assert response.json()["draft_id"] == "draft-image"


def test_commercial_api_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/commercial/drafts/unknown")

    assert response.status_code == 401


def test_download_generated_file_from_outputs_dir(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    workflow = CommercialWorkflowService()
    target_file = Path(workflow.settings.outputs_dir) / "test-download-file.txt"
    target_file.write_text("ok", encoding="utf-8")

    try:
        response = client.get(f"/api/v1/commercial/files/{target_file.name}", cookies=auth_cookie)
    finally:
        target_file.unlink(missing_ok=True)

    assert response.status_code == 200
    assert response.content == b"ok"


def test_download_generated_file_rejects_outside_outputs(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    outside_file = tmp_path / "outside.txt"
    outside_file.write_text("secret", encoding="utf-8")
    monkeypatch.setattr(CommercialWorkflowService, "_resolve_generated_file", lambda self, filename: outside_file)

    response = client.get("/api/v1/commercial/files/outside.txt", cookies=auth_cookie)

    assert response.status_code == 404


def test_web_offer_form_and_redirect(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        CommercialService,
        "list_managers",
        lambda self: [{"id": 1, "fio": "Иван Иванов", "contact_number": "", "email": ""}],
    )

    async def fake_create(self, **kwargs):
        return _sample_draft("draft-web")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    page_response = client.get("/web/offers/new", cookies=auth_cookie)
    submit_response = client.post(
        "/web/offers/new",
        data={
            "text": "ПБ 78-12-8п 2",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "5",
            "delivery_conditions": "Самовывоз",
            "payment_conditions": "100% предоплата",
        },
        cookies=auth_cookie,
        follow_redirects=False,
    )

    assert page_response.status_code == 200
    assert "Новое коммерческое предложение" in page_response.text
    assert submit_response.status_code == 303
    assert submit_response.headers["location"] == "/web/offers/drafts/draft-web"


def test_build_order_data_preserves_input_sequence() -> None:
    service = CommercialService()
    order = PlateOrder()
    parse_result = ParseResult(
        order=order,
        normalized_text="",
        line_plate_load_details=[
            {(7.1, 1.2, 8.0, "71"): 1},
            {(5.9, 1.2, 8.0, "59"): 1},
        ],
    )
    procurement_items = [
        {"length": 5.9, "width": 1.2, "qty": 1, "load_code": 8, "length_dm_raw": "59"},
        {"length": 7.1, "width": 1.2, "qty": 1, "load_code": 8, "length_dm_raw": "71"},
    ]

    order_data = service._build_order_data(procurement_items, [], order, parse_result)

    assert len(order_data) == 2
    assert order_data[0]["length_m"] == pytest.approx(7.1)
    assert order_data[1]["length_m"] == pytest.approx(5.9)


def test_resolve_wide_plates_applies_line_id_with_normalized_lines(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "original_text": "Плиты ПБ 59-15-8п 2",
            "ocr_text": "",
            "input_text": "Плиты ПБ 59-15-8п 2\nПБ 59-12-8п 1",
            "normalized_lines": ["ПБ 59-15-8п 2", "ПБ 59-12-8п 1"],
            "plate_batches": [],
            "wide_plate_lines": [{"id": "wide-1", "line": "ПБ 59-15-8п 2", "qty": 2}],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(
        workflow,
        "_load_draft_or_raise",
        lambda _draft_id: draft_payload,
    )

    captured: dict[str, str] = {}

    def fake_generate_preview(*, text: str | None = None, parse_result: ParseResult | None = None):
        _ = parse_result
        preview_text = text or ""
        captured["text"] = preview_text
        fake_parse_result = ParseResult(
            order=PlateOrder(),
            normalized_text=preview_text,
            normalized_lines=[line for line in preview_text.splitlines() if line.strip()],
        )
        return CommercialPreviewResult(
            parse_result=fake_parse_result,
            optimization_context=OptimizationContext(
                order=PlateOrder(),
                optimization_result={"total_plates": 0, "total_cost": 0.0},
            ),
            order_data=[],
            price_rows=[],
            breakdown_tables=[],
            total_sum=0.0,
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)
    monkeypatch.setattr(
        workflow,
        "_build_preview_metadata",
        lambda **kwargs: {"wide_plate_lines": [], "wide_plates_resolved": True},
    )
    monkeypatch.setattr(workflow.draft_store, "replace_preview", lambda *args, **kwargs: None)
    monkeypatch.setattr(workflow, "get_draft_details", lambda _draft_id: {"draft_id": _draft_id})

    result = workflow.resolve_wide_plates(
        "draft-1",
        decisions=[{"line_id": "wide-1", "action": "exclude"}],
    )

    assert result["draft_id"] == "draft-1"
    assert captured["text"] == "ПБ 59-12-8п 1"
