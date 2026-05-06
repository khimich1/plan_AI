from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import quote

import pytest
from fastapi.testclient import TestClient

from app.core.http_errors import MSG_INTERNAL, MSG_PARSE_FAILED, MSG_VALIDATION
from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_service import CommercialService
from app.services.commercial_service import CommercialPreviewResult
from app.schemas.commercial import WizardNextRequiredAction, WizardStepId
from app.services.commercial_upload_validation import reset_commercial_ocr_rate_limiter_for_tests
from app.services.commercial_workflow_service import CommercialWorkflowService, _safe_ocr_temp_suffix
from app.services.draft_store import DraftStore, UnsafeDraftIdError
from core.exceptions import PlateParseError

# Minimal valid 1×1 PNG (magic bytes + structure) for upload validation tests.
_MINIMAL_PNG_BYTES = (
    b"\x89PNG\r\n\x1a\n"
    b"\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
    b"\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
)


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
            "manager_id": 1,
            "manager_name": "Иван Иванов",
            "discount_percent": 5.0,
            "conditions_mode": "standard",
            "delivery_conditions": "Самовывоз",
            "payment_conditions": "100% предоплата",
            "warnings": [],
            "unparsed_lines": [],
            "normalized_text": "ПБ 78-12-8п 2",
            "normalized_lines": ["ПБ 78-12-8п 2"],
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "current_step": "client",
        },
        "wizard_state": {
            "current_step": "client",
            "can_proceed_to": [],
            "next_required_action": "post_calculate",
            "validation_errors": [],
        },
        "files": [],
        "saved_offer": None,
        "totals": {"subtotal": 1639.34, "vat_amount": 360.66, "total_with_vat": 2000.0},
        "offer_identity": {
            "offer_number": f"WEB_{draft_id[:8].upper()}",
            "offer_date": "05.05.2026",
            "file_stem": f"kp_{draft_id[:8]}",
        },
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
def auth_cookie(client: TestClient) -> dict[str, str]:
    """Attach session to ``client``. Prefer relying on ``client.cookies`` instead of per-request ``cookies=``."""
    token = create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    return {"app_session": token}


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
        assert kwargs["image_bytes"] == _MINIMAL_PNG_BYTES
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
        files={"image": ("scan.png", _MINIMAL_PNG_BYTES, "image/png")},
    )

    assert response.status_code == 200
    assert response.json()["draft_id"] == "draft-image"


def test_commercial_api_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/commercial/drafts/unknown")

    assert response.status_code == 401


def test_commercial_parse_does_not_leak_unexpected_exception_detail(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    leak_token = "KP004_SYNTHETIC_UNEXPECT_ABC123XYZ"

    def boom(self, text: str):
        raise RuntimeError(leak_token)

    monkeypatch.setattr(CommercialService, "parse", boom)
    response = client.post("/api/v1/commercial/parse", json={"text": "x"})

    assert response.status_code == 500
    body = response.json()
    assert leak_token not in response.text
    assert body["detail"] == MSG_INTERNAL


def test_commercial_parse_does_not_leak_plate_parse_error_message(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    leak_token = "KP004_SYNTHETIC_PLATE_PARSE_SECRET"

    def boom(self, text: str):
        raise PlateParseError(leak_token)

    monkeypatch.setattr(CommercialService, "parse", boom)
    response = client.post("/api/v1/commercial/parse", json={"text": "x"})

    assert response.status_code == 400
    body = response.json()
    assert leak_token not in response.text
    assert body["detail"] == MSG_PARSE_FAILED


def test_commercial_from_form_does_not_leak_value_error_message(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    leak_token = "KP004_SYNTHETIC_VALIDATION_SECRET"

    async def boom(self, **kwargs):
        raise ValueError(leak_token)

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", boom)

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
    )

    assert response.status_code == 400
    body = response.json()
    assert leak_token not in response.text
    assert body["detail"] == MSG_VALIDATION


@pytest.mark.parametrize(
    "unsafe_id",
    [
        "../escape",
        r"..\escape",
        "foo/bar",
        r"foo\bar",
        "/absolute",
        "",
        " trim",
        "has space",
        "double..dot",
        "x" * 129,
        "тест",
        "café",
        "draft\u200bid",
        "pb_\U0001f600_plate",
        "%2e%2e%2fetc%2fpasswd",
        "%252e%252e%252fstatic",
    ],
)
def test_draft_store_rejects_unsafe_ids(
    unsafe_id: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path))
    get_settings.cache_clear()
    store = DraftStore()
    with pytest.raises(UnsafeDraftIdError):
        store.validate_draft_id(unsafe_id)


@pytest.mark.parametrize(
    "safe_id",
    [
        "868a336d4a624fd59b6363733840a153",
        "draft-web",
        "draft-123",
        "WEB_01ABCD",
        "a",
    ],
)
def test_draft_store_accepts_safe_ids(
    safe_id: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path))
    get_settings.cache_clear()
    DraftStore.validate_draft_id(safe_id)


def test_draft_store_validate_rejects_non_string_ids() -> None:
    with pytest.raises(UnsafeDraftIdError):
        DraftStore.validate_draft_id(None)  # type: ignore[arg-type]
    with pytest.raises(UnsafeDraftIdError):
        DraftStore.validate_draft_id(b"deadbeef")  # type: ignore[arg-type]


def test_draft_store_validate_rejects_whitespace_only_id(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path))
    get_settings.cache_clear()
    store = DraftStore()
    with pytest.raises(UnsafeDraftIdError):
        store.validate_draft_id("   ")
    with pytest.raises(UnsafeDraftIdError):
        store.validate_draft_id("\t\n")


def test_draft_store_get_path_rejects_symlink_escaping_base_dir(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Resolved path must stay under drafts_dir (followed symlinks cannot escape)."""
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    outside_dir = tmp_path / "outside"
    outside_dir.mkdir()
    target_file = outside_dir / "payload.json"
    target_file.write_text("{}", encoding="utf-8")
    link_path = drafts_dir / "evil.json"
    try:
        os.symlink(target_file, link_path)
    except OSError:
        pytest.skip("could not create symlink for DraftStore containment test")

    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()
    store = DraftStore()

    with pytest.raises(UnsafeDraftIdError):
        store._get_path("evil")


def test_commercial_get_draft_rejects_path_traversal_id(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    malicious = quote("../etc/passwd", safe="")
    response = client.get(f"/api/v1/commercial/drafts/{malicious}")
    assert response.status_code == 404


def test_commercial_get_draft_rejects_double_url_encoded_traversal(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    inner = quote("../etc/passwd", safe="")
    double_encoded = quote(inner, safe="")
    response = client.get(f"/api/v1/commercial/drafts/{double_encoded}")
    assert response.status_code == 404


def test_commercial_patch_meta_rejects_path_traversal_id(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    malicious = quote("../etc/passwd", safe="")
    response = client.patch(
        f"/api/v1/commercial/drafts/{malicious}/meta",
        json={"client_name": "X"},
    )
    assert response.status_code == 404


def test_commercial_patch_plates_rejects_path_traversal_id(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    malicious = quote("../etc/passwd", safe="")
    response = client.patch(
        f"/api/v1/commercial/drafts/{malicious}/plates",
        data={"mode": "append", "text": ""},
    )
    assert response.status_code == 404


def test_safe_ocr_temp_suffix_whitelists_extensions() -> None:
    assert _safe_ocr_temp_suffix("scan.png") == ".png"
    assert _safe_ocr_temp_suffix("../../x.JPEG") == ".jpeg"
    assert _safe_ocr_temp_suffix("report.pdf") == ".pdf"


def test_commercial_upload_rejects_oversized_file(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("COMMERCIAL_UPLOAD_MAX_BYTES", "1024")
    get_settings.cache_clear()
    body = b"x" * 1025
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
        files={"image": ("big.png", body, "image/png")},
    )
    assert response.status_code == 413


def test_commercial_upload_rejects_bad_magic_bytes(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
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
        files={"image": ("fake.png", b"not-a-real-image-bytes", "image/png")},
    )
    assert response.status_code == 400


def test_commercial_upload_rate_limit_returns_429(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("COMMERCIAL_OCR_UPLOADS_PER_HOUR", "2")
    get_settings.cache_clear()
    reset_commercial_ocr_rate_limiter_for_tests()

    async def fake_create(self, **kwargs):
        return _sample_draft("draft-rl")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    form = {
        "text": "",
        "manager_id": "1",
        "client_name": "ООО Тест",
        "discount_percent": "0",
        "delivery_conditions": "",
        "payment_conditions": "",
    }
    files = {"image": ("a.png", _MINIMAL_PNG_BYTES, "image/png")}
    assert client.post("/api/v1/commercial/from-form", data=form, files=files).status_code == 200
    assert client.post("/api/v1/commercial/from-form", data=form, files=files).status_code == 200
    assert client.post("/api/v1/commercial/from-form", data=form, files=files).status_code == 429


def test_commercial_create_draft_accepts_valid_png_upload(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    (tmp_path / "drafts").mkdir()
    get_settings.cache_clear()
    reset_commercial_ocr_rate_limiter_for_tests()

    async def fake_create(self, **kwargs):
        assert kwargs["image_bytes"] == _MINIMAL_PNG_BYTES
        assert kwargs["image_filename"] == "plates.png"
        return _sample_draft("draft-png")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft", fake_create)

    response = client.post(
        "/api/v1/commercial/drafts",
        data={"text": ""},
        files={"image": ("plates.png", _MINIMAL_PNG_BYTES, "application/octet-stream")},
    )
    assert response.status_code == 200
    assert response.json()["draft_id"] == "draft-png"


def test_download_generated_file_from_outputs_dir(
    client: TestClient,
    auth_cookie: dict[str, str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()

    draft_id = "a" * 32
    store = DraftStore()
    order = PlateOrder()
    oc = OptimizationContext(order=order)
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=oc,
        order_data=[],
        metadata={
            "owner_user_id": 1,
            "generated_files": [
                {
                    "kind": "pdf",
                    "filename": "test-download-file.txt",
                    "display_name": "t",
                    "download_url": "",
                },
            ],
        },
    )

    workflow = CommercialWorkflowService()
    target_file = Path(workflow.settings.outputs_dir) / "test-download-file.txt"
    target_file.write_text("ok", encoding="utf-8")

    try:
        response = client.get(
            "/api/v1/commercial/files/test-download-file.txt",
            params={"draft_id": draft_id},
        )
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
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()

    draft_id = "c" * 32
    store = DraftStore()
    order = PlateOrder()
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[],
        metadata={
            "owner_user_id": 1,
            "generated_files": [
                {
                    "kind": "pdf",
                    "filename": "outside.txt",
                    "display_name": "t",
                    "download_url": "",
                },
            ],
        },
    )

    outside_file = tmp_path / "outside.txt"
    outside_file.write_text("secret", encoding="utf-8")
    monkeypatch.setattr(CommercialWorkflowService, "_resolve_generated_file", lambda self, filename: outside_file)

    response = client.get(
        "/api/v1/commercial/files/outside.txt",
        params={"draft_id": draft_id},
    )

    assert response.status_code == 404


@pytest.fixture()
def client_two_users(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key")
    get_settings.cache_clear()
    monkeypatch.setattr(
        AuthRepository,
        "list_users",
        lambda self: [
            {
                "id": 1,
                "username": "alice",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            },
            {
                "id": 2,
                "username": "bob",
                "role": "manager",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            },
        ],
    )
    app = create_app()
    return TestClient(app)


def test_draft_idor_forbids_non_owner(
    client_two_users: TestClient,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()
    draft_id = "b" * 32
    store = DraftStore()
    order = PlateOrder()
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[{"name": "n", "qty": 1, "length_m": 1.0, "width_m": 1.0, "unit_price": 1.0}],
        metadata={"owner_user_id": 1},
    )
    token = create_session_token({"id": 2, "username": "bob", "role": "manager"}, ttl_seconds=300)
    client_two_users.cookies.set("app_session", token)
    response = client_two_users.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert response.status_code == 403


def test_draft_idor_allows_owner(
    client_two_users: TestClient,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()
    draft_id = "d" * 32
    store = DraftStore()
    order = PlateOrder()
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[{"name": "n", "qty": 1, "length_m": 1.0, "width_m": 1.0, "unit_price": 1.0}],
        metadata={"owner_user_id": 1},
    )
    token = create_session_token({"id": 1, "username": "alice", "role": "admin"}, ttl_seconds=300)
    client_two_users.cookies.set("app_session", token)
    response = client_two_users.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert response.status_code == 200
    assert response.json()["draft_id"] == draft_id


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

    page_response = client.get("/web/offers/new", follow_redirects=False)
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
        follow_redirects=False,
    )

    assert page_response.status_code == 303
    assert page_response.headers["location"] == "/commercial-offer/new"
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


def test_wizard_state_ingest_plates_validation_errors() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {"current_step": WizardStepId.plates.value, "wide_plate_lines": [], "wide_plates_resolved": True},
        "order_data": [],
    }
    state = wf.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.ingest_plates
    assert state["validation_errors"] == ["Список плит пустой."]
    assert state["can_proceed_to"] == []


def test_wizard_state_wide_plates_blocks_forward() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wf.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == [WizardStepId.wide_plates]
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates
    assert state["validation_errors"] == ["Сначала обработайте плиты шире 12 дм."]


def test_wizard_state_wide_plates_step_locked_until_resolve() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {
            "current_step": WizardStepId.wide_plates.value,
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wf.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.wide_plates
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates
    assert state["validation_errors"] == ["Сначала обработайте плиты шире 12 дм."]


def test_wizard_state_client_requires_calculate() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {
            "current_step": WizardStepId.client.value,
            "manager_id": 1,
            "client_name": "ООО А",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wf.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.client
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.post_calculate
    assert state["validation_errors"] == []


def test_wizard_state_result_after_calculate() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {
            "current_step": WizardStepId.result.value,
            "manager_id": 1,
            "client_name": "ООО А",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wf.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.result
    assert state["next_required_action"] == WizardNextRequiredAction.none
    assert state["validation_errors"] == []


def test_wizard_state_select_manager_validation_errors() -> None:
    wf = CommercialWorkflowService()
    payload = {
        "metadata": {
            "current_step": WizardStepId.manager.value,
            "manager_id": None,
            "client_name": "",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wf.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.select_manager
    assert state["validation_errors"] == ["Выберите менеджера."]
