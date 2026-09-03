from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_draft_service import CommercialDraftService
from app.services.draft_store import DraftStore
from tests.test_commercial_web_flow import _MINIMAL_PNG_BYTES

pytest_plugins = ["tests.commercial_http_fixtures"]

_KNOWN_TEXT = "ПБ 78-12-8п 2"
_FULL_OCR_TEXT = "ПБ 34-15-10п 15\nПБ 60-12-8п 3"


def _drafts_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir(exist_ok=True)
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()
    return drafts_dir


def _draft_file_bytes(drafts_dir: Path, draft_id: str) -> bytes:
    return (drafts_dir / f"{draft_id}.json").read_bytes()


def _fake_extract(ocr_text: str, *, expected_product_type: str | None = None, captured: dict[str, Any] | None = None):
    async def fake_extract(
        self,
        *,
        image_bytes: bytes,
        image_filename: str | None,
        product_type: str = "plates",
    ):
        if captured is not None:
            captured["product_type"] = product_type
            captured["image_bytes"] = image_bytes
        if expected_product_type is not None:
            assert product_type == expected_product_type
        return ocr_text, {
            "ocr_verify_failed": False,
            "ocr_corrections": [{"action": "replaced", "from": "44-15-10п", "to": "34-15-10п"}],
        }

    return fake_extract


def _create_text_draft(client: TestClient) -> str:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={"text": _KNOWN_TEXT},
    )
    assert response.status_code == 200, response.text
    return response.json()["draft_id"]


def test_ocr_page_returns_normalized_text_and_does_not_mutate_draft(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    drafts_dir = _drafts_dir(tmp_path, monkeypatch)
    draft_id = _create_text_draft(client)
    before_get = client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert before_get.status_code == 200
    before_body = before_get.json()
    before_bytes = _draft_file_bytes(drafts_dir, draft_id)
    loaded = DraftStore().load_raw_json(draft_id)
    assert loaded is not None
    assert loaded["metadata"]["input_text"] == _KNOWN_TEXT

    monkeypatch.setattr(
        CommercialDraftService,
        "extract_text_from_image",
        _fake_extract(_FULL_OCR_TEXT),
    )

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["normalized_text"] == _FULL_OCR_TEXT
    assert body["ocr_verify_failed"] is False
    assert body["ocr_corrections"] == [
        {"action": "replaced", "from": "44-15-10п", "to": "34-15-10п"},
    ]
    assert set(body) == {"normalized_text", "ocr_verify_failed", "ocr_corrections"}

    after_bytes = _draft_file_bytes(drafts_dir, draft_id)
    assert after_bytes == before_bytes
    after_get = client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert after_get.status_code == 200
    after_body = after_get.json()
    assert after_body["metadata"]["input_text"] == before_body["metadata"]["input_text"]
    assert after_body["metadata"]["plate_batches"] == before_body["metadata"]["plate_batches"]
    assert after_body["metadata"]["wide_plate_lines"] == before_body["metadata"]["wide_plate_lines"]
    assert after_body["order_data"] == before_body["order_data"]


def test_ocr_page_uses_draft_product_type(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    drafts_dir = _drafts_dir(tmp_path, monkeypatch)
    draft_id = "p" * 32
    order = PlateOrder()
    DraftStore().replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[{"name": "С120.35-12", "qty": 2, "product_type": "piles"}],
        metadata={
            "owner_user_id": 1,
            "product_type": "piles",
            "input_text": "С120.35-12 B25 2",
        },
    )
    captured: dict[str, Any] = {}
    monkeypatch.setattr(
        CommercialDraftService,
        "extract_text_from_image",
        _fake_extract("С120.35-12 2", expected_product_type="piles", captured=captured),
    )

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 200, response.text
    assert captured["product_type"] == "piles"
    assert (drafts_dir / f"{draft_id}.json").is_file()


def test_ocr_page_404_unknown_draft(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _drafts_dir(tmp_path, monkeypatch)
    response = client.post(
        f"/api/v1/commercial/drafts/{'z' * 32}/ocr-page",
        files={"image": ("page.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 404


def test_ocr_page_400_missing_or_empty_image(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _drafts_dir(tmp_path, monkeypatch)
    draft_id = _create_text_draft(client)

    missing = client.post(f"/api/v1/commercial/drafts/{draft_id}/ocr-page")
    assert missing.status_code == 400

    empty = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", b"", "image/png")},
    )
    assert empty.status_code == 400

    bad_magic = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", b"not-a-real-image-bytes", "image/png")},
    )
    assert bad_magic.status_code == 400


def test_ocr_page_503_when_ocr_disabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _drafts_dir(tmp_path, monkeypatch)
    draft_id = _create_text_draft(client)
    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    get_settings.cache_clear()

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 503


def test_ocr_page_403_foreign_draft(
    client_two_users: TestClient,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.security.session import create_session_token

    _drafts_dir(tmp_path, monkeypatch)
    draft_id = "b" * 32
    order = PlateOrder()
    DraftStore().replace_preview(
        draft_id,
        order=order,
        optimization_context=OptimizationContext(order=order),
        order_data=[{"name": "n", "qty": 1, "length_m": 1.0, "width_m": 1.0, "unit_price": 1.0}],
        metadata={"owner_user_id": 1, "product_type": "plates", "input_text": _KNOWN_TEXT},
    )
    token = create_session_token({"id": 2, "username": "bob", "role": "manager"}, ttl_seconds=300)
    client_two_users.cookies.set("app_session", token)
    response = client_two_users.post(
        f"/api/v1/commercial/drafts/{draft_id}/ocr-page",
        files={"image": ("page.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 403
