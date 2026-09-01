"""MNA-401/402: multi/append unified export + mono R3 + service wiring.

Detect multi mode → unified columns ``№ | Тип | Наименование | Кол-во | Цена | Сумма``.
Mono one-type one-shot keeps existing plate/pile/step layouts (R3).

Mode detection (``is_unified_commercial_document``):
- ``False`` for mono one-shot (0–1 append batch, single product_type)
- ``True`` when ≥2 distinct ``product_type`` values in ``order_data``
- ``True`` when ``len(append_batches) > 1`` (incl. same-type plates→plates append)
- ``True`` when ≥2 distinct non-empty ``append_batch_id`` on lines

MNA-402: ``CommercialExportService`` / archive regen must forward ``append_batches``
(or equivalent) into generators so same-type multi-append stays unified and both
paths agree on mixed layout + PB-only delivery.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable
from unittest.mock import MagicMock

import pandas as pd
import pytest

from core.commercial_line_format import format_line_name
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

UNIFIED_HEADERS = ["№", "Тип", "Наименование", "Кол-во", "Цена", "Сумма"]
MONO_PLATES_HEADERS = ["№", "Наименование", "Кол-во", "Ед.", "Вес(кг)", "Цена", "Сумма"]
MONO_PILES_HEADERS = ["№", "Наименование", "Класс бетона", "Кол-во", "Цена", "Сумма"]
MONO_STEPS_HEADERS = ["№", "Наименование", "Кол-во", "Цена", "Сумма"]


def _require_is_unified() -> Callable[..., bool]:
    try:
        from core.commercial_offer import is_unified_commercial_document
    except ImportError as exc:  # pragma: no cover - RED until MNA-401
        pytest.fail(f"is_unified_commercial_document missing: {exc}")
    return is_unified_commercial_document


def _require_table_headers(module: str) -> Callable[..., list[str]]:
    """Resolve ``commercial_offer_table_headers`` from pdf or xlsx module."""
    if module == "pdf":
        import core.commercial_offer as mod
    else:
        import core.commercial_offer_xlsx as mod
    fn = getattr(mod, "commercial_offer_table_headers", None)
    if fn is None:
        pytest.fail(f"{mod.__name__}.commercial_offer_table_headers missing (MNA-401)")
    return fn


# --- fixtures / helpers -------------------------------------------------------


def _plate(**overrides: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "line_id": "ln_plate",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "mark": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 2,
        "unit_price": 1000.0,
        "weight": 500.0,
        # Explicit grade so resolve_concrete_grade short-circuits (no pb.db).
        "concrete_grade": "М500",
    }
    base.update(overrides)
    return base


def _pile(**overrides: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "line_id": "ln_pile",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 3,
        "unit_price": 44634.03,
    }
    base.update(overrides)
    return base


def _step(**overrides: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "line_id": "ln_step",
        "product_type": "steps",
        "product_kind": "step",
        "name": "ЛС-12",
        "mark": "ЛС-12",
        "qty": 10,
        "unit_price": 1500.0,
    }
    base.update(overrides)
    return base


def _xlsx_header_and_body(
    order_data: list[dict[str, Any]],
    *,
    append_batches: list[dict[str, Any]] | None = None,
    logistics_cost: float = 0.0,
    pile_logistics_cost: float = 0.0,
    pile_trip_overrides: dict[str, int] | None = None,
    pile_catalog_db_path: str | None = None,
) -> tuple[list[str], pd.DataFrame]:
    kwargs: dict[str, Any] = {
        "order_data": order_data,
        "offer_number": "MNA401",
        "offer_date": "12.08.2026",
        "customer_name": "ООО Тест",
        "kp_db_id": 401,
        "logistics_cost": logistics_cost,
        "pile_logistics_cost": pile_logistics_cost,
        "pile_trip_overrides": pile_trip_overrides,
        "pile_catalog_db_path": pile_catalog_db_path,
    }
    if append_batches is not None:
        kwargs["append_batches"] = append_batches
    buf = generate_commercial_offer_xlsx(**kwargs)
    df = pd.read_excel(buf, sheet_name="КП", header=None)
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
    assert header_idx is not None, "table header row with № not found in XLSX"
    body = df.iloc[header_idx + 1 :].copy()
    body.columns = range(len(body.columns))
    return headers, body


def _body_column(body: pd.DataFrame, headers: list[str], name: str) -> list[str]:
    assert name in headers, f"column {name!r} missing from headers {headers}"
    col = headers.index(name)
    out: list[str] = []
    for _, row in body.iterrows():
        cells = list(row.tolist())
        if all(pd.isna(v) or str(v).strip() == "" for v in cells):
            continue
        if col >= len(cells) or pd.isna(cells[col]):
            out.append("")
        else:
            out.append(str(cells[col]).strip())
    return out


# --- Mode detection -----------------------------------------------------------


def test_is_unified_false_for_mono_plates_one_shot() -> None:
    """R3: one type, no multi append → classic layout (not unified)."""
    is_unified = _require_is_unified()
    order = [_plate(line_id="a"), _plate(line_id="b", name="ПБ 56-6-8п", mark="ПБ 56-6-8п")]
    assert is_unified(order) is False
    assert is_unified(order, append_batches=None) is False
    assert (
        is_unified(
            order,
            append_batches=[
                {
                    "batch_id": "only-one",
                    "product_type": "plates",
                    "line_ids": ["a", "b"],
                }
            ],
        )
        is False
    )


def test_is_unified_false_for_mono_piles_one_shot() -> None:
    is_unified = _require_is_unified()
    order = [_pile()]
    assert is_unified(order) is False
    assert (
        is_unified(
            order,
            append_batches=[
                {"batch_id": "b1", "product_type": "piles", "line_ids": ["ln_pile"]}
            ],
        )
        is False
    )


def test_is_unified_false_for_legacy_plates_without_product_type() -> None:
    """Legacy mono lines (no product_type) stay classic."""
    is_unified = _require_is_unified()
    order = [
        {
            "name": "ПБ 60-12-8п",
            "length_m": 6.0,
            "width_m": 1.2,
            "qty": 1,
            "unit_price": 1000.0,
        }
    ]
    assert is_unified(order) is False


def test_is_unified_true_for_mixed_product_types() -> None:
    is_unified = _require_is_unified()
    order = [_plate(line_id="p1"), _pile(line_id="s1")]
    assert is_unified(order) is True


def test_is_unified_true_when_append_batches_gt_one_same_type() -> None:
    """Plates→plates append: still unified even though one product_type."""
    is_unified = _require_is_unified()
    order = [
        _plate(line_id="p1", append_batch_id="batch-a"),
        _plate(line_id="p2", append_batch_id="batch-b", name="ПБ 48-12-8п"),
    ]
    batches = [
        {"batch_id": "batch-a", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "batch-b", "product_type": "plates", "line_ids": ["p2"]},
    ]
    assert is_unified(order, append_batches=batches) is True


def test_is_unified_true_from_multiple_append_batch_ids_without_metadata() -> None:
    """Export may only receive order_data — detect multi from append_batch_id."""
    is_unified = _require_is_unified()
    order = [
        _plate(line_id="p1", append_batch_id="batch-a"),
        _pile(line_id="s1", append_batch_id="batch-b"),
    ]
    assert is_unified(order) is True
    assert is_unified(order, append_batches=[]) is True


def test_is_unified_false_for_empty_order() -> None:
    is_unified = _require_is_unified()
    assert is_unified([]) is False


# --- Shared headers helper (PDF + XLSX contract) ------------------------------


def test_table_headers_mono_plates_unchanged() -> None:
    order = [_plate()]
    assert _require_table_headers("pdf")(order) == MONO_PLATES_HEADERS
    assert _require_table_headers("xlsx")(order) == MONO_PLATES_HEADERS


def test_table_headers_mono_piles_unchanged() -> None:
    order = [_pile()]
    assert _require_table_headers("pdf")(order) == MONO_PILES_HEADERS
    assert _require_table_headers("xlsx")(order) == MONO_PILES_HEADERS


def test_table_headers_mono_steps_unchanged() -> None:
    order = [_step()]
    assert _require_table_headers("pdf")(order) == MONO_STEPS_HEADERS
    assert _require_table_headers("xlsx")(order) == MONO_STEPS_HEADERS


def test_table_headers_unified_for_mixed() -> None:
    order = [_plate(line_id="p1"), _pile(line_id="s1")]
    assert _require_table_headers("pdf")(order) == UNIFIED_HEADERS
    assert _require_table_headers("xlsx")(order) == UNIFIED_HEADERS


def test_table_headers_unified_for_multi_append_same_type() -> None:
    order = [
        _plate(line_id="p1", append_batch_id="a"),
        _plate(line_id="p2", append_batch_id="b"),
    ]
    batches = [
        {"batch_id": "a", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "b", "product_type": "plates", "line_ids": ["p2"]},
    ]
    assert _require_table_headers("pdf")(order, append_batches=batches) == UNIFIED_HEADERS
    assert _require_table_headers("xlsx")(order, append_batches=batches) == UNIFIED_HEADERS


# --- XLSX mono regression (R3) ------------------------------------------------


def test_xlsx_mono_plates_column_layout_unchanged() -> None:
    headers, _body = _xlsx_header_and_body([_plate()])
    assert headers == MONO_PLATES_HEADERS
    assert "Тип" not in headers


def test_xlsx_mono_piles_column_layout_unchanged() -> None:
    headers, _body = _xlsx_header_and_body([_pile()])
    assert headers == MONO_PILES_HEADERS
    assert "Тип" not in headers


def test_xlsx_mono_plates_with_stamped_product_type_stays_classic() -> None:
    """MNA-102 stamps product_type on mono — must not flip to unified alone."""
    order = [
        _plate(line_id="a"),
        _plate(line_id="b", name="ПБ 56-6-8п", mark="ПБ 56-6-8п"),
    ]
    # One sealed batch (calculate seals once) is still mono — do not pass
    # append_batches here; detector contract is covered in unit tests above.
    headers, _body = _xlsx_header_and_body(order)
    assert headers == MONO_PLATES_HEADERS


# --- XLSX multi / unified -----------------------------------------------------


def test_xlsx_mixed_uses_unified_headers_with_type_column() -> None:
    order = [_plate(line_id="p1"), _pile(line_id="s1")]
    headers, _body = _xlsx_header_and_body(order)
    assert headers == UNIFIED_HEADERS


def test_xlsx_mixed_names_use_format_line_name_with_grade() -> None:
    plate = _plate(line_id="p1")
    pile = _pile(line_id="s1", mark="С30.15-3", name="С30.15-3", concrete_grade="B25")
    headers, body = _xlsx_header_and_body([plate, pile])
    names = _body_column(body, headers, "Наименование")
    assert names[0] == format_line_name(plate)
    assert names[1] == format_line_name(pile)
    assert names[1] == "С30.15-3 (B25)"
    assert "Класс бетона" not in headers


def test_xlsx_mixed_rows_are_chronological_not_grouped_by_type() -> None:
    """Плиты→Сваи→Плиты: order preserved; no regrouping by type."""
    order = [
        _plate(line_id="p1", name="ПБ-A", mark="ПБ-A"),
        _pile(line_id="s1", mark="Свая-B", name="Свая-B"),
        _plate(line_id="p2", name="ПБ-C", mark="ПБ-C", qty=1),
    ]
    headers, body = _xlsx_header_and_body(order)
    types = _body_column(body, headers, "Тип")
    names = _body_column(body, headers, "Наименование")
    assert types[:3] == ["Плиты", "Сваи", "Плиты"]
    assert names[0] == format_line_name(order[0])
    assert names[1] == format_line_name(order[1])
    assert names[2] == format_line_name(order[2])


def test_xlsx_same_type_multi_append_uses_unified_layout() -> None:
    """Same-type multi: ≥2 append_batch_id on lines → unified (no metadata required)."""
    order = [
        _plate(line_id="p1", append_batch_id="b1"),
        _plate(line_id="p2", append_batch_id="b2", name="ПБ 48-12-8п", mark="ПБ 48-12-8п"),
    ]
    headers, body = _xlsx_header_and_body(order)
    assert headers == UNIFIED_HEADERS
    types = _body_column(body, headers, "Тип")
    assert types[:2] == ["Плиты", "Плиты"]


# --- Delivery row (PB only) ---------------------------------------------------


def test_xlsx_unified_includes_delivery_when_pb_delivery_positive() -> None:
    order = [
        _plate(line_id="p1", qty=10, weight=2000.0, unit_price=100.0),
        _pile(line_id="s1"),
    ]
    headers, body = _xlsx_header_and_body(order, logistics_cost=5000.0)
    assert headers == UNIFIED_HEADERS
    names = _body_column(body, headers, "Наименование")
    assert any("доставк" in n.lower() for n in names), names


def test_xlsx_unified_omits_delivery_when_logistics_zero() -> None:
    order = [_plate(line_id="p1"), _pile(line_id="s1")]
    headers, body = _xlsx_header_and_body(order, logistics_cost=0.0)
    assert headers == UNIFIED_HEADERS
    names = _body_column(body, headers, "Наименование")
    assert not any("доставк" in n.lower() for n in names), names


def test_xlsx_unified_omits_delivery_when_no_plate_weight() -> None:
    """Non-plates only: no PB mass → no delivery row even if trip cost set."""
    order = [_pile(line_id="s1"), _step(line_id="t1")]
    headers, body = _xlsx_header_and_body(order, logistics_cost=9000.0)
    assert headers == UNIFIED_HEADERS
    names = _body_column(body, headers, "Наименование")
    assert not any("доставк" in n.lower() for n in names), names


def _seed_export_pile_catalog(tmp_path: Path) -> str:
    from core.kp_db_schema import init_schema
    from core.pile_catalog import PileCatalogEntry, upsert_pile_catalog

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    upsert_pile_catalog(
        db_path,
        [PileCatalogEntry("С60.30", 6.0, 300, 0.55, 1380.0, 14)],
    )
    return db_path


def test_xlsx_two_delivery_rows_when_plate_and_pile_ready(tmp_path: Path) -> None:
    db_path = _seed_export_pile_catalog(tmp_path)
    order = [
        _plate(line_id="p1", qty=65, length_m=1.0, width_m=1.0),
        _pile(line_id="s1", mark="С60.30", name="С60.30", qty=14, unit_price=50.0),
    ]
    headers, body = _xlsx_header_and_body(
        order,
        logistics_cost=1000.0,
        pile_logistics_cost=2000.0,
        pile_catalog_db_path=db_path,
    )
    names = _body_column(body, headers, "Наименование")
    assert "Доставка плит" in names
    assert "Доставка свай" in names


def test_xlsx_omits_pile_delivery_when_pending(tmp_path: Path) -> None:
    db_path = _seed_export_pile_catalog(tmp_path)
    order = [
        _pile(
            line_id="s1",
            mark="C18-40T8",
            name="C18-40T8",
            qty=49,
            unit_price=10.0,
            product_type="bridge_piles",
            product_kind="bridge_pile",
        )
    ]
    headers, body = _xlsx_header_and_body(
        order,
        pile_logistics_cost=9000.0,
        pile_catalog_db_path=db_path,
    )
    names = _body_column(body, headers, "Наименование")
    assert not any("доставк" in n.lower() for n in names), names


def test_xlsx_pile_only_ready_shows_pile_delivery_row(tmp_path: Path) -> None:
    db_path = _seed_export_pile_catalog(tmp_path)
    order = [_pile(line_id="s1", mark="С60.30", name="С60.30", qty=14, unit_price=50.0)]
    headers, body = _xlsx_header_and_body(
        order,
        pile_logistics_cost=1500.0,
        pile_catalog_db_path=db_path,
    )
    names = _body_column(body, headers, "Наименование")
    assert "Доставка свай" in names
    assert "Доставка плит" not in names
    assert not any("услуга по доставке" in n.lower() for n in names)


# --- PDF smoke via shared headers (no PDF text extractor in deps) -------------


def test_pdf_generate_mixed_does_not_raise_and_uses_unified_headers() -> None:
    order = [_plate(line_id="p1"), _pile(line_id="s1")]
    assert _require_table_headers("pdf")(order) == UNIFIED_HEADERS
    buf = generate_commercial_offer_pdf(
        order_data=order,
        offer_number="MNA401",
        offer_date="12.08.2026",
        customer_name="ООО Тест",
        kp_db_id=401,
        logistics_cost=0.0,
    )
    data = buf.getvalue()
    assert data[:4] == b"%PDF"
    assert len(data) > 100


def test_pdf_generate_accepts_append_batches_kwarg_for_same_type_multi() -> None:
    order = [
        _plate(line_id="p1", append_batch_id="b1"),
        _plate(line_id="p2", append_batch_id="b2"),
    ]
    batches = [
        {"batch_id": "b1", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "b2", "product_type": "plates", "line_ids": ["p2"]},
    ]
    buf = generate_commercial_offer_pdf(
        order_data=order,
        offer_number="MNA401",
        offer_date="12.08.2026",
        customer_name="ООО Тест",
        append_batches=batches,
        logistics_cost=1000.0,
    )
    assert buf.getvalue()[:4] == b"%PDF"
    assert _require_table_headers("pdf")(order, append_batches=batches) == UNIFIED_HEADERS


# --- MNA-402: CommercialExportService forwards append_batches + PB logistics ---


def _same_type_multi_batches() -> list[dict[str, Any]]:
    return [
        {"batch_id": "batch-a", "product_type": "plates", "line_ids": ["p1"]},
        {"batch_id": "batch-b", "product_type": "plates", "line_ids": ["p2"]},
    ]


def _same_type_multi_order_without_line_batch_ids() -> list[dict[str, Any]]:
    """Same-type multi where only metadata.append_batches proves multi-append.

    Lines intentionally omit ``append_batch_id`` so unified layout requires
    forwarding ``append_batches`` into the generators.
    """
    return [
        _plate(line_id="p1", name="ПБ-A", mark="ПБ-A"),
        _plate(line_id="p2", name="ПБ-B", mark="ПБ-B", qty=1),
    ]


def _export_service_payload(
    order_data: list[dict[str, Any]],
    *,
    append_batches: list[dict[str, Any]] | None = None,
    logistics_cost: float = 0.0,
    product_type: str = "plates",
) -> dict[str, Any]:
    metadata: dict[str, Any] = {
        "client_name": "ООО MNA402",
        "manager_name": "Менеджер",
        "discount_percent": 0.0,
        "logistics_cost": logistics_cost,
        "delivery_conditions": "",
        "payment_conditions": "",
        "product_type": product_type,
        "generated_files": [],
    }
    if append_batches is not None:
        metadata["append_batches"] = append_batches
    return {"order_data": order_data, "metadata": metadata}


def _make_export_service(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    file_generation_service: Any | None = None,
):
    from app.core.settings import get_settings
    from app.services.commercial_export_service import CommercialExportService
    from app.services.draft_store import DraftStore
    from app.services.file_generation_service import FileGenerationService

    outputs = tmp_path / "outputs"
    drafts = tmp_path / "drafts"
    outputs.mkdir(parents=True, exist_ok=True)
    drafts.mkdir(parents=True, exist_ok=True)
    monkeypatch.setenv("OUTPUTS_DIR", str(outputs))
    monkeypatch.setenv("DRAFTS_DIR", str(drafts))
    get_settings.cache_clear()

    monkeypatch.setattr(
        "app.services.commercial_export_service.ensure_order_priced",
        lambda *args, **kwargs: None,
    )

    draft_store = MagicMock(spec=DraftStore)
    fgs = file_generation_service or FileGenerationService()
    return CommercialExportService(draft_store=draft_store, file_generation_service=fgs)


def test_export_service_passes_append_batches_to_pdf_generator(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-402: wizard generate-files must forward metadata.append_batches to PDF."""
    batches = _same_type_multi_batches()
    order = _same_type_multi_order_without_line_batch_ids()
    fgs = MagicMock()

    def _write_pdf(**kwargs: Any) -> str:
        path = Path(kwargs["output_path"])
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"%PDF-FAKE")
        return str(path)

    fgs.generate_offer_pdf.side_effect = _write_pdf

    service = _make_export_service(tmp_path, monkeypatch, file_generation_service=fgs)
    service.generate_files(
        "draft402a",
        _export_service_payload(order, append_batches=batches),
        file_types=("pdf",),
    )

    fgs.generate_offer_pdf.assert_called_once()
    pdf_kwargs = fgs.generate_offer_pdf.call_args.kwargs
    assert "append_batches" in pdf_kwargs, (
        "CommercialExportService must pass append_batches into generate_offer_pdf"
    )
    assert pdf_kwargs["append_batches"] == batches


def test_export_service_passes_append_batches_to_xlsx_generator(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-402: wizard generate-files must forward metadata.append_batches to XLSX."""
    batches = _same_type_multi_batches()
    order = _same_type_multi_order_without_line_batch_ids()
    fgs = MagicMock()

    def _write_xlsx(**kwargs: Any) -> str:
        path = Path(kwargs["output_path"])
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"PK\x03\x04")
        return str(path)

    fgs.generate_offer_xlsx.side_effect = _write_xlsx

    service = _make_export_service(tmp_path, monkeypatch, file_generation_service=fgs)
    service.generate_files(
        "draft402b",
        _export_service_payload(order, append_batches=batches),
        file_types=("xlsx",),
    )

    fgs.generate_offer_xlsx.assert_called_once()
    xlsx_kwargs = fgs.generate_offer_xlsx.call_args.kwargs
    assert "append_batches" in xlsx_kwargs, (
        "CommercialExportService must pass append_batches into generate_offer_xlsx"
    )
    assert xlsx_kwargs["append_batches"] == batches


def test_export_service_same_type_multi_xlsx_unified_via_append_batches(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Without line append_batch_id, only forwarded append_batches flips to unified."""
    batches = _same_type_multi_batches()
    order = _same_type_multi_order_without_line_batch_ids()
    assert _require_is_unified()(order) is False
    assert _require_is_unified()(order, append_batches=batches) is True

    service = _make_export_service(tmp_path, monkeypatch)
    files = service.generate_files(
        "draft402c",
        _export_service_payload(order, append_batches=batches),
        file_types=("xlsx",),
    )
    assert files and files[0]["kind"] == "xlsx"
    xlsx_path = service.resolve_generated_file(files[0]["filename"])
    assert xlsx_path.exists()

    df = pd.read_excel(xlsx_path, sheet_name="КП", header=None)
    headers: list[str] = []
    for _, row in df.iterrows():
        vals = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if vals and vals[0] == "№":
            raw = [v if pd.notna(v) else "" for v in row.tolist()]
            while raw and raw[-1] == "":
                raw.pop()
            headers = [str(c).strip() if c != "" else "" for c in raw]
            break
    assert headers == UNIFIED_HEADERS, (
        "export generate-files must produce unified XLSX for same-type multi-append"
    )


def test_export_service_mixed_xlsx_includes_pb_only_delivery(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-402 / MNA-201: mixed export delivery row only from plates cargo."""
    order = [
        _plate(line_id="p1", qty=10, weight=2000.0, unit_price=100.0),
        _pile(line_id="s1"),
    ]
    service = _make_export_service(tmp_path, monkeypatch)
    files = service.generate_files(
        "draft402d",
        _export_service_payload(
            order,
            logistics_cost=5000.0,
            product_type="mixed",
            append_batches=[
                {"batch_id": "b1", "product_type": "plates", "line_ids": ["p1"]},
                {"batch_id": "b2", "product_type": "piles", "line_ids": ["s1"]},
            ],
        ),
        file_types=("xlsx",),
    )
    xlsx_path = service.resolve_generated_file(files[0]["filename"])
    df = pd.read_excel(xlsx_path, sheet_name="КП", header=None)
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
    assert header_idx is not None
    assert headers == UNIFIED_HEADERS
    body = df.iloc[header_idx + 1 :].copy()
    body.columns = range(len(body.columns))
    names = _body_column(body, headers, "Наименование")
    assert any("доставк" in n.lower() for n in names), names


def test_export_and_archive_agree_on_append_batches_kwarg_for_same_type_multi(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """MNA-402: wizard export and archive regen pass the same append_batches kwarg."""
    import asyncio

    from app.services.archive_service import ArchiveService

    batches = _same_type_multi_batches()
    order = _same_type_multi_order_without_line_batch_ids()

    export_fgs = MagicMock()
    captured: dict[str, Any] = {}

    def _capture_xlsx(**kwargs: Any) -> str:
        captured["export"] = kwargs.get("append_batches")
        path = Path(kwargs["output_path"])
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"PK\x03\x04")
        return str(path)

    export_fgs.generate_offer_xlsx.side_effect = _capture_xlsx
    export_service = _make_export_service(
        tmp_path / "export", monkeypatch, file_generation_service=export_fgs
    )
    export_service.generate_files(
        "draft402e",
        _export_service_payload(order, append_batches=batches),
        file_types=("xlsx",),
    )

    repository = MagicMock()
    repository.get_by_id.return_value = {
        "kp_id": 402,
        "creation_date": "12.08.2026",
        "customer_name": "ООО MNA402",
        "manager_name": "Менеджер",
        "discount_percent": 0.0,
        "logistics_cost": 0.0,
        "delivery_conditions": "",
        "payment_conditions": "",
        "owner_user_id": 1,
        "product_type": "plates",
        "plates": order,
        "append_batches": batches,
    }
    monkeypatch.setattr(
        "app.services.archive_service.order_data_from_kp_info",
        lambda _raw: order,
    )

    class FakeBuffer:
        def getvalue(self) -> bytes:
            return b"PK\x03\x04"

    fake_xlsx = MagicMock(return_value=FakeBuffer())

    def _capture_archive(*args: Any, **kwargs: Any) -> FakeBuffer:
        captured["archive"] = kwargs.get("append_batches")
        return FakeBuffer()

    fake_xlsx.side_effect = _capture_archive
    monkeypatch.setattr(
        "app.services.archive_service.generate_commercial_offer_xlsx",
        fake_xlsx,
    )

    archive_out = tmp_path / "archive"
    archive_out.mkdir()
    archive_service = ArchiveService(repository=repository, outputs_dir=archive_out)
    asyncio.run(
        archive_service.generate_document(402, "xlsx", user={"id": 1, "role": "admin"})
    )

    assert captured.get("export") == batches, (
        "export must forward append_batches; got "
        f"{captured.get('export')!r}"
    )
    assert captured.get("archive") == batches, (
        "archive regen must forward append_batches; got "
        f"{captured.get('archive')!r}"
    )
    assert captured["export"] == captured["archive"]
