"""Type table for commercial product drafts (Q1). No live ``if product_type`` branches."""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from typing import Any, Literal

from app.schemas.commercial import WizardStepId
from core.bridge_pile_price_db import list_available_grades
from core.commercial_offer_xlsx import DB_PATH
from core.fbs_price_db import list_available_grades as list_fbs_available_grades
from core.ocr_gpt import (
    apply_bridge_piles_with_ai,
    apply_fbs_with_ai,
    apply_marches_with_ai,
    apply_piles_with_ai,
    apply_steps_with_ai,
)

PreviewFn = Callable[..., Any]
MetadataFn = Callable[..., dict[str, Any]]
AiFn = Callable[..., Awaitable[Any]]
GradesFn = Callable[[str], list[str]]


@dataclass(frozen=True)
class ProductDraftSpec:
    product_type: str
    wizard_step: WizardStepId
    batches_key: str
    type_mismatch_update: str
    type_mismatch_ai: str | None
    type_mismatch_grades: str | None
    invalid_mode_error: str
    ai_empty_error: str
    grades_empty_error: str | None
    has_grades: bool
    needs_plate_ctx: bool
    allow_empty_create: bool
    empty_create_grade: str | None
    use_preview_order: bool
    update_reject_map: dict[str, str] | None
    ai_payload_style: Literal["rich", "compact"]
    grades_skip_unavailable: bool
    generate_preview: PreviewFn
    build_metadata: MetadataFn
    apply_ai: AiFn
    list_available_grades: GradesFn | None


def _preview_plates(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.commercial_service.generate_preview(text=text, plate_order_ctx=plate_order_ctx)


def _preview_piles(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.pile_service.generate_preview(text, db_path=str(DB_PATH))


def _preview_marches(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.march_service.generate_preview(text, db_path=str(DB_PATH))


def _preview_steps(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.step_service_product.generate_preview(text, db_path=str(DB_PATH))


def _preview_bridge_piles(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.bridge_pile_service.generate_preview(text, db_path=str(DB_PATH))


def _preview_fbs(wf: Any, text: str, *, plate_order_ctx: Any = None) -> Any:
    return wf.fbs_service.generate_preview(text, db_path=str(DB_PATH))


def _metadata_plates(draft_service: Any, **kwargs: Any) -> dict[str, Any]:
    return draft_service.build_preview_metadata(
        preview=kwargs["preview"],
        base_metadata=kwargs["base_metadata"],
        source_type=kwargs["source_type"],
        original_text=kwargs["original_text"],
        ocr_text=kwargs["ocr_text"],
        input_text=kwargs["input_text"],
        last_source_filename=kwargs["last_source_filename"],
        plate_batches=kwargs["batches"],
        wide_plates_resolved=kwargs["wide_plates_resolved"],
        source_metadata=kwargs["source_metadata"],
        owner_user_id=kwargs.get("owner_user_id"),
    )


def _typed_metadata(method_name: str, batches_key: str) -> MetadataFn:
    def build(draft_service: Any, **kwargs: Any) -> dict[str, Any]:
        return getattr(draft_service, method_name)(
            preview=kwargs["preview"],
            base_metadata=kwargs["base_metadata"],
            source_type=kwargs["source_type"],
            original_text=kwargs["original_text"],
            ocr_text=kwargs["ocr_text"],
            input_text=kwargs["input_text"],
            last_source_filename=kwargs["last_source_filename"],
            source_metadata=kwargs["source_metadata"],
            owner_user_id=kwargs.get("owner_user_id"),
            **{batches_key: kwargs["batches"]},
        )

    return build


async def _call_ai(
    fn: AiFn,
    text_kw: str,
    current_text: str,
    instruction: str,
    image_path: str | None,
) -> Any:
    kwargs: dict[str, Any] = {text_kw: current_text, "user_instruction": instruction}
    if image_path is not None:
        kwargs["image_path"] = image_path
    return await fn(**kwargs)


async def _ai_plates(current_text: str, instruction: str, image_path: str | None) -> Any:
    # Tests monkeypatch this name on commercial_workflow_service; resolve at call time.
    from app.services.commercial_workflow_service import apply_plates_with_ai as runtime_ai

    return await _call_ai(runtime_ai, "current_plates_text", current_text, instruction, image_path)


async def _ai_piles(current_text: str, instruction: str, image_path: str | None) -> Any:
    return await _call_ai(apply_piles_with_ai, "current_piles_text", current_text, instruction, image_path)


async def _ai_marches(current_text: str, instruction: str, image_path: str | None) -> Any:
    return await _call_ai(apply_marches_with_ai, "current_marches_text", current_text, instruction, image_path)


async def _ai_steps(current_text: str, instruction: str, image_path: str | None) -> Any:
    return await _call_ai(apply_steps_with_ai, "current_steps_text", current_text, instruction, image_path)


async def _ai_bridge_piles(current_text: str, instruction: str, image_path: str | None) -> Any:
    return await _call_ai(
        apply_bridge_piles_with_ai,
        "current_bridge_piles_text",
        current_text,
        instruction,
        image_path,
    )


async def _ai_fbs(current_text: str, instruction: str, image_path: str | None) -> Any:
    return await _call_ai(apply_fbs_with_ai, "current_fbs_text", current_text, instruction, image_path)


def _bridge_available_grades(mark: str) -> list[str]:
    return list(list_available_grades(mark, db_path=str(DB_PATH)))


def _fbs_available_grades(mark: str) -> list[str]:
    return list(list_fbs_available_grades(mark, db_path=str(DB_PATH)))


_PLATES_UPDATE_REJECT = {
    "piles": "Для КП на сваи используйте endpoint /piles.",
    "marches": "Для КП на лестничные марши используйте endpoint /marches.",
    "steps": "Для КП на ступени используйте endpoint /steps.",
}


SPECS: dict[str, ProductDraftSpec] = {
    "plates": ProductDraftSpec(
        product_type="plates",
        wizard_step=WizardStepId.plates,
        batches_key="plate_batches",
        type_mismatch_update="",
        type_mismatch_ai=None,
        type_mismatch_grades=None,
        invalid_mode_error="Некорректный режим обновления списка плит.",
        ai_empty_error="ИИ не смог обработать список плит. Попробуйте уточнить инструкцию.",
        grades_empty_error=None,
        has_grades=False,
        needs_plate_ctx=True,
        allow_empty_create=False,
        empty_create_grade=None,
        use_preview_order=True,
        update_reject_map=_PLATES_UPDATE_REJECT,
        ai_payload_style="rich",
        grades_skip_unavailable=False,
        generate_preview=_preview_plates,
        build_metadata=_metadata_plates,
        apply_ai=_ai_plates,
        list_available_grades=None,
    ),
    "piles": ProductDraftSpec(
        product_type="piles",
        wizard_step=WizardStepId.piles,
        batches_key="pile_batches",
        type_mismatch_update="Черновик не является КП на сваи.",
        type_mismatch_ai="ИИ-редактирование свай доступно только для КП на сваи.",
        type_mismatch_grades="Черновик не является КП на сваи.",
        invalid_mode_error="Некорректный режим обновления списка свай.",
        ai_empty_error="ИИ не смог обработать список свай. Попробуйте уточнить инструкцию.",
        grades_empty_error="Список свай пустой.",
        has_grades=True,
        needs_plate_ctx=False,
        allow_empty_create=True,
        empty_create_grade="B25",
        use_preview_order=False,
        update_reject_map=None,
        ai_payload_style="rich",
        grades_skip_unavailable=False,
        generate_preview=_preview_piles,
        build_metadata=_typed_metadata("build_pile_preview_metadata", "pile_batches"),
        apply_ai=_ai_piles,
        list_available_grades=None,
    ),
    "marches": ProductDraftSpec(
        product_type="marches",
        wizard_step=WizardStepId.marches,
        batches_key="march_batches",
        type_mismatch_update="Черновик не является КП на лестничные марши.",
        type_mismatch_ai="ИИ-редактирование маршей доступно только для КП на лестничные марши.",
        type_mismatch_grades="Черновик не является КП на лестничные марши.",
        invalid_mode_error="Некорректный режим обновления списка маршей.",
        ai_empty_error="ИИ не смог обработать список маршей. Попробуйте уточнить инструкцию.",
        grades_empty_error="Список маршей пустой.",
        has_grades=True,
        needs_plate_ctx=False,
        allow_empty_create=True,
        empty_create_grade="B25",
        use_preview_order=False,
        update_reject_map=None,
        ai_payload_style="rich",
        grades_skip_unavailable=False,
        generate_preview=_preview_marches,
        build_metadata=_typed_metadata("build_march_preview_metadata", "march_batches"),
        apply_ai=_ai_marches,
        list_available_grades=None,
    ),
    "steps": ProductDraftSpec(
        product_type="steps",
        wizard_step=WizardStepId.steps,
        batches_key="step_batches",
        type_mismatch_update="Черновик не является КП на ступени.",
        type_mismatch_ai="ИИ-редактирование ступеней доступно только для КП на ступени.",
        type_mismatch_grades=None,
        invalid_mode_error="Некорректный режим обновления списка ступеней.",
        ai_empty_error="ИИ не смог обработать список ступеней. Попробуйте уточнить инструкцию.",
        grades_empty_error=None,
        has_grades=False,
        needs_plate_ctx=False,
        allow_empty_create=True,
        empty_create_grade=None,
        use_preview_order=False,
        update_reject_map=None,
        ai_payload_style="rich",
        grades_skip_unavailable=False,
        generate_preview=_preview_steps,
        build_metadata=_typed_metadata("build_step_preview_metadata", "step_batches"),
        apply_ai=_ai_steps,
        list_available_grades=None,
    ),
    "bridge_piles": ProductDraftSpec(
        product_type="bridge_piles",
        wizard_step=WizardStepId.bridge_piles,
        batches_key="bridge_pile_batches",
        type_mismatch_update="Черновик не является КП на мостовые сваи.",
        type_mismatch_ai="ИИ-редактирование доступно только для КП на мостовые сваи.",
        type_mismatch_grades="Черновик не является КП на мостовые сваи.",
        invalid_mode_error="Некорректный режим обновления списка мостовых свай.",
        ai_empty_error="ИИ не вернул распознанный список мостовых свай.",
        grades_empty_error="Список мостовых свай пустой.",
        has_grades=True,
        needs_plate_ctx=False,
        allow_empty_create=True,
        empty_create_grade="B25",
        use_preview_order=False,
        update_reject_map=None,
        ai_payload_style="compact",
        grades_skip_unavailable=True,
        generate_preview=_preview_bridge_piles,
        build_metadata=_typed_metadata("build_bridge_pile_preview_metadata", "bridge_pile_batches"),
        apply_ai=_ai_bridge_piles,
        list_available_grades=_bridge_available_grades,
    ),
    "fbs": ProductDraftSpec(
        product_type="fbs",
        wizard_step=WizardStepId.fbs,
        batches_key="fbs_batches",
        type_mismatch_update="Черновик не является КП на ФБС.",
        type_mismatch_ai="ИИ-редактирование доступно только для КП на ФБС.",
        type_mismatch_grades="Черновик не является КП на ФБС.",
        invalid_mode_error="Некорректный режим обновления списка ФБС.",
        ai_empty_error="ИИ не вернул распознанный список ФБС.",
        grades_empty_error="Список ФБС пустой.",
        has_grades=True,
        needs_plate_ctx=False,
        allow_empty_create=True,
        empty_create_grade="B25",
        use_preview_order=False,
        update_reject_map=None,
        ai_payload_style="compact",
        grades_skip_unavailable=True,
        generate_preview=_preview_fbs,
        build_metadata=_typed_metadata("build_fbs_preview_metadata", "fbs_batches"),
        apply_ai=_ai_fbs,
        list_available_grades=_fbs_available_grades,
    ),
}


def get_spec(product_type: str) -> ProductDraftSpec:
    spec = SPECS.get(product_type)
    if spec is None:
        raise ValueError("Некорректный тип продукта.")
    return spec
