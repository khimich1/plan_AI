# Plan: Надёжное OCR свай (GigaChat) + optional plate speed

> **Spec:** [`ai_docs/specs/pile-ocr-reliability.md`](../../specs/pile-ocr-reliability.md)  
> **Дата:** 2026-07-31  
> **Статус:** IMPLEMENT ✅ (OCR-501) — manual smoke OCR-502 pending  
> **D12 (plate speed):** **P0 defer** — plate OCR не меняем

---

## Implementation order

```
1. Normalizer (TDD)          ← no API, fast feedback
2. Pile parser gate + verify policy
3. GigaChat extract_piles + run_pile_ocr_pipeline
4. Wire CommercialDraftService (product_type)
5. Fixtures (PNG + mock JSON)
6. Integration tests
7. Prompt enrichment
8. ~~Plate speed~~ — out of scope (D12 P0)
```

---

## Tasks

### Phase A — Domain (no external API)

- [ ] **OCR-101:** Normalizer rules R1–R4 в `pile_text_normalizer.py`
  - Acceptance: pilot strings → `С90.30-11 189` etc.; AC-1 green
  - Verify: `pytest tests/test_pile_ocr_normalizer.py -q`
  - Files: `core/pile_text_normalizer.py`, `tests/test_pile_ocr_normalizer.py`

- [ ] **OCR-102:** `apply_pile_parser_gate` в `core/ocr/pile_parser_gate.py`
  - Acceptance: parsed pile → unchanged; bad line → `parser_rejected`, cap confidence
  - Verify: extend `tests/test_ocr_parser_gate.py` or new file
  - Files: `core/ocr/pile_parser_gate.py`, tests

- [ ] **OCR-103:** `should_run_pile_verify` (extend `verify_policy.py` or sibling)
  - Acceptance: 3 good pile rows + small image → skip; low conf → run
  - Verify: `pytest tests/test_ocr_verify_policy.py -q`
  - Files: `core/ocr/verify_policy.py` or `pile_verify_policy.py`, tests

### Phase B — GigaChat pipeline

- [ ] **OCR-201:** `_sync_extract_piles` in `gigachat.py`
  - Acceptance: uses `build_pile_parser_system_prompt()`; returns JSON list
  - Verify: unit test with mocked GigaChat client
  - Files: `core/ocr/providers/gigachat.py`, `tests/test_ocr_gigachat_provider.py`

- [ ] **OCR-202:** `run_pile_ocr_pipeline` in `pipeline.py`
  - Acceptance: extract → pile parser gate → auto verify → `build_result_payload`
  - Verify: `pytest tests/test_pile_ocr_pipeline.py -q` (mock provider)
  - Files: `core/ocr/pipeline.py`, `tests/test_pile_ocr_pipeline.py`

- [ ] **OCR-203:** `recognize_text_smart(..., product_type="piles")` entry point
  - Acceptance: piles branch calls `run_pile_ocr_pipeline`; default plates unchanged
  - Verify: mock branch test
  - Files: `core/ocr/recognition.py`

### Phase C — Service wiring

- [ ] **OCR-301:** `CommercialDraftService.extract_text_from_image(product_type=...)`
  - Acceptance: pile paths pass `product_type="piles"`
  - Verify: mock `recognize_text_smart` asserts product_type
  - Files: `commercial_draft_service.py`, tests

- [ ] **OCR-302:** `resolve_source_input` + workflow pile create/update
  - Acceptance: `_create_pile_draft` / `update_draft_piles` use pile OCR
  - Verify: AC-2 integration test
  - Files: `commercial_workflow_service.py`, `test_commercial_pile_flow.py`

- [ ] **OCR-303:** OCR repair in preview path
  - Acceptance: `normalize_pile_order_text` applies R1–R4 after OCR text assembly
  - Verify: AC-1 end-to-end via `CommercialPileService.generate_preview`

### Phase D — Fixtures & prompt

- [ ] **OCR-401:** Fixtures `tests/fixtures/pile_ocr/`
  - Acceptance: `pilot_table.png` (from pilot screenshot), `lines.txt`, `gigachat_extract_response.json`
  - Verify: files committed; CI reads them
  - Files: fixtures dir, copy from assets

- [ ] **OCR-402:** Enrich `pile_format_prompt.py` (AC-7)
  - Acceptance: examples «Сваи С90.30-11», «189 шт», «13и» suffix
  - Verify: string assert in test

### Phase E — Regression

- [ ] **OCR-501:** Full test suite
  - Verify: `pytest tests/test_commercial_pile_flow.py tests/test_commercial_web_flow.py tests/test_ocr_verify_policy.py -q`
  - AC-3, AC-6, AC-9

- [ ] **OCR-502:** Manual smoke checklist
  - Upload pilot PNG in wizard → 3 rows + B25 prices
  - Document in report

### Phase F — Out of scope (D12 P0)

Plate OCR speed — отложено. AC-10 = N/A (no code change).

---

## Verification checkpoints

| After | Command | Expected |
|-------|---------|----------|
| Phase A | `pytest tests/test_pile_ocr_normalizer.py tests/test_ocr_verify_policy.py -q` | green |
| Phase B | `pytest tests/test_pile_ocr_pipeline.py -q` | green |
| Phase C | `pytest tests/test_commercial_pile_flow.py -q` | AC-2 green |
| Phase E | full commercial + ocr tests | all green |

---

## Risks during implement

1. GigaChat provider Protocol only has `extract_plates` — extend with `extract_piles` or generic `product_type` param
2. Verify reuse — ensure pile JSON fits existing verify prompt token limits
3. PNG fixture size — keep <500KB if possible for CI speed

---

## Parallel work

| Can parallel | Must be sequential |
|--------------|-------------------|
| OCR-101 + OCR-102 + OCR-103 | OCR-202 after OCR-201 |
| OCR-402 prompt | OCR-301 after OCR-203 |
| OCR-401 fixtures | OCR-302 after OCR-301 |
