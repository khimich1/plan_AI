# Report: КП — несколько наименований (append loop)

**Date:** 2026-08-12  
**Orchestration:** `orch-2026-08-12-14-05-kp-multi-append`  
**Status:** ✅ Completed  
**Spec:** [`ai_docs/specs/kp-multi-nomenclature-append.md`](../../specs/kp-multi-nomenclature-append.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md`](../plans/2026-08-12-kp-multi-nomenclature-append.md)  
**Idea:** [`ai_docs/ideas/kp-multi-nomenclature-append.md`](../../ideas/kp-multi-nomenclature-append.md)

## Summary

Менеджер собирает **одно КП** из нескольких заходов (в т.ч. повтор типа), со sticky клиентом/скидкой, логистикой только по весу ПБ, unified PDF/XLSX при multi/append, и может **дописать уже сохранённое КП** из архива (Q1=C, тот же `kp_id`, только статус «в работе»). Все **20 задач** MNA-001…MNA-702 выполнены; automated release gate (CP-7 Full Verify) зелёный.

## What Was Built

### Phase 0 — Domain helpers
- **MNA-001:** `total_order_cargo_weight_kg(..., product_types={"plates"})` — рейсы только от плит.
- **MNA-002:** `format_line_name` — марка + класс бетона в скобках для unified export.

### Phase 1 — Draft model + append API
- **MNA-101:** схема `line_id` / `product_type` / `append_batch_id` на строках; `metadata.append_batches`.
- **MNA-102:** stamp identity при построении `order_data`.
- **MNA-103:** `POST .../append/start`, `mode=append` merge, undo last batch, `DELETE .../lines/{line_id}`.
- **MNA-104:** skip client со 2-го цикла (BE + FE `wizardStepOrder`).

### Phase 2 — Calculation + PB logistics
- **MNA-201:** pricing/calc используют plates-only cargo.
- **MNA-202:** mixed calculate + одна скидка на все строки; wide-plate gate по `order_data`.

### Phase 3 — Persistence
- **MNA-301:** `line_id` на line-таблицах; `kp_meta.product_type = mixed`.
- **MNA-302:** multi-table create, сквозной `position_number`.
- **MNA-303:** read merge sort by `position_number` + `product_type` на строках.
- **MNA-304:** update sync by `line_id`; status gate «в работе»; preserve plate DB ids.

### Phase 4 — Export
- **MNA-401:** unified `№ | Тип | Наименование | Кол-во | Цена | Сумма`; mono R3 без регрессии.
- **MNA-402:** wizard export ↔ archive regen согласованы; `append_batches` пробрасывается в генераторы.

### Phase 5 — Wizard UX
- **MNA-501:** колонка «Тип», CTA «Добавить другое наименование», undo/delete, trip field gated.
- **MNA-502:** sticky header loop (`start-append-cycle`), skip client, full reset «Создать новое КП».

### Phase 6 — Archive resume C
- **MNA-601:** hydrate draft из сохранённого КП (`saved_offer.kp_id` / `resume_kp_id`).
- **MNA-602:** multi badges `product_types`, contains-type filter, CTA только «в работе».

### Phase 7 — Production + E2E
- **MNA-701:** production candidates включают `mixed` с плитами; в работу идут только `kp_plates`.
- **MNA-702:** E2E suite SC-1…SC-9 + mono hydrate/`saved_offer` retention fixes.

## Success Criteria (SC-1…SC-9)

| SC | Criterion | Status | Evidence |
|----|-----------|--------|----------|
| SC-1 | Плиты→Сваи→Плиты → один `kp_id` | ✅ | `test_sc1_*` in `tests/test_commercial_multi_append_flow.py` |
| SC-2 | Unified PDF/XLSX: порядок, тип, grade в имени, одна скидка | ✅ | `test_sc2_*` + export mixed suite |
| SC-3 | Skip client со 2-го захода | ✅ | `test_sc3_*` + wizard step service / FE order |
| SC-4 | Undo/delete по `line_id` / batch | ✅ | `test_sc4_*` + draft append API |
| SC-5 | Delivery только от веса plates; без plates — без доставки | ✅ | `test_sc5_*` ×2 |
| SC-6 | Append к сохранённому КП из архива (тот же `kp_id`, «в работе») | ✅ | `test_sc6_*` ×2 + hydrate fixes |
| SC-7 | Несколько бейджей типов в архиве | ✅ | `test_sc7_*` + archive UI |
| SC-8 | Mono без append — без регрессии | ✅ | `test_sc8_*` + R3 export headers |
| SC-9 | Tests green; production только plates (mixed-with-plates OK) | ✅ | gate below + `test_production_mixed_inclusion.py` |

Manual smoke (archive → append → PDF trips) — **unchecked** in plan CP-7; recommended before prod announce.

## Completed Tasks

| Task | Name | Phase | Tests (task run) | Completed |
|------|------|-------|------------------|-----------|
| MNA-001 | Plates-only cargo weight helper | 0 | 12/12 | 14:26 |
| MNA-002 | format_line_name (grade in name) | 0 | 14/14 | 14:28 |
| MNA-101 | Schema line fields + append_batches | 1 | 14/14 | 14:41 |
| MNA-102 | Stamp line_id + product_type | 1 | 13/13 | 15:08 |
| MNA-103 | Append / undo / delete API | 1 | 30/30 | 15:35 |
| MNA-104 | Skip client on cycle ≥2 | 1 | 27/27 | 15:00 |
| MNA-201 | PB-only pricing/calculation wiring | 2 | 24/24 | 15:00 |
| MNA-202 | Mixed calculate + shared discount | 2 | 41/41 | 15:50 |
| MNA-301 | DB migration line_id + mixed meta | 3 | 10/10 | 14:37 |
| MNA-302 | Multi-table create persistence | 3 | 12/12 | 15:52 |
| MNA-303 | Read merge by position_number | 3 | 16/16 | 16:10 |
| MNA-304 | Update kp_id sync + status gate | 3 | 29/29 | 16:32 |
| MNA-401 | Unified export + mono regression | 4 | 24/24 | 16:10 |
| MNA-402 | Export service + archive regen | 4 | 37/37 | 16:32 |
| MNA-501 | Result UI Тип + CTA + undo/delete | 5 | 16/16 | 16:50 |
| MNA-502 | Wizard loop sticky header | 5 | 11/11 | 17:19 |
| MNA-601 | Hydrate draft from saved KP | 6 | 15/15 | 16:50 |
| MNA-602 | Archive multi badges + CTA | 6 | 48/48 | 17:33 |
| MNA-701 | Production mixed-with-plates | 7 | 5/5 | 16:50 |
| MNA-702 | E2E flow + regression gate | 7 | 112/112 (+ FE 121) | 18:05 |

## Key files changed

**Domain / core:**  
`core/cargo_delivery_pricing.py`, `core/commercial_line_format.py`, `core/commercial_pricing.py`, `core/commercial_offer.py`, `core/commercial_offer_xlsx.py`, `core/commercial_offer_layout.py`, `core/kp_db_schema.py`, `core/kp_persistence_service.py`, `core/kp/offers_read.py`, `core/kp/offers_write.py`, `core/kp_order_data.py`

**App / API:**  
`app/schemas/commercial.py`, `app/schemas/archive.py`, `app/api/v1/endpoints/commercial.py`, `app/api/v1/endpoints/archive.py`, `app/services/commercial_{draft,workflow,wizard_step,calculation,export}_service.py`, `app/services/archive_service.py`, `app/repositories/kp_repository.py`

**Frontend:**  
`CalculationResultStep`, `CommercialOfferWizard`, `WizardProgress`, `wizardDraftStore`, `wizardStepOrder`, `commercialOfferApi`, archive list/drawer/page/API

**Tests (new/extended):**  
`tests/test_commercial_multi_append_flow.py`, `tests/test_commercial_draft_append.py`, `tests/test_commercial_export_mixed.py`, `tests/test_kp_persistence_mixed.py`, `tests/test_production_mixed_inclusion.py`, FE commercial-offer / commercial-archive vitest

## Test metrics (release gate MNA-702)

| Suite | Result |
|-------|--------|
| Core pytest (`multi_append_flow` + `draft_append` + `export_mixed` + `kp_persistence_mixed`) | **112 passed** |
| Frontend typecheck | ✅ |
| Frontend vitest (`commercial-offer` + `commercial-archive`) | **121 passed** / 23 files |
| Frontend build | ✅ |
| Broader `pytest -k "commercial or wizard or kp_persistence or archive"` | Core green; sandbox-only `sqlite3`/`EROFS` noise excluded from gate |

Per-task verify runs (overlapping suites) summed to **510/510** across the orchestration; final gate numbers above are authoritative.

**Orchestration window:** ~14:05 → ~18:06 (+03:00), ~4h.

## Technical decisions

- **Draft-first identity:** every line has `line_id` + `product_type`; `append_batches` is draft/undo metadata (not PDF segments).
- **PB-only logistics (Q2):** cargo weight filtered to plates; trip UI active iff ≥1 plate line.
- **Unified vs mono (R3):** multi / post-append → unified columns; mono one-shot keeps legacy templates.
- **Archive C (Q1):** same `kp_id`, regenerate files overwrite (R1 — no version history).
- **Status gate (R2):** append/update only when `kp_meta.status = «в работе»`.
- **Persistence:** create multi-table insert; update sync-by-`line_id` (preserve plate production fields).
- **Production:** `plates|mixed` candidates; only `kp_plates` rows enter plans.

## Known residuals / follow-ups

1. **AI / grade paths can wipe other types in `order_data`** — `apply_ai_*` / `update_draft_*_grades` still assign `order_data=preview.order_data` without the append compose merge. Text/OCR `mode=append` paths are covered; AI/grade rebuilds remain a risk on multi drafts.
2. **Same-type multi after DB round-trip** — `append_batches` / `append_batch_id` are draft-only today; archive regen of same-type multi may fall back to mono layout unless batches are mocked or persisted on KP raw. Mixed multi-type still unifies via distinct `product_type` on lines.
3. **WizardProgress titles** — when client is skipped, labels can still show `"3. Результат"` (cosmetic numbering).
4. **Manual smoke** — live archive «в работе» → append → download PDF → verify trips from plates only (CP-7 unchecked).
5. **Out of scope (locked):** non-plates cargo weight, PDF version history, append outside «в работе», production for non-plates, bot path.

## Related documentation

- Spec: [`ai_docs/specs/kp-multi-nomenclature-append.md`](../../specs/kp-multi-nomenclature-append.md)
- Plan: [`ai_docs/develop/plans/2026-08-12-kp-multi-nomenclature-append.md`](../plans/2026-08-12-kp-multi-nomenclature-append.md)
- Feature: [`ai_docs/develop/features/kp-multi-nomenclature-append.md`](../features/kp-multi-nomenclature-append.md)
- Handoff: [`ai_docs/develop/handoffs/2026-08-12-kp-multi-nomenclature-append.md`](../handoffs/2026-08-12-kp-multi-nomenclature-append.md)

## Next steps

1. Manual browser smoke on live stack (`./run+logs.sh`).
2. Optionally persist `append_batches` (or equivalent layout signal) for same-type multi after save.
3. Route AI/grade rebuilds through append-safe compose.
4. Renumber `WizardProgress` titles when `skipClient`.

## Status

**Implemented** (2026-08-12). Plan tasks MNA-001…MNA-702 complete for MVP scope. Workspace archived to `.cursor/workspace/completed/orch-2026-08-12-14-05-kp-multi-append/`.
