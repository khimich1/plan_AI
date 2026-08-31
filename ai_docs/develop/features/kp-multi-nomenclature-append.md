# КП — несколько наименований (append loop)

**Status:** ✅ Implemented  
**Date:** 2026-08-12  
**Orchestration:** `orch-2026-08-12-14-05-kp-multi-append`  
**Report:** [2026-08-12-kp-multi-nomenclature-append-implementation.md](../reports/2026-08-12-kp-multi-nomenclature-append-implementation.md)  
**Spec:** [kp-multi-nomenclature-append.md](../../specs/kp-multi-nomenclature-append.md)

## Description

Одно коммерческое предложение собирается из нескольких заходов ввода (в т.ч. повтор номенклатуры). На result-шаге менеджер добавляет другое наименование, пропускает клиента со 2-го цикла, держит одну скидку, считает доставку только по весу плит и получает unified PDF/XLSX. Можно дописать уже сохранённое КП из архива (статус «в работе») на том же `kp_id`.

## How It Works

1. Result → «Добавить другое наименование» → picker → input → result (client skip).
2. Draft lines carry `line_id` + `product_type`; `metadata.append_batches` supports undo.
3. Save creates/updates one `kp_id`, splits lines across `kp_*` tables, sets `mixed` when >1 type.
4. Multi/append export uses unified columns; mono one-shot keeps legacy templates.
5. Archive shows one badge per product type; CTA gated to «в работе».

## Usage

- New KP: `/commercial-offer/new` → complete first cycle → «Добавить другое наименование».
- Resume: archive drawer (status «в работе») → «Добавить другое наименование».
- Undo last batch / delete line returns to result.

## API Endpoints (high level)

- `POST /api/v1/commercial/drafts/{id}/append/start` — begin append cycle
- Type PATCH with `mode=append` — merge lines without wiping other types
- `POST .../undo-last` / `DELETE .../lines/{line_id}` — undo / delete
- Archive hydrate / save update same `kp_id` when resume

## Components

- `CalculationResultStep` — Тип column, CTA, undo/delete, trip gate
- `CommercialOfferWizard` / `wizardDraftStore` — sticky append loop
- `ArchiveOfferList` / `OfferDetailsDrawer` — multi badges + resume CTA

## Known Issues

- AI/grade rebuild paths may wipe other types in draft `order_data`
- Same-type multi after DB may need persisted `append_batches` for unified regen
- Wizard progress titles cosmetic when client skipped

## Related Tasks

- MNA-001 … MNA-702 (see plan / report)
