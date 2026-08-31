# Report: КП на мостовые сваи (bridge piles) MVP

**Date:** 2026-08-05  
**Spec:** [`ai_docs/specs/kp-bridge-piles.md`](../../specs/kp-bridge-piles.md)  
**Plan:** [`ai_docs/develop/plans/2026-08-05-kp-bridge-piles.md`](../plans/2026-08-05-kp-bridge-piles.md)

## Summary

MVP «КП на мостовые сваи» реализован end-to-end: picker «Мостовые сваи» → text/OCR/AI → после «Список верен» preview с классом бетона (только доступные B25/B30) → client → PDF/XLSX → save в `kp_bridge_piles`, архив с бейджем/фильтром. Цены: `bridge_pile_prices` в `pb.db` (импорт только лист «Прайс»). Марка в КП/PDF — как ввёл менеджер; алиасы T/В — только для lookup. Bulk «класс ко всем» — skip + warning (решение A). Production whitelist только `plates`.

Шаблоны: **piles** (grade UX + PDF), **marches/steps** (multi-product wiring + hide preview until confirm).

## AC checklist

| AC | Status | Tests / verify |
|----|--------|----------------|
| AC-1 picker «Мостовые сваи» → `bridge_piles` | ✅ | `ProductTypePicker` + vitest |
| AC-2 input text/photo/AI | ✅ | `BridgePileInputStep`, OCR bridge piles |
| AC-3 parser mark+grade+qty; merge; keep spelling | ✅ | `test_bridge_pile_line_parser.py` |
| AC-4 preview available grades only; no variant picker | ✅ | `KpBridgePilePreviewPanel` |
| AC-5 CLI import лист «Прайс»; aliases; skip zeros | ✅ | 114 price rows (synonym expansion); import tests |
| AC-6 missing mark+grade → block | ✅ | pricing + flow calculate |
| AC-7 client/result; pdf+xlsx only | ✅ | export non-plate file types |
| AC-8 PDF/XLSX: mark as typed, grade, qty, price | ✅ | commercial_offer / xlsx bridge branch |
| AC-9 save `kp_bridge_piles` + meta | ✅ | persistence + flow test |
| AC-10 archive badge + filter | ✅ | archive UI + types |
| AC-11 plate/pile/step/march regression | ✅ | focused pytest 83 passed |
| AC-12 not in production | ✅ | `test_production_bridge_pile_exclusion.py` |
| AC-13 shared `kp_id` | ✅ | shared `KP_offers` |
| AC-14–17 mirror piles | ✅ | grades bulk skip+warning, drawer, regen |

## Price import

```bash
python scripts/import_bridge_pile_prices_from_xlsx.py \
  "банк знаний/Прайс на мостовые сваи от 03.08.2026.xlsx" --sheet Прайс
# → 114 строк в bridge_pile_prices (алиасы развернуты; нули пропущены)
```

Grades: только `B25` / `B30`. Lookup нормализует C↔С, B↔В, T↔Т.

## Key files

**Backend:** `core/bridge_pile_*`, `scripts/import_bridge_pile_prices_from_xlsx.py`, `app/services/commercial_bridge_pile_service.py`, endpoints `.../bridge-piles`, OCR gate/pipeline.

**Frontend:** `ProductTypePicker`, `BridgePileInputStep`, `KpBridgePilePreviewPanel`, wizard/API/archive wiring.

**Tests:** `tests/test_bridge_pile_*`, `tests/test_commercial_bridge_pile_*`, `tests/test_production_bridge_pile_exclusion.py`, frontend vitest.

## Verification run

```bash
pytest tests/test_bridge_pile_*.py tests/test_commercial_bridge_pile_*.py \
  tests/test_production_bridge_pile_exclusion.py \
  tests/test_commercial_{pile,march,step}_flow.py \
  tests/test_commercial_wizard_step_service.py -q
# → 83 passed

cd frontend && npm run typecheck && npm run test -- --run && npm run build
# → green
```

## Smoke test

1. Import prices (CLI above).
2. UI: picker «Мостовые сваи» → `C8-35T1 2` + `C8-35В4 1` → «Список верен».
3. Bulk B25 on mix with `C13-40T3` → warning, B30-only row skipped.
4. Client → calculate → PDF shows typed marks → archive badge «Мостовые сваи».

## Remaining / next

- **ФБС** — next product iteration (see handoff).
- Manual browser smoke on live stack recommended after restart `./run+logs.sh`.
- No generic multi-product framework (by design).

## Status

**Implemented** (2026-08-05). Plan tasks BP-001…BP-604 complete for MVP scope.
