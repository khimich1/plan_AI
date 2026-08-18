# Report: ГСМ — кнопка «Экспорт zip»

**Date:** 2026-08-17  
**Status:** ✅ Completed  
**Idea:** [`../../ideas/gsm-waybill-export-buttons.md`](../../ideas/gsm-waybill-export-buttons.md)  
**Feature:** [`../features/gsm-module-putevye-listy.md`](../features/gsm-module-putevye-listy.md)

## Summary

Закрыт последний шаг цикла ПЛ в UI: бухгалтер скачивает zip бланков с вкладки **Период**, без curl/API. Backend `POST /gsm/waybills/export` не менялся. Отдельный confirm дня в MVP нет — статус `exported` уже участвует в цепочке бака/одометра.

## What changed

- `httpClient.download` принимает POST (CSRF + blob); `gsmApi.exportWaybills`.
- `GsmPeriodView`: кнопка **«Экспорт zip»** в строке формы.
- Двухуровневый gate (`exportGate.ts`): hard-block `manual_intervention` / `unsolvable`; soft confirm на `balance_route`, `hook_above_threshold`, `weekend_anchor`; confirm повторного экспорта.
- Пустой период — кнопка disabled. Ошибки `gsm_export_soffice_*` — русский Alert с подсказкой про LibreOffice.

## Tests

```bash
cd frontend && npx vitest run --pool=threads src/features/gsm src/shared/api/httpClient.test.ts
```

79 passed (17 files), including `exportGate.test.ts`, `GsmPeriodView.test.tsx`, `gsmApi.test.ts`, `httpClient.test.ts`.

## Not in this slice

- Confirm дня в drawer / bulk confirm
- Экспорт нескольких машин из UI
- Health-check `soffice` до клика
