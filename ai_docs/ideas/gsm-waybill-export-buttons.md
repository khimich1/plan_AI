# ГСМ: кнопки завершения цикла ПЛ (экспорт zip)

Дата: 2026-08-17. Статус: implemented (UI export zip).
Report: [`../develop/reports/2026-08-17-gsm-waybill-export-buttons.md`](../develop/reports/2026-08-17-gsm-waybill-export-buttons.md).
Родитель: [`gsm-module-putevye-listy.md`](gsm-module-putevye-listy.md).
UI: `frontend/src/features/gsm/components/GsmPeriodView.tsx`.
API: `POST /api/v1/gsm/waybills/export` (`GsmExportService`).

## Problem Statement

Как бухгалтеру или админу за один сеанс в `/gsm` пройти путь
«импорт → генерация → правки → готовые бланки», не выходя в API/curl?

## Контекст: где обрывается цикл

| Действие | UI | API |
|----------|----|-----|
| Автогенерация периода | ✅ «Сгенерировать» | ✅ |
| Ручной ПЛ | ✅ «Ручной ПЛ» | ✅ |
| Правка дня | ✅ drawer «Сохранить» | ✅ PATCH |
| Подтверждение | ❌ | ✅ `POST /waybills/{id}/confirm` |
| Экспорт zip | ❌ | ✅ `POST /waybills/export` |

Backend экспорта готов: zip «ПЛ DD.MM.YY.xls» через шаблон + openpyxl +
LibreOffice `soffice`. После успеха дни переводятся в `exported`.
`get_last_confirmed_waybill` учитывает и `confirmed`, и `exported` — для
цепочки одометра/бака на следующий период отдельный confirm не обязателен.

## Recommended Direction

**Export-first MVP** с минимальным diff: одна primary-кнопка **«Экспорт zip»**
в строке формы (`GsmPeriodView`, рядом с «Сгенерировать» / «Ручной ПЛ»).

Подключить `gsmApi.exportWaybills` → blob download + mutation.
Отдельный confirm в MVP **не включаем** — финальное действие сеанса =
скачивание zip; статус `exported` достаточен для следующей генерации.

### Двухуровневый gate перед экспортом

Warnings в коде уже разделены по смыслу (`core/gsm/generator.py`,
`frontend/src/features/gsm/lib/waybillWarnings.ts`):

| Код | Уровень | Поведение при экспорте |
|-----|---------|------------------------|
| `manual_intervention` | жёсткий | Кнопка disabled; день в списке «исправьте» |
| `unsolvable` (period) | жёсткий | Кнопка disabled |
| `balance_route` | мягкий | Confirm-диалог «есть предупреждения…» |
| `hook_above_threshold` | мягкий | Confirm-диалог |
| `weekend_anchor` | мягкий | Confirm-диалог |

**Жёсткий стоп:** любой день с `manual_intervention` (`isProblematicDay`) или
period warning `unsolvable` после генерации.

**Мягкий confirm:** остальные day-level warnings — кнопка активна, но
диалог перед скачиванием. Жёлтые `balance_route` часто нормальный компромисс
солвера; hard block заставил бы править то, что система уже считает приемлемым.

### Прочие UX-решения

- **Пустой период:** «Экспорт» disabled, пока `useGsmWaybillsQuery` не вернул
  хотя бы один ПЛ за выбранный период.
- **Повторный экспорт:** confirm-диалог, если в периоде есть дни со статусом
  `exported` («Период уже экспортировался, скачать снова?»).
- **Ошибка LibreOffice:** Alert через `formatGsmError` + подсказка админу
  («нужен LibreOffice / soffice на сервере»). Коды: `gsm_export_soffice_missing`,
  `gsm_export_soffice_timeout`, `gsm_export_soffice_failed`.

## Key Assumptions to Validate

- [ ] `soffice` доступен в среде бухгалтера — пробный экспорт одного дня
- [ ] Экспорт draft без отдельного confirm — приемлемо для бухгалтерии
- [ ] Двухуровневый gate (hard manual/unsolvable + soft confirm на жёлтые) —
  достаточная защита от битого комплекта
- [ ] Экспорт одной машины за раз (`vehicle_ids: [current]`) покрывает
  основной сценарий

## MVP Scope

**Frontend**

- `gsmApi.exportWaybills(payload)` — POST + blob response + download helper
- `useExportGsmWaybillsMutation` в `useGsmQueries.ts`
- helpers: `canExportHardBlock`, `hasSoftExportWarnings` (waybills +
  periodWarnings)
- `GsmPeriodView`: кнопка «Экспорт zip», disabled/tooltip, soft confirm,
  reexport confirm
- расширить `formatGsmError` / `gsmErrors.ts` для export-кодов

**Тесты**

- `gsmApi.test.ts` — POST export, blob handling
- `GsmPeriodView.test.tsx` — disabled при manual/empty, soft confirm, reexport

**Backend**

- без изменений (API уже есть)

## Not Doing (and Why)

- **Bulk confirm периода** — нет bulk API; N+1 запросов некрасиво; `exported`
  закрывает цепочку
- **«Подтвердить день» в drawer** — фаза 1.1, если понадобится аудит-трейл
- **Экспорт нескольких машин** — не нужен для первого сеанса; API принимает
  `vehicle_ids[]`, UI — только текущая машина
- **Auto-confirm перед export** — лишняя сложность и риск частичного state
  при падении export
- **Новый wizard / экран** — против ограничения small diff
- **Health-check endpoint для soffice** — достаточно alert с подсказкой при
  первой ошибке

## Open Questions

- Нужен ли inline CTA «Скачать бланки» в зелёном алерте после генерации
  (дублирует кнопку в форме)?
- Блокировать ли экспорт при period-level warnings кроме `unsolvable`
  (например `weekend_anchor` только на period, не на day)?
- Когда делать drawer confirm — по запросу бухгалтера после sign-off MVP?

## Связанные артефакты

- Feature doc: [`../develop/features/gsm-module-putevye-listy.md`](../develop/features/gsm-module-putevye-listy.md)
- Deploy note (LibreOffice): implementation report GSM module
