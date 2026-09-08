# Конструктор КП — фикс-пак 1 (по итогам ревью 2026-09-07)

Дата: 2026-09-07. Основание: полное ревью модуля `frontend/src/features/commercial-offer` (5 осей), пункт плана «1. Bugfix».

## Что исправлено

1. **`isProceeding` проведён по-настоящему** — во всех 6 шагах визарда вместо хардкода
   `isProceeding={false}` теперь передаётся `calculateMutation.isPending`.
   Эффект: кнопка «Готово, далее» блокируется и показывает «Переход...» во время расчёта,
   убрана возможность двойной отправки calculate.
   Файл: `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx` (6 мест).

2. **Bounds-check в `handleLineGradeChange`** — ветки FBS и bridge_piles приведены к остальным:
   добавлена проверка `lineIndex < 0 || lineIndex >= rows.length` перед обращением к строке.
   Раньше вне-диапазонный индекс молча инициировал re-ingest списка без изменений.
   Файл: `frontend/src/features/commercial-offer/components/CommercialOfferWizard.tsx`.

3. **`draftStorage.save/clear` — try/catch** — переполнение квоты sessionStorage
   (большой драфт с OCR-текстами 12 страниц) или недоступность хранилища больше не роняют
   эффект React. Персист драфта — best-effort.
   Файл: `frontend/src/features/commercial-offer/store/draftStorage.ts`.

## Тесты

- Новый регрессионный тест: `frontend/src/features/commercial-offer/store/draftStorage.test.ts`
  (round-trip, corrupt JSON → null, QuotaExceededError в save, SecurityError в clear).

## Верификация

- `npm run typecheck` — чисто.
- `npx vitest run src/features/commercial-offer` — 56 файлов / 373 теста, все зелёные
  (368 существующих + 5 новых).
- Линтер по изменённым файлам — без замечаний.

## Дальше по плану ревью

2. Конфиг по изделиям (`Record<ProductType, ...>`) — убрать тернарники и фабрики мутаций/API.
3. Слияние input-шагов и preview-панелей (~−2500 строк дублей).
4. Расчленение `CommercialOfferWizard` (1823 строки) + объединение sync-кейсов редьюсера.
5. Типизация `order_data` (discriminated union по `product_type`).
6. Тесты на уровне визарда до/после пп. 2–4.
