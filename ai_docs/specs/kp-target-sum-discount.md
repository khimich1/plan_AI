# Spec: КП — целевая сумма и подтверждение скидки >16%

> **Источник идеи:** [`ai_docs/ideas/kp-target-sum-discount.md`](../ideas/kp-target-sum-discount.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT (ожидает запуска)  
> **Статус:** ready for implementation  
> **Связанные модули:** `CalculationResultStep`, `OfferDetailsDrawer`, `ResetConfirmDialog` (паттерн), `core/commercial_pricing.calculate_total_cost`, archive `PATCH .../discount`, commercial metadata/calculate
> **План и задачи:** [`ai_docs/develop/plans/2026-08-05-kp-target-sum-discount.md`](../develop/plans/2026-08-05-kp-target-sum-discount.md)

---

## Assumptions I'm Making

1. **Скидка только на продукцию.** Как в `calculate_total_cost`: `discount_percent` уменьшает сумму позиций; `delivery_service_total_rub` (рейс × число рейсов) **не** скидывается. Целевая сумма T = сумма позиций со скидкой + доставка.
2. **Обратная формула:**  
   `base_plates = Σ(unit_price × qty)` при 0%  
   `delivery = delivery_service_total`  
   `discount% = 100 × (1 − (T − delivery) / base_plates)`  
   при `base_plates > 0` и `delivery ≤ T ≤ base_plates + delivery`.
3. **Порог:** строго **`discount > 16`** (16.00 без модалки; 16.01 — с модалкой). Для всех `product_type`.
4. **Keyword:** точное совпадение `ПОДТВЕРЖДАЮ` (как `confirmKeyword` в `ResetConfirmDialog`). Регистр/пробелы: trim, регистр как в эталоне (кириллица uppercase).
5. **Нет записи в БД** факта подтверждения; нет ролей руководителя; нет блокировки PDF/XLSX.
6. **Хранение** по-прежнему только `discount_percent` (0–100). Целевая сумма — UI-вход, не колонка в БД.
7. **Округление (предложение):** `%` округляем до **2 знаков** после запятой; после применения серверного пересчёта итог может отличаться от T на **≤ 1 ₽** — это ок, отдельный «добивающий» residual не делаем.
8. **T выше базы без скидки:** ошибка валидации, скидку не применяем (наценки нет). Сообщение с фактическим максимумом (`base_plates + delivery`).
9. **T ниже доставки:** ошибка (отрицательная сумма позиций невозможна).
10. **`base_plates = 0`:** поле целевой суммы disabled / ошибка «нет позиций для расчёта».
11. **Двусторонняя синхронизация черновиков:** ввод `%` → пересчёт черновика целевой суммы; ввод целевой суммы → пересчёт черновика `%`. Применение — по OK (как сейчас у скидки), с модалкой если итогвый `% > 16`.
12. **Обратный расчёт на клиенте** из уже известных `unit_price`/`qty`/`delivery_service_total` (архив) или `order_data` + logistics → delivery (визард); отдельный API `target_sum → discount` **не** обязателен в MVP. При сомнении в delivery в визарде — взять из последнего `totals` / пересчёт через существующий calculate с `discount=0` только если без этого нельзя стабильно посчитать.
13. **Смена логистики** после большой скидки: модалка **не** повторяется, пока `%` не меняют заново через apply.
14. **Web-only**, существующий auth; без новых зависимостей.

> Assumptions Q1–Q5 approved 2026-08-05; planning proceeds with them unchanged.

---

## Decisions locked (из ideation)

| # | Тема | Решение |
|---|------|---------|
| D1 | Пользователь v1 | Только менеджер |
| D2 | >16% | Модалка + keyword `ПОДТВЕРЖДАЮ` + ОК/Отмена (паттерн очистки БД) |
| D3 | Текст | «ВНИМАНИЕ: ДЛЯ СКИДКИ ВЫШЕ 16% НЕОБХОДИМО ОДОБРЕНИЕ РУКОВОДСТВА.» |
| D4 | Ввод | Целевая сумма (₽) **и** скидка (%), синхрон в обе стороны |
| D5 | Состав T | Итого с НДС **включая** доставку |
| D6 | Продукты | Все типы КП |
| D7 | Аудит confirm | Не писать в КП |
| D8 | Точки UI | Визард шаг результата + архив drawer |

---

## Objective

Менеджер подгоняет итог КП под бюджет клиента («у клиента 2 млн») за секунды и не может «молча» применить скидку выше 16% без явного подтверждения.

### User stories

| # | Как менеджер… | Я хочу… | Чтобы… |
|---|---------------|---------|--------|
| US-1 | на шаге результата / в архиве | ввести целевую сумму в ₽ | система сама посчитала скидку % |
| US-2 | знаю нужный % | ввести скидку как сейчас | целевая/итог обновились согласованно |
| US-3 | скидка вышла >16% | увидеть жёсткое предупреждение с ОК/Отмена и вводом `ПОДТВЕРЖДАЮ` | случайно не отдать большую скидку |
| US-4 | нажал Отмена | откатить черновик | КП осталось с прежней скидкой |
| US-5 | работаю с любым типом продукции | одинаковое поведение | не учить разные правила |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| «Ввёл 2 млн — скидка сама» | Из T и базы считается `%`; после apply `total_with_vat` ≈ T (допуск ≤ 1 ₽) |
| «И % работает» | Изменение `%` → черновик T = f(%); apply как сегодня |
| «>16% — плашка» | Apply при `% > 16` без keyword невозможен; с keyword — применяется |
| «Защита от дурака» | ОК disabled пока input ≠ `ПОДТВЕРЖДАЮ` |
| «Архив = визард» | Оба экрана: целевая сумма + confirm |
| «Без наценки» | T > max → ошибка, discount не меняется |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2 — **минимальные изменения** (существующие update discount / calculate) |
| Domain | `core/commercial_pricing.calculate_total_cost` как источник истины формулы |
| Frontend | React 19, TypeScript, Vitest; UI: `Modal` / паттерн `ResetConfirmDialog` |
| Tests | pytest (если появится shared pure-функция %); Vitest для lib + компонентов |

---

## Commands

```bash
# Backend
source venv/bin/activate
pytest tests/ -k "discount or commercial_pricing or archive" -q

# Frontend
cd frontend && npm run typecheck
cd frontend && npm run test -- --run src/features/commercial-offer src/features/commercial-archive
cd frontend && npm run build

# Dev
./run+logs.sh
```

---

## Project Structure

```
ai_docs/ideas/kp-target-sum-discount.md          # идея
ai_docs/specs/kp-target-sum-discount.md          # этот spec
frontend/src/features/commercial-offer/
  components/steps/CalculationResultStep.tsx     # поле T + % + confirm
  lib/discountFromTargetSum.ts                   # NEW: pure math (предпочтительно)
  lib/discountFromTargetSum.test.ts
  components/HighDiscountConfirmDialog.tsx       # NEW или shared
frontend/src/features/commercial-archive/
  components/OfferDetailsDrawer.tsx              # то же UX
  components/HighDiscountConfirmDialog.tsx       # лучше shared в shared/ui или commercial-offer и re-export
core/commercial_pricing.py                       # reference only; optional pure helper если вынесем в core
app/services/archive_service.py                  # без новых полей, если клиентский расчёт достаточен
```

Предпочтение: **shared pure function** (TS и при необходимости Python-зеркало для тестов) + **один** confirm-диалог, переиспользуемый визардом и архивом.

---

## Code Style

Паттерн confirm — как админка:

```tsx
// Аналог ResetConfirmDialog: OK disabled until typedKeyword === "ПОДТВЕРЖДАЮ"
<Modal open={open} onClose={onCancel} title="Подтверждение скидки">
  <Alert tone="warning">
    ВНИМАНИЕ: ДЛЯ СКИДКИ ВЫШЕ 16% НЕОБХОДИМО ОДОБРЕНИЕ РУКОВОДСТВА.
  </Alert>
  <label>
    Введите <code>ПОДТВЕРЖДАЮ</code>, чтобы подтвердить:
    <input value={typedKeyword} onChange={...} />
  </label>
  <Button onClick={onCancel}>Отмена</Button>
  <Button disabled={typedKeyword.trim() !== "ПОДТВЕРЖДАЮ"} onClick={onConfirm}>
    ОК
  </Button>
</Modal>
```

Чистая математика (без React):

```ts
const APPROVAL_THRESHOLD_PERCENT = 16;

export function discountPercentFromTargetSum(params: {
  targetTotalWithVat: number;
  basePlatesTotalWithVat: number; // Σ unit_price * qty, discount 0
  deliveryTotal: number;
}): { ok: true; discountPercent: number } | { ok: false; error: string } {
  // validate bounds; return round(percent, 2)
}

export function targetSumFromDiscountPercent(params: {
  discountPercent: number;
  basePlatesTotalWithVat: number;
  deliveryTotal: number;
}): number {
  // base * (1 - d/100) + delivery, round 2
}

export function requiresHighDiscountConfirmation(discountPercent: number): boolean {
  return discountPercent > APPROVAL_THRESHOLD_PERCENT;
}
```

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Unit (Vitest) | Формула: типовой кейс T=2e6; границы `T = delivery`, `T = max`; `T > max` / `T < delivery`; порог 16 vs 16.01; округление 2 знака |
| Component | Confirm: OK disabled без keyword; Отмена не вызывает apply; с keyword вызывает apply |
| Integration (optional) | Archive/wizard handlers: apply path вызывает существующий mutation с посчитанным % |
| Pytest | Только если вынесем helper в `core/`; иначе не обязательно |

Покрытие: все ветки валидации `discountPercentFromTargetSum` + порог confirm.

---

## Boundaries

**Always:**
- Переиспользовать существующий apply discount (wizard calculate/meta + archive `PATCH /discount`)
- Единый порог и текст для всех product types
- Тесты на pure math до UI
- Отмена = никакого изменения сохранённой скидки

**Ask first:**
- Новый API endpoint вместо клиентского расчёта
- Запись audit-полей в БД
- Изменение формулы НДС/доставки в `calculate_total_cost`
- Конфиг порога в env/БД (сейчас константа)

**Never:**
- Наценка (отрицательная скидка) через это UI
- Блокировка скачивания файлов в v1
- Approve-workflow / роли руководителя в v1
- Разные скидки по строкам
- Commit секретов / живых БД

---

## Success Criteria

1. В визарде (шаг результата) есть поле «Целевая сумма (₽)» рядом со «Скидка (%)»; черновики синхронизируются в обе стороны.
2. В архиве (`OfferDetailsDrawer`) — то же.
3. Apply целевой суммы вызывает тот же путь сохранения `discount_percent`, что и %; итог с НДС ≈ T (≤ 1 ₽).
4. При `% > 16` без ввода `ПОДТВЕРЖДАЮ` скидка не применяется; с вводом — применяется; Отмена откатывает черновик.
5. T > max или T < delivery → понятная ошибка, без apply и без модалки.
6. Unit-тесты формулы и порога зелёные; typecheck/build фронта зелёные.
7. Поведение одинаково для plates / piles / steps / marches / bridge_piles / fbs.

---

## Out of scope (Not Doing)

- Роли / статус `awaiting_approval`
- `discount_acknowledged_at` в БД
- Блокировка PDF/XLSX
- Построчные скидки
- Env-конфиг порога (можно follow-up)
- Telegram / бот

---

## Open Questions (resolved before PLAN)

| # | Вопрос | Предложение spec |
|---|--------|------------------|
| Q1 | Допуск округления после apply | **≤ 1 ₽**, `%` до 2 знаков |
| Q2 | T выше суммы без скидки | **Ошибка**, без наценки |
| Q3 | Где считать обратный % | **Клиент** (pure TS); API не обязателен |
| Q4 | Live-sync черновиков при вводе или только по OK? | **При вводе** (onChange) синхронизировать парное поле; apply по OK |
| Q5 | Визард: откуда брать `deliveryTotal` / `basePlates`, если в `totals` нет явного delivery? | Считать `basePlates` из `order_data`; delivery = `total_with_vat(at 0%) − basePlates` через один calculate с discount=0 **или** из уже известной формулы рейсов, если UI уже показывает доставку |

Q1–Q5 подтверждены пользователем как записано. Единственный implementation checkpoint, не требующий продуктового решения заранее: проверить, что вывод delivery из текущего wizard total совпадает с серверной формулой (см. PLAN ниже).

---

## PLAN ✅

Полный технический план и декомпозиция Phase 3: [`2026-08-05-kp-target-sum-discount.md`](../develop/plans/2026-08-05-kp-target-sum-discount.md).

### Components and dependency order

1. `DISC-001` — новая чистая TS-библиотека `discountFromTargetSum` с единой формулой, округлением, границами и порогом.
2. `DISC-002` — один `HighDiscountConfirmDialog`, повторяющий lifecycle `ResetConfirmDialog`.
3. `DISC-003` — wizard `CalculationResultStep` + существующий apply в `CommercialOfferWizard`.
4. `DISC-004` — archive `OfferDetailsDrawer` + существующий PATCH mutation.
5. `DISC-005` — cross-surface regression and handoff.

`DISC-001` и `DISC-002` предшествуют обеим UI-интеграциям. После них `DISC-003` и `DISC-004` можно выполнять параллельно; финальная проверка последовательна.

### Data sources

- Archive берёт `baseProducts` из исходных `unit_price × qty` product lines и `delivery` из `delivery_service_total_rub`.
- Wizard берёт `baseProducts` из `draft.order_data`; так как response сейчас не содержит delivery отдельно, план выводит его из текущего server total и применённой скидки. До merge это сверяется с backend total на нулевой и ненулевой логистике.
- В обоих случаях apply передаёт только рассчитанный `discount_percent` через уже существующие пути; target sum остаётся локальным draft.

### Implementation risks and gates

- Rounding: shared module rounds `%` to 2 decimals and checks returned total against target with tolerance ≤1 ₽.
- Delivery inference: если сверка wizard покажет отличие >0.01 ₽, implementation останавливается для решения об explicit delivery field в существующем calculate response; новый endpoint не добавляется без отдельного согласования.
- Cancel: target и discount drafts restore from current saved data, mutation is never called.
- Checkpoints: pure-math unit tests → dialog tests → wizard/archive integration tests → typecheck, focused Vitest, build, targeted pytest.

## TASKS ✅

Задачи `DISC-001` … `DISC-005` определены в plan file; каждая содержит зависимости, точный список (не более пяти) файлов, acceptance criteria и verify command. Рекомендуемый первый task: **DISC-001 — shared target-sum math and tests**.
