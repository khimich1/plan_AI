# Spec: КП — номер после сохранения и цены со скидкой на шаге 3

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**План**: [../develop/plans/2026-09-09-kp-saved-number-and-line-discount.md](../develop/plans/2026-09-09-kp-saved-number-and-line-discount.md)  
**Дата**: 2026-09-09  
**One-pager**: [../ideas/kp-saved-number-and-line-discount.md](../ideas/kp-saved-number-and-line-discount.md)  
**Связанные модули:** `CommercialExportService.build_offer_identity`, `commercial_draft_lifecycle.save_offer` / `get_draft_details`, `CalculationResultStep`, `SaveOfferSection`, `DownloadFilesSection`, архив `generate_document` (`offer_number = str(kp_id)`)

---

## Assumptions I'm Making

1. **Номер КП = `kp_id` как строка** (`"1188"`), без префикса `КП-`. Так уже в `archive_service.generate_document` и в hydrate-фикстурах.
2. **`WEB_*` только пока нет `saved_offer.kp_id` / `resume_kp_id`.** Формат как сейчас: `WEB_{draft_id[:8].upper()}` (подчёркивание, не дефис).
3. **Не предсказывать `get_next_kp_number()`.** Мёртвые тесты `test_build_offer_identity_uses_predicted_kp_number` переписать под WEB/saved, не воскрешать predicted.
4. **Первый save:** сначала запись в архив → `kp_id`, потом генерация документов с этим номером. Сейчас `save_offer` зовёт `generate_files(..., ("xlsx",))` *до* persist — поэтому XLSX внутри save тоже с `WEB_*`.
5. **После save UI не должен предлагать старые WEB-файлы.** Реген всех уже лежащих в `generated_files` видов (`pdf` / `xlsx` / `breakdown`; `schema` — только если уже была). Новые виды сами не создаём.
6. **Скидка на шаге 3 — display-only.** `order_data[].unit_price` = прайс. Отображение: `unit_price * (1 - discount_percent/100)`, формула как в `core/commercial_offer.py` / `commercial_offer_xlsx.py`. Целевая сумма по-прежнему от прайса.
7. **Оба бага в одном MVP.** Архивный drawer (`formatLinePrice`) не трогаем, если уже показывает `discounted_price`.
8. **Web-only**, существующий auth; без новых таблиц и без новых зависимостей.

→ Поправьте сейчас, иначе реализация идёт по этому списку.

---

## Decisions locked (из ideation)

| # | Тема | Решение |
|---|------|---------|
| D1 | Identity | Saved/resume → `str(kp_id)`; иначе `WEB_{draft[:8].upper()}` |
| D2 | Предсказание номера | Нет |
| D3 | Резерв `kp_id` при старте черновика | Нет |
| D4 | Скачивание до save | Остаётся (не направление C) |
| D5 | Файлы после save | Реген с новым номером; не оставлять WEB-имена в списке |
| D6 | Таблица шага 3 | Заменить цену и сумму строки на цену со скидкой |
| D7 | `unit_price` в данных | Не затирать |
| D8 | Формат номера | `str(kp_id)`, как архив |
| D9 | Чистка `outputs/` | Не удаляем старые WEB-файлы с диска |

---

## Objective

Менеджер на шаге результата видит в таблице те же цены, что уйдут в PDF. После «В архив» офис и клиент видят один номер — тот же `kp_id`, что в архиве и в шапке документов.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер, ещё не нажал «В архив» | видеть `WEB_*` в identity/черновых файлах | не путать черновик с номером из архива |
| US-2 | менеджер, сохранил КП | видеть в плашке «в архиве: 1188» (не `WEB_745903ED`) | знать номер для разговора с офисом |
| US-3 | менеджер, после save качает PDF/XLSX | получить файлы с номером 1188 в имени и в шапке | отправить клиенту тот же номер, что в архиве |
| US-4 | офис | найти КП по цифре с бумаги | поиск совпал с `kp_id` |
| US-5 | менеджер, нажал «Применить скидку» 10% | увидеть в строках цену и сумму со скидкой | сверить таблицу с PDF до отправки |
| US-6 | менеджер, правит целевую сумму | чтобы база считалась от прайса | обратный % не сломался |

### Reframed success criteria

| Требование | Измеримый критерий |
|------------|-------------------|
| «WEB только пока не сохранено» | `get_draft_details` / `save_offer.result_card.offer_number`: нет `saved_offer.kp_id` → `WEB_…`; есть → `str(kp_id)` |
| «После save файлы с номером КП» | После save список `files` не содержит `WEB_` / `kp_{draft[:8]}`; `offer_number` в PDF/XLSX = `str(kp_id)`; `file_stem` начинается с `kp_{kp_id}_` |
| «Повторная генерация» | `POST generate files` после save снова пишет номер `kp_id`, не `WEB_*` |
| «Resume/дополнение» | Черновик с `resume_kp_id` / `saved_offer.kp_id` сразу с этим номером, без `WEB_*` |
| «Скидка в вебе» | После apply 10% колонки Цена/Сумма на шаге 3 = прайс×0.9 (и ×qty); итоги как сейчас (уже со скидкой) |
| «Прайс жив» | `order_data[].unit_price` после apply не меняется; целевая сумма по-прежнему сходится |

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | Python 3, FastAPI, Pydantic v2 — `CommercialExportService`, `commercial_draft_lifecycle` |
| Domain | `kp_id` из `kp_repository.save_offer` / `update_offer_from_order_data`; формула скидки как в `core/commercial_offer.py` |
| Frontend | React 19, TypeScript; `CalculationResultStep`, `SaveOfferSection`, `DownloadFilesSection` |
| Tests | pytest: identity + save/generate порядок; Vitest: display-price helper + шаг 3 / SaveOfferSection |

---

## Commands

```bash
# Backend
source venv/bin/activate
pytest tests/test_commercial_web_flow.py tests/test_commercial_draft_append.py -q
pytest tests/ -k "offer_identity or save_offer or generate_files" -q

# Frontend
cd frontend && npm run typecheck
cd frontend && npm run test -- --run src/features/commercial-offer

# Dev
./run+logs.sh
```

---

## Project Structure

```
ai_docs/ideas/kp-saved-number-and-line-discount.md   # идея
ai_docs/specs/kp-saved-number-and-line-discount.md   # этот spec

app/services/commercial_export_service.py            # build_offer_identity(draft_id, metadata)
app/services/commercial_draft_lifecycle.py           # save: persist → generate; result_card из kp_id
app/services/commercial_workflow_service.py          # тонкий фасад, без своей копии identity

frontend/src/features/commercial-offer/
  lib/formatOfferNumbers.ts                          # или рядом: discounted unit/sum
  lib/lineDiscountDisplay.ts                         # NEW: pure display helper (предпочтительно)
  components/steps/CalculationResultStep.tsx         # Цена/Сумма со скидкой
  components/SaveOfferSection.tsx                    # плашка уже из result_card.offer_number
  components/DownloadFilesSection.tsx                # без отдельной логики, если files с бэка верные

tests/test_commercial_web_flow.py                    # identity: WEB vs saved; без predicted
```

---

## Current vs target

### Identity (баг)

Сейчас `CommercialExportService.build_offer_identity(draft_id)` всегда:

```text
offer_number = WEB_{draft[:8].upper()}
file_stem    = kp_{draft[:8]}_{YYYYMMDD_HHMMSS}
```

`get_draft_details` и `save_offer` возвращают этот payload даже когда `saved_offer.kp_id` уже есть. `save_offer` генерирует xlsx **до** `save_offer()` в репозиторий — шапка файла = WEB.

Архив при этом генерирует документы с `offer_number = str(kp_id)`.

Сиротские тесты ждут `workflow._build_offer_identity` с `predicted_kp_id` / saved `777`. Метода нет. Predicted **не** восстанавливаем.

### Скидка (баг)

`calculate_total_cost` считает `discounted_price` только для итога. PDF/XLSX применяют фактор при генерации. Шаг 3 рисует `item.unit_price` и `formatOfferSum(qty, unit_price)`. После «Применить скидку» `hydrate-draft` обновляет `totals` и `discount_percent`, строки визуально те же.

---

## Code Style

Identity — один метод, читает metadata, не угадывает следующий id:

```python
def build_offer_identity(
    self,
    draft_id: str,
    metadata: dict[str, Any] | None = None,
) -> tuple[str, str, str]:
    now = datetime.now()
    kp_id = _resolved_kp_id(metadata or {})  # saved_offer.kp_id or resume_kp_id
    if kp_id is not None:
        offer_number = str(int(kp_id))
        stem_id = offer_number
    else:
        offer_number = f"WEB_{draft_id[:8].upper()}"
        stem_id = draft_id[:8]
    offer_date = now.strftime("%d.%m.%Y")
    file_stem = f"kp_{stem_id}_{now.strftime('%Y%m%d_%H%M%S')}"
    return offer_number, offer_date, file_stem
```

Порядок первого save (обязателен):

```text
1. persist KP → kp_id          # xlsx_path может быть пустым на этот шаг
2. metadata.saved_offer = {kp_id, ...}
3. generate_files(pdf, xlsx, [breakdown если был])
4. при необходимости дописать xlsx_path в КП
5. result_card.offer_number = str(kp_id)
```

Resume (уже есть `kp_id`): generate сразу с этим номером; порядок persist/generate не критичен, identity не `WEB_*`.

Display скидки (чистая функция, без мутации draft):

```ts
export const discountedUnitPrice = (
  unitPrice: unknown,
  discountPercent: unknown,
): number | null => {
  const price = toNumber(unitPrice);
  const percent = toNumber(discountPercent) ?? 0;
  if (price === null) return null;
  const clamped = Math.min(Math.max(percent, 0), 100);
  return price * (1 - clamped / 100);
};
```

В `CalculationResultStep` все ветки таблицы (плиты / марки+класс / ступени) показывают `discountedUnitPrice(item.unit_price, draft.metadata.discount_percent)` и сумму от этой цены. Не писать скидку обратно в `order_data`.

`SaveOfferSection` уже берёт `lastSaveResult?.result_card.offer_number ?? draft.offer_identity.offer_number` — после бэкенд-фикса плашка зеленеет сама, отдельный хак на фронте не нужен.

---

## Testing Strategy

| Уровень | Что |
|---------|-----|
| Pytest identity | Нет `kp_id` → `WEB_` + stem `kp_{draft[:8]}_`; есть `saved_offer.kp_id=777` → `"777"` + `kp_777_`; `resume_kp_id` то же |
| Pytest save | После первого save `result_card.offer_number == str(kp_id)` и `files` pdf/xlsx не WEB; generate после save не возвращает `WEB_` |
| Pytest regression | Resume update (MNA-304) по-прежнему тот же `kp_id`, не create |
| Vitest helper | 10% от 42508 → 38257.2; 0% = прайс; 100% = 0; невалидная цена → null |
| Component | `CalculationResultStep`: при `discount_percent: 10` в ячейке не сырой `unit_price`; `SaveOfferSection`: saved → номер из `result_card` |

Сироту `test_build_offer_identity_uses_predicted_kp_number` **заменить**, не чинить через `get_next_kp_number`.

Покрытие: обе оси MVP (identity + display discount). Не требовать e2e PDF-парсинга, если pytest проверяет `offer_number`, уходящий в `generate_offer_pdf` / `generate_offer_xlsx`.

---

## Boundaries

**Always:**
- Один источник identity: export service + metadata (`saved_offer` / `resume_kp_id`)
- Persist → generate на первом save
- Display-only скидка; прайс в `order_data` неприкосновенен
- Тесты identity до правки UI-мелочей
- Формат номера как в архиве (`str(kp_id)`)

**Ask first:**
- Резерв/предсказание номера
- Запрет скачивания до save
- Префикс `КП-` в номере
- Возвращать `discounted_price` в API `order_data` (если display-only на клиенте недостаточно для паритета округления)
- Удаление старых файлов из `outputs/`

**Never:**
- `get_next_kp_number()` как offer_number черновика
- Запись скидки в `unit_price`
- Новые колонки БД / новые product types
- Менять формулу `calculate_total_cost` «чтобы таблица сошлась» — таблица должна читать ту же формулу, что документы
- Commit секретов / живых БД

---

## Success Criteria

1. Несохранённый черновик: `offer_identity.offer_number` вида `WEB_XXXXXXXX`; имена файлов с hex черновика.
2. После «В архив»: плашка `в архиве: {kp_id}`; `offer_identity.offer_number === str(kp_id)`.
3. После save кнопки «Скачать» отдают файлы, в имени и в `offer_number` генерации — этот `kp_id`, не `WEB_*`.
4. Повторное «Сформировать» / дозагрузка после save не возвращает WEB-identity.
5. Resume/редактирование из архива: номер сразу `kp_id`.
6. «Применить скидку» на шаге 3 меняет Цена и Сумма каждой строки (плиты и прочие типы на этом шаге); итоги как сейчас.
7. Целевая сумма / повторный apply % работают на прежнем `unit_price`.
8. pytest identity+save и vitest display+шаг 3 зелёные; typecheck фронта зелёный.

---

## Out of scope (Not Doing)

- Предсказание / резерв номера
- Скрытие документов до save
- Две колонки прайс + скидка
- Переименование уже скачанных пользователем файлов на его диске
- Удаление WEB-артефактов из `outputs/`
- Архивный drawer (уже `discounted_price`), кроме регрессии
- Схема раскладки «сама после save», если менеджер её не формировал
- Смена модели скидки (построчные разные %, скидка только на итог без строк)

---

## Open Questions

1. ~~Предсказывать номер до save?~~ Нет (D2).
2. ~~Авто-реген PDF при save или только при следующей генерации?~~ Авто-реген уже существующих видов, иначе UI продолжает показывать WEB-файлы (D5).
3. Округление display: как в документах (`price * factor`, без отдельного `round` до 2 знаков до форматирования) — **да**, `toLocaleString` уже режет отображение. Если визуально разъедется с PDF на копейку — тогда вынести общий round; не блокирует старт.
4. ~~Нужно ли человеку подтвердить этот spec, прежде чем план/код?~~ План запрошен 2026-09-09; код не начинать, пока план не одобрен.
