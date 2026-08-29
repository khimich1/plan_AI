# Spec: Живая подсветка нераспознанных строк в поле ввода КП

**Статус**: спека принята 28.08.2026 · план: [ai_docs/develop/plans/2026-08-28-unparsed-line-live-highlight.md](../develop/plans/2026-08-28-unparsed-line-live-highlight.md) · implement после явного старта
**Дата**: 28.08.2026
**One-pager**: [ai_docs/ideas/unparsed-line-live-highlight.md](../ideas/unparsed-line-live-highlight.md)
**Контекст контура**: [stabilizaciya-p1-commercial-2026-08-28.md](./stabilizaciya-p1-commercial-2026-08-28.md)

## Objective

**Проблема.** Менеджер добавляет позиции текстом. Строки, которые парсер не понял, не чинятся в поле ввода: поле очищается, ошибка всплывает мёртвой жёлтой плашкой в «Состав КП» и едет до шага 3. С плашкой ничего нельзя сделать.

**Цель.** Пока печатают список, каждая непустая строка после короткой паузы либо нормальная, либо красная с причиной. «Обработать текст» / «Добавить к списку» нельзя нажать, пока проверка не закончилась или есть красные. В состав КП через текстовый ввод нераспознанное не попадает. Плашка «Не попали в состав» перестаёт быть основным сигналом.

**Пользователь:** менеджер, который набивает или дописывает список изделий в мастере КП.

**Успех:** на текстовом пути нельзя отправить красную строку; видно *какую* строку чинить; то же поведение у всех шести типов; шаг 3 не орёт о строках, которые пользователь уже не может поправить в этом поле.

## ASSUMPTIONS

Поправьте до PLAN / IMPLEMENT — иначе работаем с этим.

1. **Стоп только на обработке текста.** «Список верен», «Готово, далее» и `WizardNextRequiredAction` не получают `resolve_unparsed`. OCR-сверка как сейчас: подсветка в `PlateListEditor`, без серой «Список верен».
2. **Фото отдельно.** Выбрано только изображение — линт текста не нужен, «Распознать фото» живая (как сейчас, если файл есть). Текст и фото в карточке источника по факту взаимоисключающие (выбор файла чистит текст).
3. **Кнопка не меняет подпись.** Серая с тем же текстом («Обработать текст» / «Добавить к списку»). Подсказка: в полёте — «Проверка списка…»; ошибка сети/HTTP — «Не удалось проверить список»; есть красные — «Исправьте красные строки». Без прыжка вёрстки.
4. **Общая карточка источника сразу**, не «сначала плиты». Иначе шесть копий разъедутся. Хук линта один, шаги только прокидывают `productType` и подписи.
5. **Линт = построчный line-parser**, не `generate_preview` и не полный `parse_plate_text` с нормализацией всего текста. Единица — физическая строка textarea (`\\n`). Пустая / из пробелов — не ошибка. Для плит после `parse_line` ещё `validate_plate_values`. «Нет в прайсе» и «шире стандартной» — не красные здесь: формат понят, дальше существующие карточки после обработки.
6. **`POST /commercial/parse` расширяем**, новый URL не плодим. `product_type` по умолчанию `plates`. Старый ответ для плит **аддитивен**: текущие поля остаются, фронт линта читает новое поле `lines`. Вызов не пишет черновик.
7. **Append-цикл** линтит текущий `product_type` шага (сваи после плит — парсер свай).
8. **Мёртвую плашку снимаем** только как UX текстового ввода: баннер «Не удалось распознать строк: N» и блок «Не попали в состав» в `Kp*PreviewPanel` и дубль в `CalculationResultStep`. Прочие warnings (нагрузка по умолчанию для Д×Ш×H и т.п.) не трогаем. OCR-нераспознанное на сверке остаётся подсветкой списка.
9. **Слэш `40,3/2,6` в парсер не учим** в этой спеке. Строка честно красная.
10. **Коммиты агент не делает**, пока явно не попросите.

→ Поправьте сейчас или PLAN пойдёт с этими допущениями.

## Tech Stack

- Backend: Python 3, FastAPI, Pydantic v2, pytest (`tests/`)
- Frontend: React 19, TypeScript, Vite, TanStack Query, vitest + Testing Library
- Парсеры: `core/plate_line_parser.py`, `pile_line_parser.py`, `step_line_parser.py`, `march_line_parser.py`, `bridge_pile_line_parser.py`, `fbs_line_parser.py`
- Контур КП после P1: `ProductDraftSpec` / `ProductDraftHandler`; HTTP по-прежнему раздельные update-URL, **кроме** этого расширения `/parse`

## Commands

```bash
# Backend (корень, venv)
pytest tests/test_commercial_web_flow.py tests/test_commercial_*_flow.py tests/test_commercial_draft_append.py -q
pytest tests/ -q -k "parse or unparsed or line_parser"

# Frontend
cd frontend && npm run test
cd frontend && npm run typecheck

# Dev
./run_local.sh
```

## Project Structure

```
app/schemas/commercial.py                          → CommercialParseRequest + parse lines DTO
app/api/v1/endpoints/commercial.py                 → POST /parse (тот же path)
app/services/commercial_line_lint.py               → НОВЫЙ: построчный lint по product_type
app/services/product_draft_config.py               → lint_fn в spec (не generate_preview)
core/*_line_parser.py                              → без смены грамматики; только вызываем
frontend/src/features/commercial-offer/api/commercialOfferApi.ts
frontend/src/features/commercial-offer/hooks/useSourceTextLint.ts   → НОВЫЙ
frontend/src/features/commercial-offer/components/SourceInputCard.tsx → НОВЫЙ, общая карточка
frontend/src/features/commercial-offer/components/PlateListEditor.tsx  → оверлей без обязательного draft
frontend/src/features/commercial-offer/lib/plateLineHighlights.ts
frontend/src/features/commercial-offer/components/steps/*InputStep.tsx → шесть шагов тонкие
frontend/src/features/commercial-offer/components/Kp*PreviewPanel.tsx
frontend/src/features/commercial-offer/components/steps/CalculationResultStep.tsx
tests/                                             → parse lint все типы + compat плит
frontend/.../*.test.ts(x)                          → хук, карточка, гейт кнопки, панели
```

## Code Style

Слои: роутер тонкий → сервис линта → `core` line-parser. Не звать ILP, прайс, `DraftStore`. Pydantic v2, `product_type: ProductType`. На фронте один хук и одна карточка; шаги не копируют дебаунс.

```python
# Контракт строки линта (индекс = номер строки textarea, 0-based)
class CommercialParseLine(BaseModel):
    index: int = Field(ge=0)
    text: str
    empty: bool = False
    ok: bool
    reason_text: str | None = None
```

```tsx
// Кнопка: подпись не меняем, только disabled + title
<Button
  disabled={!canSubmitSource}
  title={sourceSubmitBlockReason} // undefined | «Проверка списка…» | «Не удалось проверить список» | «Исправьте красные строки»
>
  {primaryRecognizeLabel}
</Button>
```

## Design

### API

`POST /api/v1/commercial/parse`  
Auth: как сейчас (`REQUIRE_ADMIN_OR_MANAGER`).

**Request** (обратная совместимость):

```json
{ "text": "ПБ 78-12-8п 2\\nплохо", "product_type": "plates", "lint_only": true }
```

- `text`: `min_length=1`, `max_length=50000` (горячий путь дебаунса)
- `product_type`: опционально, default `"plates"`, тот же `ProductType` литерал, что у черновика
- `lint_only`: опционально, default `false`. Живой линт фронта шлёт `true` и **не** вызывает `CommercialService.parse` даже для плит. Без флага ответ плит по-прежнему аддитивен (старые ключи + `lines`).

**Response:**

- Для `plates` без `lint_only` — все текущие поля (`order`, `normalized_text`, `unparsed_lines`, `warnings`, `wide_plate_lines`, `dobor_pairs`, `diagnostics`) **плюс** `product_type` и `lines`.
- Для остальных типов **или** `lint_only: true` — `product_type`, `lines`, `unparsed_lines` (строки `ok=false`, без суффикса «пропущено» в `text`). Поля заказа/оптимизации не обязательны; не звать preview-сервисы.
- `lines`: по одной записи на каждую физическую строку входа, включая пустые (`empty: true`, `ok: true`, `reason_text: null`).

Пустой textarea фронт **не** шлёт (кнопка и так серая из‑за пустого источника).

Ошибки: невалидный `product_type` → 422; как сейчас, внутренности парсера в 500 не утекают.

### Линт (сервер)

Новый сервис (имя ориентир `commercial_line_lint.py`), в `ProductDraftSpec` — `lint_lines(text) -> list[LineLint]`, не `generate_preview`.

| Тип | Вызов на непустую строку | Красная если |
|-----|--------------------------|--------------|
| plates | `parse_line` + при parsed `validate_plate_values` | не parsed или validation не ok |
| piles / steps / marches / bridge_piles / fbs | соответствующий `parse_*_line` | не `parsed` |

Цена, нагрузка «нет в прайсе», ширина «wide» на линт не влияют.

### Фронт

**`useSourceTextLint({ text, productType, enabled })`**

- `enabled`: есть непустой текст и нет выбранного фото
- дебаунс **500 ms** после последнего ввода
- `POST /parse` с `product_type` и `lint_only: true`; отмена устаревших ответов (seq / AbortController)
- сразу при изменении текста: `isPending = true` (кнопка серая, не ждём 500 ms)
- пустой текст: `isPending = false`, красных нет, запрос не уходит

**Карточка источника (`SourceInputCard`)** — одна на шесть шагов: textarea с оверлеем как у `PlateListEditor` (красные = `unparsed`, title = `reason_text`), файл, кнопки. `PlateListEditor` принимает подсветку снаружи (карта index → kind/title), без обязательного полного `draft` для режима линта.

**Гейт кнопки обработки / добавления:**

`disabled`, если любое:

- нет источника (пустой текст и нет файла) — как сейчас
- идёт OCR/обработка/AI — как сейчас
- линт включён и (`isPending` или `isError` или есть `ok === false` у непустой строки)

«Распознать фото» при файле без текста: линт выключен, гейт линта не серит.

**Предпросмотр и шаг 3:** не показывать warning «Не удалось распознать строк: N» и список «Не попали в состав». Остальные предупреждения оставить. Метаданные `unparsed_lines` в черновике после OCR пока могут писаться — сверка с фото их читает; в предпросмотре после «Список верен» этот список пользователю не дублируем мёртвой плашкой.

## Testing Strategy

**Backend (pytest)**

- Совместимость: `POST /parse` без `product_type` на плитном тексте — 200, прежние ключи на месте, плюс `lines`
- Плиты: валидная + пустая + `ПБ 40,3/2,6-8п` → у слэша `ok=false`, пустая `empty=true`/`ok=true`
- Каждый из шести `product_type`: хотя бы одна ok и одна not ok
- Линт не дергает optimize / DraftStore (моки/шпион, если легко; иначе отсутствие полей preview)
- Невалидный product_type → 422
- RBAC: без cookie как у текущего `/parse`

**Frontend (vitest)**

- Хук: дебаунс, отбрасывание stale, пустой текст без fetch, `isError` при отказе `/parse`
- Кнопка серая при pending, при красных и при ошибке сети (title «Не удалось проверить список»)
- Гейт футера «Добавить к списку» через `resolveSourceSubmitDisabled`
- Файл без текста — кнопка не серится линтом
- Карточка: красная строка + title с причиной
- `lint_only: true` на плитах не зовёт `CommercialService.parse`; текст длиннее 50000 → 422
- Панели предпросмотра / шаг 3: нет «Не попали в состав» и нет баннера «Не удалось распознать строк»

**Вручную (после зелёных тестов, в браузере):** дописать к списку плит одну ломаную строку — краснеет, кнопка серая; исправить — кнопка живая; append свай — свой парсер.

## Boundaries

- **Always:** тесты parse-compat и гейта кнопки зелёные до объявления done; линт без preview/ILP; один хук + одна карточка источника; не блокировать «Далее» / «Список верен»
- **Ask first:** новый URL вместо расширения `/parse`; учить парсер слэшу; серить «Список верен»; менять `WizardNextRequiredAction`
- **Never:** второй парсер на TypeScript; вызов `generate_preview` **или полного `CommercialService.parse`** с дебаунса (`lint_only`); третья карточка wide/unpriced; коммит секретов; схлопывание остальных 32 URL мастера

## Success Criteria

- [x] Текстовое поле на шаге изделия (все шесть типов) после паузы красит нераспознанные непустые строки
- [x] «Обработать текст» / «Добавить к списку» нельзя нажать, пока проверка в полёте или есть красные
- [x] Пустые строки не краснеют и не блокируют
- [x] Только фото — линт не мешает «Распознать фото»
- [x] `/parse` без `product_type` по-прежнему работает для плит (старые ключи)
- [x] Предпросмотр и шаг 3 не показывают мёртвый список нераспознанных и баннер «Не удалось распознать строк: N»
- [ ] Wide / unpriced / OCR-сверка не сломаны этим изменением
- [x] Ошибка `/parse`: кнопка серая, title «Не удалось проверить список»
- [x] Живой линт шлёт `lint_only: true` и не тянет полный `parse` плит
- [x] Текст длиннее 50000 символов → 422
- [ ] `pytest` commercial safety net и `cd frontend && npm run test` + `typecheck` зелёные

## Follow-up after review (29.08.2026)

Разбор замечаний [Quality review] — что чиним в этом срезе:

| # | Замечание | Решение |
|---|-----------|---------|
| 1 | Title при ошибке сети врёт «Исправьте красные строки» | Чиним: отдельный title «Не удалось проверить список» |
| 2 | Нет теста на `isError` / серую кнопку при сети | Чиним: хук + карточка |
| 3 | Гейт «Добавить к списку» в футере без теста | Чиним: тесты `resolveSourceSubmitDisabled` |
| 4 | Нет теста оверлея карточки | Чиним: красная строка + `title` причины |
| 5 | Пустые строки в `PlateInputStep.tsx` | Чиним: схлопнуть лишние blank lines |
| 6 | Двойной проход плит (`lint` + `service.parse`) на дебаунсе | Чиним: `lint_only` на существующем `/parse`, фронт шлёт `true`. Новый URL не делаем |
| 7 | Нет `max_length` у `text` | Чиним: 50000 символов → 422 |
| 8 | FYI про OCR-плашку | Не код |
| 9 | Шпион импортов «линт не зовёт preview» | Не трогаем: для текущего модуля достаточно |

## Open Questions

Нет блокирующих. План: [2026-08-28-unparsed-line-live-highlight.md](../develop/plans/2026-08-28-unparsed-line-live-highlight.md).
