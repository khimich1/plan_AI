# Spec: OCR retry без часового капа + сверка с исправленным текстом

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT 🔄 (focused tests green; awaiting user accept)  
**Дата**: 2026-09-03  
**Plan:** [../develop/plans/2026-09-03-kp-ocr-retry-no-hourly-cap.md](../develop/plans/2026-09-03-kp-ocr-retry-no-hourly-cap.md)  
**One-pager**: [../ideas/kp-ocr-retry-no-hourly-cap.md](../ideas/kp-ocr-retry-no-hourly-cap.md)  
**Related**: [kp-multi-page-screenshots.md](./kp-multi-page-screenshots.md), [kp-ocr-wait-and-ai-on-review.md](./kp-ocr-wait-and-ai-on-review.md), [ocr-verify-apply-and-upscale.md](./ocr-verify-apply-and-upscale.md), [unparsed-line-live-highlight.md](./unparsed-line-live-highlight.md)

## Objective

**Проблема.** (1) Очередь фото до 12 страниц упирается в `COMMERCIAL_OCR_UPLOADS_PER_HOUR=10` — 11-й кадр и retry после десятого дают 429. (2) Плохое распознавание выбранного кадра нельзя запустить заново без `append` в черновик. (3) Менеджер правит список справа, а блок «Нестандартная ширина» всё ещё предлагает split по замороженному OCR (`44-15-10п 5` → `34-3-10п 5` вместо `34-15-10п 15`).

**Цель.** Снять часовой кап. Кнопка «Перераспознать» под выбранным фото (на **всех** шагах ввода с фото) заменяет **только текст этой страницы** — в том числе если OCR отдал 2 строки из 10. Сверка страницы с **текущим текстом редактора** на всех шести типах изделий. Если OCR не удался — менеджер сам добавляет новое фото в конец очереди.

**Пользователь:** менеджер на batch-review конструктора КП (фото слева, список справа).

**Успех:** 12 фото в одном заходе без 429; retry выбранного кадра не дублирует позиции; после правки `44-15-10п 5` → `34-15-10п 15` карточка wide и qty считаются от новой строки.

---

## ASSUMPTIONS I'M MAKING

1. **Кап.** Default `COMMERCIAL_OCR_UPLOADS_PER_HOUR=0` = выключен. Сейчас Field `ge=1` — меняем на `ge=0`. `check_commercial_ocr_rate_limit`: при `lim <= 0` сразу return, события не пишем. Логин и `COMMERCIAL_UPLOAD_MAX_BYTES` не трогаем. Тесты с явным `=2` оставляем.
2. **Апскейл / 8 МиБ / `OCR_MAX_API_CALLS`.** Вне scope.
3. **Кнопка.** Под карточкой «Исходное фото» (под картинкой, не в шапке зума). Текст: «Перераспознать». Только **выбранная** страница. Replace её `batchReviewText` (+ editor, если страница активна). Остальные страницы не трогаем. Типичный кейс: на фото 10 строк, OCR вернул 2 — retry должен заменить список целиком новым проходом, не дописывать к двум строкам.
4. **Когда кнопка жива — везде, где видно выбранное фото с `File`.** `ready`, `error` **и `confirmed`**. Не только «ещё не подтвердили». После retry `confirmed` → снова `ready`: список изменился, «Список верен» заново. Единственный disable: эта страница сейчас `running` / `pending` (уже в полёте) или нет `File`. Не interrupt чужого кадра в очереди.
5. **Нет OCR-only API сегодня.** `update` с картинкой — только `append` | `replace` всего ввода. Retry **нельзя** делать через `append` (дубли) и нельзя через `replace` на multi-page (сотрёт другие страницы). Нужен новый endpoint, который гоняет тот же OCR-пайплайн и **не пишет черновик**.
6. **Single-page** (нет multi-сессии): retry может идти существующим `update` + `mode=replace` + тот же `File` — это и есть одна страница.
7. **Ошибка OCR.** Страница → `error`, тот же `File` остаётся. Сообщение — существующий `getErrorMessage`. Не удаляем страницу, не предлагаем «заменить файл». Менеджер сам добавляет кадр в конец очереди (как D11 multi-page спеки).
8. **Сверка с текстом — на всех шагах ввода** (плиты, сваи, марши, ступени, мостовые, ФБС). Список, подсветки и гейт «Список верен» этой страницы считаются по **live** тексту редактора, не по замороженному OCR. У плит дополнительно: wide-карточки / подсветка `wide` от live-строк (пример `44-15-10п 5` → правка `34-15-10п 15` → split с qty 15). У остальных типов нет `WidePlatesInlineSection` — не выдумываем wide, но не оставляем сверку на старом OCR-списке из 2 строк.
9. **Apply wide на сверке.** Сначала flush editor в черновик (`updateInput` text, без картинки), затем существующий `resolve_wide_plates` с `sourceLine` = **live** строки (после flush серверные id). Иначе Apply ищет старый OCR-line и менеджер правит дважды. Действие `confirm` (оставить 15 дм) по-прежнему легитимно — после Apply `wide_plates_resolved` как сейчас, calculate не блокируется.
10. **Детекция wide на FE.** MVP: канонические/голые марки вида `ПБ? L-W-нагрузка qty` (то, что в редакторе сверки), ширина в дм > 12. Полный порт `get_wide_plate_lines` (WxL в метрах) — не обязателен, если такие строки на сверке не встречаются.
11. **Типы изделий.** Кнопка retry **и** сверка с текстом редактора — все шесть `*InputStep` с фото. Карточка «Нестандартная ширина» по-прежнему только у плит (её нет у свай/маршей и т.д.).
12. **После retry обновить сигналы этой страницы.** Новый `normalized_text` в редактор; `ocr_verify_failed` страницы с нового ответа. Жёлтый баннер corrections — с нового прохода (или пустой), не от hydrate с 2 строками. AI-инструкция и unpriced по-прежнему не перепроектируем.
13. **Провайдер 429.** Существующий AI-provider error path. Не подменяем его текстом про «лимит загрузок».
14. **Коммиты** — только по просьбе. Не убивать `./run+logs.sh`. Без новых npm/pip.

→ Поправьте сейчас, иначе это locked для PLAN.

---

## Decisions locked (из ideation)

| # | Тема | Решение |
|---|------|---------|
| D1 | Часовой кап | `0` = выкл. Не «поднять до 40». |
| D2 | 4× / 8 МиБ | Не делаем |
| D3 | Кнопка | Выбранная картинка, replace текста страницы. **Везде** (`ready` / `error` / `confirmed`). |
| D4 | OCR fail | Менеджер добавляет фото в конец очереди |
| D5 | Сверка | На **всех шагах** ввода: страница сверяется с исправленным текстом |
| D6 | Merge при retry | Нет |
| D7 | Авто-retry | Нет |
| D8 | Неполный OCR | Retry — основной ответ на «распознало 2 строки из 10», не merge к двум строкам |
| D9 | Retry confirmed | Список новый → страница снова `ready`, «Список верен» ещё раз |

---

## User Stories

- Как **менеджер**, загружаю 12 фото подряд и не ловлю «Превышен лимит загрузок».
- Как **менеджер**, на любом шаге ввода с фото жму «Перераспознать» под текущим кадром — даже если уже нажал «Список верен» — и вижу **полный** новый список этой страницы (не 2 строки из 10); другие страницы на месте.
- Как **менеджер**, если кадр снова не распознался, добавляю другое фото в конец очереди — без отдельного мастера.
- Как **менеджер**, исправив `44-15-10п 5` на `34-15-10п 15`, вижу split и qty от исправленной строки, а не повтор той же ошибки OCR.

---

## Tech Stack

| Слой | Стек |
|------|------|
| Backend | FastAPI, существующий OCR (`resolve_source_input` / pipeline), Pydantic settings |
| Frontend | React 19, TS, Vitest, `useMultiPageRecognize`, `PlateInputStep` / peers, `WidePlatesInlineSection` |
| Docs | `core/config/settings.py`, `.env.example`, при необходимости `ai_docs/develop/architecture/rate-limiting.md` |

## Commands

```
# Кап
pytest tests/test_commercial_web_flow.py tests/test_commercial_ocr_policy.py -q -k "rate_limit or ocr_upload or OCR_UPLOADS"

# Retry API (имя теста уточнится в PLAN)
pytest tests/test_commercial_ocr_page.py -q

# FE
cd frontend && npm run test -- \
  src/features/commercial-offer/lib/widePlateSuggestion \
  src/features/commercial-offer/lib/liveWidePlateLines \
  src/features/commercial-offer/hooks/useMultiPageRecognize \
  src/features/commercial-offer/components/steps/PlateInputStep \
  src/features/commercial-offer/components/steps/PileInputStep \
  src/features/commercial-offer/components/WidePlatesInlineSection
cd frontend && npm run typecheck
```

## Project Structure

```
core/config/settings.py                          → default 0, ge=0
app/services/commercial_upload_validation.py     → skip limiter if lim <= 0
app/api/v1/endpoints/commercial.py               → POST drafts/{id}/ocr-page
app/services/product_draft_handler.py            → OCR без persist (тонкая обёртка)
app/schemas/commercial.py                        → response { normalized_text, ... }

frontend/.../api/commercialOfferApi.ts           → ocrPage(draftId, file)
frontend/.../hooks/useMultiPageRecognize.ts      → rerunPage(id)
frontend/.../components/steps/*InputStep.tsx     → кнопка под фото
frontend/.../lib/liveWidePlateLines.ts           → NEW: wide from editor text
frontend/.../lib/widePlateSuggestion.ts          → уже есть split
frontend/.../components/WidePlatesInlineSection.tsx
frontend/.../components/CommercialOfferWizard.tsx → gate + Apply flush-then-resolve

tests/test_commercial_web_flow.py                → cap 0; cap 2 по-прежнему 429
frontend/.../lib/liveWidePlateLines.test.ts
```

## Code Style

Кнопка — ghost/secondary рядом с подсказкой зума, не новый Card.

```tsx
<Button
  type="button"
  variant="ghost"
  onClick={() => void onRerecognize()}
  disabled={disabled || isRerecognizing}
>
  {isRerecognizing ? "Распознавание..." : "Перераспознать"}
</Button>
```

Limiter:

```python
def check_commercial_ocr_rate_limit(user_id: int) -> None:
    lim = get_settings().commercial_ocr_uploads_per_hour
    if lim <= 0:
        return
    _ocr_upload_limiter.check(user_id, max_events=lim)
```

Новый endpoint **не** вызывает update draft. Ответ — текст страницы (+ флаги OCR, которые UI уже умеет: `ocr_verify_failed`). Без нового product_type в form — тип берётся из черновика.

Live wide: не матчить карточки к `metadata.wide_plate_lines` по старой строке. Ключ — индекс/нормализованная **текущая** строка редактора.

## Testing Strategy

| Уровень | Что |
|---------|-----|
| pytest cap | `=2` → 3-я загрузка 429 (регресс). `=0` → ≥11 загрузок без 429 |
| pytest ocr-page | OCR вызывается; draft `input_text` / batches **байт-в-байт те же**; ответ содержит `normalized_text`; 404 чужой draft; без картинки 400; OCR выключен 503 |
| unit live-wide | `44-15-10п 5` → wide qty 5; после замены на `34-15-10п 15` → wide qty 15, suggestion содержит 12 и 3 и **15**; `34-12-10п 15` → не wide |
| RTL кнопка | есть под фото на `ready`, `error` и `confirmed`; нет/disabled на `running`; клик → loading; ошибка → страница error, File на месте |
| RTL/unit retry multi | 2 страницы, rerun второй → текст только второй страницы; первая не изменилась |
| typecheck | зелёный |

Моки OCR на бэке — как в `test_commercial_ocr_policy` / web_flow (не live GigaChat).

## Boundaries

**Always**
- `0` выключает кап; `>0` оставляет limiter
- Retry выбранной страницы без `append` и без wipe остальных
- Wide на сверке **плит** от live-текста; на всех шагах — список/гейт страницы от live-текста
- Кнопка на `confirmed`; после retry страница снова `ready`
- Ошибка OCR → `error` + существующее «добавить в конец»
- Провайдерские ошибки не маскировать текстом про лимит загрузок

**Ask first**
- Порт полного `get_wide_plate_lines` (WxL) на FE
- Порт live-wide на не-плиты (там нет wide-карточки)
- Unpriced live-sync
- Redis/shared rate limit
- Redis/shared rate limit
- Новый npm/pip

**Never**
- Апскейл 4×, снятие 8 МиБ, смена `OCR_MAX_API_CALLS`
- In-place «заменить файл» в слоте
- Авто-retry без клика
- Merge ручных правок с новым OCR
- Хранение фото на сервере / IndexedDB
- Коммит секретов / `plita.db`
- Трогать `./run+logs.sh`

## Success Criteria

| # | Критерий | Метод |
|---|----------|--------|
| S1 | Default кап 0; 11+ OCR-загрузок одного user без 429 | pytest |
| S2 | Явный `COMMERCIAL_OCR_UPLOADS_PER_HOUR=2` → 3-я загрузка 429, старый detail | pytest регресс |
| S3 | Multi-page: «Перераспознать» на странице 2 заменяет только её список; страница 1 без изменений; в draft нет четвёртого батча | pytest + unit FE |
| S4 | Single-page retry = replace этой единственной страницы | unit / RTL |
| S5 | Fail OCR → статус error, File жив, можно добавить ещё файл в хвост | RTL / hook |
| S6 | Редактор `34-15-10п 15` после OCR `44-15-10п 5` → wide suggestion от 34-15 qty 15, не `34-3-10п 5` | unit |
| S7 | Строка больше не wide (ширина ≤12) → карточка и блок «Список верен» не требуют wide | unit + RTL |
| S8 | Кнопка на всех шести `*InputStep` с фото; live-wide карточка только у плит; live-текст сверки на всех шагах | RTL минимум plates + один не-плитный step |
| S9 | OCR вернул 2 строки, retry вернул полный список → в редакторе новый полный текст, не «2 + хвост» | unit FE |
| S10 | Retry `confirmed` → текст заменён, статус `ready`, без четвёртого батча | unit hook |
| S11 | Команды Testing Strategy зелёные | локально |

## API contract (новый)

```
POST /api/v1/commercial/drafts/{draft_id}/ocr-page
Auth: admin | manager, ownership draft
Content-Type: multipart/form-data
  image: file (обязателен; те же magic JPEG/PNG/PDF и COMMERCIAL_UPLOAD_MAX_BYTES)
```

**200**

```json
{
  "normalized_text": "34-15-10п 15\n...",
  "ocr_verify_failed": false,
  "ocr_corrections": []
}
```

UI retry подставляет `normalized_text` в список страницы и обновляет verify/corrections **этой** страницы. Не раздувать полным `CommercialDraftDetails`.

**Ошибки:** 400 нет/битый файл; 404 draft; 413 размер; 503 OCR выключен; provider 502/503 как у текущего update-with-image.

**Побочных эффектов на draft нет** (текст, batches, `wide_plate_lines`, order_data).

Кап: при `0` этот POST тоже не считает. При `>0` — та же `prepare_commercial_ocr_upload` (считается как загрузка).

## Open Questions

- Имя пути: `/ocr-page` vs `/recognize-page` — **зафиксировано в PLAN:** `POST /api/v1/commercial/drafts/{draft_id}/ocr-page`.
