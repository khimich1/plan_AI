# Handoff: OCR apply-policy + 2× → SPECIFY в новом окне

> **Дата:** 2026-08-22  
> **Ветка:** текущая рабочая (этап B **в working tree, не закоммичен**)  
> **Статус:** Idea ✅ · **Spec ⏳** · Plan ❌ · Implement ❌ (кроме уже лежащего этапа B)  
> **Цель файла:** открыть **новое окно** и сразу писать спеку по `/spec-driven-development`, без повторного ideation.  
> **Не коммитить** без явной просьбы пользователя.

---

## Как стартовать новую сессию (скопируй в первый промпт)

```
/spec-driven-development переходи к спеке

Контекст: прочитай ai_docs/develop/handoffs/2026-08-22-ocr-verify-apply-and-upscale.md
и ai_docs/ideas/ocr-verify-apply-and-upscale.md

Идея 2+4 уже зафиксирована — не запускай idea-refine заново.
Этап B (auto_small_image) уже в коде и его НЕ откатываем.
Эта спека: (2) Apply Verify только при непустом corrections + (4) 2× Lanczos + контраст на мелких.
Не пиши код, пока спека не утверждена. Не коммить без просьбы.
```

### Чеклист агента в новом окне

1. Прочитать **этот** handoff целиком.
2. Прочитать `.cursor/skills/plan-web-context/SKILL.md`.
3. Прочитать `.cursor/skills/spec-driven-development/SKILL.md` — фаза **SPECIFY only**.
4. Источник правды по направлению: [`ai_docs/ideas/ocr-verify-apply-and-upscale.md`](../../ideas/ocr-verify-apply-and-upscale.md).
5. Образец спеки OCR: [`ai_docs/specs/ocr-small-screenshot-verify.md`](../../specs/ocr-small-screenshot-verify.md) (этап B, уже в коде).
6. Вынести ASSUMPTIONS сразу, затем полный spec (6 разделов SDD + Decisions locked).
7. Сохранить в `ai_docs/specs/ocr-verify-apply-and-upscale.md`. **Не** переходить к PLAN, пока пользователь не утвердит спеку.

**Режим:** `/spec-driven-development`, не `/orchestrate` и не `/implement`.

---

## Артефакты

| Артефакт | Путь | Статус |
|----------|------|--------|
| Idea (этап B: force Verify на мелких) | [`ai_docs/ideas/ocr-small-screenshot-verify.md`](../../ideas/ocr-small-screenshot-verify.md) | done |
| Spec этапа B | [`ai_docs/specs/ocr-small-screenshot-verify.md`](../../specs/ocr-small-screenshot-verify.md) | implemented в WT |
| Plan этапа B | [`ai_docs/develop/plans/2026-08-22-ocr-small-screenshot-verify.md`](../plans/2026-08-22-ocr-small-screenshot-verify.md) | VS-101…106 done; S9 = этот пилот |
| **Idea 2+4 (эта сессия)** | [`ai_docs/ideas/ocr-verify-apply-and-upscale.md`](../../ideas/ocr-verify-apply-and-upscale.md) | **locked, писать спеку отсюда** |
| Spec 2+4 | `ai_docs/specs/ocr-verify-apply-and-upscale.md` | **создать** |
| Plan 2+4 | — | после утверждения спеки |

Код этапа B (уже в дереве, uncommitted):

- `core/ocr/verify_policy.py` — `_auto_image_decision`, `short_side_px`, `min_short_side`
- `core/ocr/image_meta.py` — `image_short_side_px` (только чтение size)
- `core/ocr/pipeline.py` — 6 runners прокидывают short side; **Verify всегда подменяет plates**, если `verified_plates` не пустой
- `core/config/settings.py` — `OCR_VERIFY_AUTO_MIN_SHORT_SIDE` default 1000
- тесты: `tests/test_ocr_verify_policy.py`, `test_ocr_image_meta.py`, `test_pile_ocr_pipeline.py`, `test_recognition_pipeline.py`

---

## Что случилось (зачем новая спека)

Пользователь хотел «увеличить resolution/DPI перед OCR». Ideation сузил до: сначала B (Verify на мелких), потом A (2×) если B не поможет.

B влили в working tree. Пилот на **том же** скрине Excel-таблицы плит:

| Факт | Значение |
|------|----------|
| Лог | `short_side_px=416` `verify_decision=auto_small_image` `api_calls=2` |
| Verify | `rows=10` `corrections=0` |
| Поведение пайплайна | `plates = parser_gate(verified_plates)` даже при пустом `corrections` |
| Оценка пользователя | **этот же файл до B (1 вызов) читался лучше** |
| Ошибки хвоста | `ПБ36` → `ПБ 63`; кириллическая **и** → латинская **u**; 10-я строка на фото обрезана |

Вывод: ставка «второй проход ловит марки» на этом кадре опровергнута. Verify подписал чужой/худший список без diff.

Повторный ideation. Пользователь выбрал **2+4**, явно **не** откат B (вариант 1) и не «сначала только 2».

---

## Что уже решено (не переспрашивать)

| ID | Решение |
|----|---------|
| Канал | Веб wizard КП; Telegram out of scope |
| Этап B | **Оставляем** `auto_small_image` |
| Apply-policy (2) | Брать Verify **только** если `verified_plates` не пустой **и** `corrections` не пустой. Иначе оставить Extract, `verify_applied=true`, reason `verify_kept_extract_empty_corrections` (имя можно уточнить в спеке, смысл нет) |
| Не синтезировать diff | Если plates отличаются, а `corrections=[]` — **не** «додумывать» add/remove. Это и есть регресс |
| Препроцесс (4) | Короткая сторона **исходника** `< min_short_side` → 2× Lanczos + лёгкий контраст → PNG в API. Без DPI, без кропа, без SR |
| Порог Verify после 2× | Считать «мелкий» по **исходнику** (416×2=832 всё ещё мелкий → второй вызов остаётся) |
| Scope | Все 6 OCR runners в `core/ocr/pipeline.py`, один helper, не 6 копий |
| UI / промпты / БД | Не менять в MVP |
| Обрезанная 10-я строка | Не обещать восстановить |
| Омоглифы `u`→`и` | **Out** (вариант 3 не выбран) |
| Отказ «переснимите» | **Out** (вариант 5 не выбран) |

Рекомендация из idea по открытому Q: `corrections=[]` и `row_count_on_image ≠ len(extract)` → **всё равно Extract**. Зафиксировать в спеке как assumption, дать пользователю поправить.

---

## Что писать в спеке (содержание, не копипаста)

Шесть блоков SDD + Decisions locked, как в `ocr-small-screenshot-verify.md`.

Обязательно развести два механизма:

```
исходник
  ├── image_short_side_px(original)     → политика auto_small_image
  └── preprocess if short < N            → 2× + contrast → PNG
           │
           ├── Extract
           └── Verify? (B, по исходнику)
                    │
                    └── select_ocr_plates(extract, verify_result)
                         ├── no/empty plates     → extract, verify_failed
                         ├── empty corrections   → extract + kept_extract reason
                         └── else                → parser_gate(verified)
```

Success criteria, которые должны быть в спеке (черновик, можно переименовать):

| # | Критерий |
|---|----------|
| S1 | `corrections=[]` + другой `plates` → в результате Extract, не Verify |
| S2 | непустой `corrections` + непустой `plates` → Verify проходит parser_gate |
| S3 | пустой `verified_plates` → extract, `verify_failed` (как сейчас) |
| S4 | `never` / `max_api_calls=1` — без Verify, без подмены (как сейчас) |
| S5 | исходник short ≥ порога → препроцесс no-op, исходные байты |
| S6 | исходник short < порога → в провайдер уходит PNG, исходный файл на диске не перезаписан |
| S7 | все 6 пайплайнов через один `select_ocr_plates` / один preprocess |
| S8 | pytest pack этапа B + новые тесты — green; без live GigaChat в CI |
| S9 | тот же скрин 416 px: хвост не хуже, чем Extract-only до B |

Контраст: в idea открыт Autocontrast vs линейный. Спека должна выбрать default (рекомендация: `ImageOps.autocontrast` с cutoff 0–1, без бинаризации) и пометить калибровку как Ask first.

---

## Ключевые файлы (куда лезть при SPECIFY / позже IMPLEMENT)

```
core/ocr/pipeline.py            # 6×: после verify сейчас plates = gate(verified) безусловно (~строка 163)
core/ocr/verify_policy.py       # B уже здесь; 2+4 не обязаны менять should_run_*
core/ocr/image_meta.py          # сейчас только size; preprocess — сюда или sibling
core/ocr/result.py              # build_result_payload: прокинуть reason/preprocess metadata
core/config/settings.py         # min_short_side уже есть; новых env по возможности не плодить
core/ocr/providers/openai.py    # load_image_payload; PDF → short_side None
tests/test_ocr_verify_policy.py
tests/test_ocr_image_meta.py
tests/test_pile_ocr_pipeline.py
tests/test_recognition_pipeline.py
```

Не трогать в этой спеке: frontend wizard, промпты Extract/Verify, parser_gate логику марок, upload validation.

---

## Долг ревью этапа B (не блокирует спеку 2+4)

Review сказал Request changes по **тестам B**, не по прод-логике B:

1. Нет теста, что пайплайн читает `ocr_verify_auto_min_short_side` из Settings (`=0` → skip на 400×300).
2. Нет pipeline-теста D5: PDF/битые байты → `auto_image_size_unknown`.
3. `test_recognize_text_smart_small_image_runs_verify` не пинит `OCR_VERIFY_MODE=auto`.

Можно упомянуть в Open Questions спеки 2+4 («закрыть долг B в том же PR?») — не смешивать в MVP 2+4 без спроса.

---

## Open Questions (вынести в ASSUMPTIONS спеки)

- `corrections=[]` + другой `row_count_on_image` → Extract? (idea: да)
- Autocontrast vs линейный contrast
- Лимит GigaChat upload после 2× PNG
- Metadata `ocr_preprocess: 2x_lanczos` для пилота — да/нет
- Долг тестов B в этом же PR?

---

## Анти-паттерны для новой сессии

- Не предлагать снова «поднять DPI».
- Не предлагать Real-ESRGAN / автокроп / откат B, пока пользователь сам не вернётся.
- Не писать PLAN/код в том же ходе, что и черновик спеки, без «спека ок».
- Не коммитить этап B «заодно».
- Не считать UI-строку «исправлено 4 строк(и) и→u» источником истины: серверный лог того прогона — `corrections=0`. Источник регресса — **подмена plates**.
