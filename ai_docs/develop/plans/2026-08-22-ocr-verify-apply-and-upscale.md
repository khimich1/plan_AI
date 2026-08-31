# Plan: OCR — Apply Verify только при corrections + 2× на мелких

> **Spec:** [`ai_docs/specs/ocr-verify-apply-and-upscale.md`](../../specs/ocr-verify-apply-and-upscale.md)  
> **Идея:** [`ai_docs/ideas/ocr-verify-apply-and-upscale.md`](../../ideas/ocr-verify-apply-and-upscale.md)  
> **Handoff:** [`ai_docs/develop/handoffs/2026-08-22-ocr-verify-apply-and-upscale.md`](../handoffs/2026-08-22-ocr-verify-apply-and-upscale.md)  
> **Предшественник:** [`2026-08-22-ocr-small-screenshot-verify.md`](./2026-08-22-ocr-small-screenshot-verify.md) (этап B в working tree — **не откатывать**)  
> **Дата:** 2026-08-22  
> **Статус:** IMPLEMENT ✅ (AU-201…209) — AU-210 manual S9 pending  
> **Не коммитить** без явной просьбы (в т.ч. этап B «заодно»)

---

## Overview

Два независимых механизма на уже существующем этапе B (`auto_small_image`):

1. **Apply-policy** — после Verify брать список модели только если `plates` не пустой **и** `corrections` не пустой. Иначе оставить Extract. Не синтезировать diff.
2. **Препроцесс** — если короткая сторона *исходника* `< min_short_side`, в API уходит 2× Lanczos + `autocontrast(cutoff=1)` как PNG. Порог Verify по-прежнему по исходнику.

Этап B, UI, промпты, parser_gate, новые env — вне этого плана. Долг тестов B — не в этом PR.

---

## Approach

Не менять `should_run_*_verify`. Политика «когда звать Verify» уже верная; ломается «что делать с ответом» и «какие пиксели видит модель».

```
image_path
  ├── image_short_side_px(original)              → политика B (без правок)
  ├── original_bytes                             → image_size_bytes для политики
  └── preprocess_image_for_ocr(path, min_short)
           │
           ├── applied → PNG bytes + image/png
           └── else    → original bytes + исходный mime
                    │
                    ├── Extract(payload)
                    └── Verify?(payload, порог по исходнику)
                             │
                             └── select_ocr_items(extract, verify_result, gate)
                                  ├── empty plates     → extract, failed, empty_verified_plates
                                  ├── empty corrections→ extract, kept_extract_empty_corrections
                                  └── else             → gate(verified), applied
```

В `pipeline.py` два helper'а на все 6 runners — не копировать if в каждый:

- `_ocr_image_payload(path, min_short_side)` → original_bytes, payload_bytes, mime, short_side, preprocess_flag
- после Verify: `select_ocr_items(...)` вместо `if not verified: fail; else: plates = gate(verified)`

`ocr_verify_applied_reason` остаётся причиной *запуска*. Новые поля payload: `ocr_verify_select_reason`, `ocr_preprocess`.

---

## Architecture Decisions

- Apply живёт в `core/ocr/verify_apply.py`, не в `verify_policy.py` (когда звать ≠ что взять).
- Препроцесс — sibling `core/ocr/image_preprocess.py`; `image_meta.py` только читает size.
- Ключ Verify-списка везде `"plates"` — как сейчас у свай/ступеней.
- Политика считает исходник (D7/D8), чтобы 2× PNG не сменил reason на `auto_file_too_large`.
- PNG > 8 MiB → no-op, исходник, лог skip. Нового env нет.
- Draft metadata получает оба новых ключа; frontend copy не трогаем.
- TDD: RED unit → GREEN unit → wiring. Пилот S9 не блокер кода.

---

## Components

| Компонент | Роль | Зависит от |
|-----------|------|------------|
| `select_ocr_items` | apply-policy, 3 ветки | — |
| `preprocess_image_for_ocr` | 2× + contrast или no-op | `min_short_side`, Pillow |
| `build_result_payload` | `ocr_verify_select_reason`, `ocr_preprocess` | — |
| `_ocr_image_payload` в pipeline | один load+preprocess на 6 runners | preprocess, image_meta |
| 6 runners | payload в провайдер; select после Verify | всё выше |
| draft mapper + schema | ключи доходят до черновика | payload |
| `ocr_pilot_compare.py` | печать полей для S9 | payload |

Внешние API, промпты, frontend UI, upload validation, `verify_policy.py` — не в этом плане.

---

## Dependency graph

```
AU-201 (RED apply) ──► AU-202 (GREEN apply) ──┐
                                               │
AU-203 (RED preprocess) ► AU-204 (GREEN pp) ──┤
                                               ├──► AU-206 (pipeline wiring)
AU-205 (payload fields) ───────────────────────┤         │
                                               │         ├──► AU-207 (draft mapper)
                                               │         ├──► AU-208 (pilot script)
                                               │         └──► AU-209 (S8 pack) ► AU-210 (S9 manual)
```

**Можно параллелить:** AU-201/202 и AU-203/204; AU-205 с любой из этих веток.  
**Только последовательно:** AU-206 после 202+204+205; AU-209 после wiring.

---

## Tasks

- [x] **AU-201:** RED — unit `select_ocr_items` (S1–S3)
  - Acceptance: тесты описаны и падают: пустой corrections + другой plates → extract; непустой corrections → `apply_gate(verified)` (gate зовётся); пустой verified → extract + `verify_failed`; другой `row_count_on_image` + `corrections=[]` → extract
  - Verify: `pytest tests/test_ocr_verify_apply.py -q` — fail (нет модуля / нет функции)
  - Files: `tests/test_ocr_verify_apply.py`
  - Scope: S · Dependencies: none

- [x] **AU-202:** GREEN — `select_ocr_items`
  - Acceptance: три `select_reason`; `verify_failed` только на пустом verified; gate не зовётся на empty corrections; не сравниваем plates поэлементно
  - Verify: `pytest tests/test_ocr_verify_apply.py -q`
  - Files: `core/ocr/verify_apply.py`, `tests/test_ocr_verify_apply.py`
  - Scope: S · Dependencies: AU-201

- [x] **AU-203:** RED — unit preprocess (S5–S6)
  - Acceptance: тесты падают: `short >= N` и `== N` → исходные байты, `applied=False`; `short < N` → PNG, геометрический 2×; `min_short_side=0` no-op; битый файл / PDF-like → `None`; исходник на диске байт-в-байт тот же; RGBA/P не падают
  - Verify: `pytest tests/test_ocr_image_preprocess.py -q` — fail
  - Files: `tests/test_ocr_image_preprocess.py`
  - Scope: S · Dependencies: none · **parallelSafe** с AU-201

- [x] **AU-204:** GREEN — `preprocess_image_for_ocr`
  - Acceptance: RGB → 2× `LANCZOS` → `ImageOps.autocontrast(..., cutoff=1)` → PNG; диск не писать; `> 8 MiB` encode → no-op + не падать (тест можно на mock encode, не на гигантском PNG в git)
  - Verify: `pytest tests/test_ocr_image_preprocess.py -q`
  - Files: `core/ocr/image_preprocess.py`, `tests/test_ocr_image_preprocess.py`
  - Scope: S · Dependencies: AU-203

- [x] **AU-205:** payload — новые поля в `build_result_payload`
  - Acceptance: kwargs `verify_select_reason=None`, `ocr_preprocess=None`; ключи в dict; старые вызовы без kwargs не ломаются
  - Verify: точечный тест в существующем файле result/pipeline **или** мини-тест рядом с AU-201; импорт `build_result_payload` без новых обязательных аргументов
  - Files: `core/ocr/result.py`, при необходимости `tests/test_recognition_pipeline.py` (только вызов payload, без wiring)
  - Scope: S · Dependencies: none · **parallelSafe** с AU-201/203

- [x] **AU-206:** Pipeline wiring (6 runners)
  - Acceptance: один `_ocr_image_payload` + `select_ocr_items` во всех 6; политика получает **original** bytes/short_side; провайдер на мелком кадре — `image/png`; лог `preprocess=` и `select=`; S1 на плитах (другой verified + `corrections=[]` → extract); S2 на сваях (непустой corrections → gate); `never` без подмены. Поправить `_mock_provider`: тест «починил ???» передаёт непустой `corrections`
  - Verify: `pytest tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py tests/test_ocr_verify_apply.py tests/test_ocr_image_preprocess.py -q`
  - Files: `core/ocr/pipeline.py`, `tests/test_recognition_pipeline.py`, `tests/test_pile_ocr_pipeline.py`
  - Scope: M · Dependencies: AU-202, AU-204, AU-205

- [x] **AU-207:** Draft metadata mapper + schema
  - Acceptance: `ocr_verify_select_reason`, `ocr_preprocess` в `_map_ocr_result_metadata` и `CommercialDraftMetadata` (`str | None = None`). UI/frontend types не обязательны
  - Verify: `pytest tests/test_commercial_ocr_policy.py -q` + импорт схемы; если есть удобный draft-тест — добавить два ключа
  - Files: `app/services/commercial_draft_service.py`, `app/schemas/commercial.py`
  - Scope: S · Dependencies: AU-205 (логически), AU-206 (чтобы ключи реально приходили из pipeline)

- [x] **AU-208:** Пилот-скрипт печатает новые поля
  - Acceptance: `ocr_preprocess` и `ocr_verify_select_reason` в stdout; CLI без новых флагов
  - Verify: чтение скрипта / `python -m py_compile scripts/ocr_pilot_compare.py`
  - Files: `scripts/ocr_pilot_compare.py`
  - Scope: XS · Dependencies: AU-205

- [x] **AU-209:** Regression pack (S8)
  - Acceptance: pack зелёный, без live API
  - Verify: `pytest tests/test_ocr_verify_apply.py tests/test_ocr_image_preprocess.py tests/test_ocr_verify_policy.py tests/test_ocr_image_meta.py tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py tests/test_commercial_ocr_policy.py -q`
  - Files: нет, если зелёный; иначе точечный фикс в файле регрессии
  - Scope: S · Dependencies: AU-206, AU-207, AU-208

- [ ] **AU-210:** Manual S9 (не блокер кода)
  - Acceptance: тот же скрин 416 px; хвост не хуже Extract-only до B; при `corrections=0` список = Extract; в выводе `ocr_preprocess=2x_lanczos` и `kept_extract_empty_corrections` (если модель снова не дала diff). Обрезанную строку не оцениваем
  - Verify: `python scripts/ocr_pilot_compare.py --image <пилот-416> --verify-mode auto`
  - Files: нет
  - Scope: — · Dependencies: AU-209

---

## Verification checkpoints

| After | Command | Expected |
|-------|---------|----------|
| AU-201 | `pytest tests/test_ocr_verify_apply.py -q` | RED |
| AU-202 | тот же | green (S1–S3 unit) |
| AU-203 | `pytest tests/test_ocr_image_preprocess.py -q` | RED |
| AU-204 | тот же | green (S5–S6) |
| AU-206 | pipeline + apply + preprocess tests | green, без live API |
| AU-209 | pack S8 | green |
| AU-210 | пилот-скрипт | ручной, не CI |

После AU-202+204 система ещё не меняет прод-поведение (helpers есть, runners старые). Поведение меняется на AU-206 — это главный checkpoint перед mapper/скриптом.

---

## Risks

| Риск | Почему | Mitigation |
|------|--------|------------|
| Забыть 1 из 6 runners | `pipeline.py` — шесть копий | `_ocr_image_payload` + один `select_ocr_items`; grep `plates = apply_` / `piles = apply_` после wiring |
| Политика видит PNG-байты | 2× раздует файл → другой reason | в `should_run_*` только `len(original_bytes)` и исходный `short_side_px` |
| Существующие mock-тесты «Verify починил» | `_mock_provider` всегда `corrections=[]` | в AU-206 явно добавить correction; это новое правильное поведение, не ослаблять тест |
| Мутация upload-файла | Pillow `save(path)` | только bytes в памяти; тест «диск не изменился» |
| RGBA / palette / PDF | падение preprocess | convert RGB; unreadable → `None` → исходник |
| PNG > лимита GigaChat | редкий огромный скрин | потолок 8 MiB → исходник + лог; без нового env |
| Смешать этап B в коммит | B уже в WT uncommitted | не коммитить, пока пользователь не попросит; не откатывать B |

---

## Out of scope

Откат B; долг тестов B; омоглифы; автокроп/DPI/SR; смена промптов и UI; рендер PDF; новые env; live vision в CI; Telegram; коммит без просьбы.

---

## Open Questions (не блокируют PLAN)

- Путь к пилотному файлу 416 px для AU-210 — уточнить перед ручным прогоном.
- Долг тестов B по-прежнему Ask first (в этот план не входит).
