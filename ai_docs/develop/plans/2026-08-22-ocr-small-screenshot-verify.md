# Plan: OCR мелких скриншотов — Verify по короткой стороне

> **Spec:** [`ai_docs/specs/ocr-small-screenshot-verify.md`](../../specs/ocr-small-screenshot-verify.md)  
> **Идея:** [`ai_docs/ideas/ocr-small-screenshot-verify.md`](../../ideas/ocr-small-screenshot-verify.md)  
> **Дата:** 2026-08-22  
> **Статус:** IMPLEMENT ✅ (VS-101…106) — VS-107 manual S9 pending  
> **Этап A (2× / контраст):** out of scope

---

## Approach

Не трогаем пиксели и UI. Читаем `min(width, height)` через Pillow, передаём в уже существующую `auto`-политику Verify. Мелкий чистый скрин, который сегодня даёт `auto_all_checks_passed`, получит второй вызов с reason `auto_small_image`.

Источник правды по решению — `core/ocr/verify_policy.py`. Пайплайн только измеряет кадр и логирует.

```
load_image_payload
    │
    ├── image_short_side_px(path)  → int | None
    │
    └── should_run_*_verify(..., short_side_px=..., settings.min_short_side)
              │
              ├── never / max_api_calls=1  → skip
              ├── always                   → verify
              ├── empty rows               → verify
              ├── _auto_image_decision     → too_large | unknown | small_image
              └── row/confidence/parse     → как сейчас
```

---

## Components

| Компонент | Роль | Зависит от |
|-----------|------|------------|
| `OcrVerifySettings.min_short_side` | порог, `0` = выкл | — |
| `_auto_image_decision` | bytes / unknown / small | settings |
| `should_run_*_verify` × 6 | новый kwarg `short_side_px` | helper |
| `Settings.ocr_verify_auto_min_short_side` | env | — |
| `image_short_side_px` | Pillow, без мутации файла | — |
| `pipeline.py` × 6 runners | измерить, прокинуть, залогировать | всё выше |

Внешние API, промпты, frontend, upload validation — не в этом плане.

---

## Implementation order

Последовательность обязательна (TDD). Параллелить почти нечего: один модуль политики.

```
1. Tests (RED)  → policy helper + short_side_px API
2. Policy       → _auto_image_decision + 6 функций
3. Settings     → env default 1000
4. image_meta   → short_side из файла
5. Pipeline     → 6 call sites + log
6. Regression   → pytest pack из спеки S8
7. Pilot S9     → вручную, не блокер merge кода
```

---

## Tasks

- [x] **VS-101:** RED — тесты политики на `short_side_px`
  - Acceptance: S1–S5 описаны тестами и падают; существующие skip-кейсы передают `short_side_px=2000` (или ≥ default); `short_side_px` — обязательный kwarg (нет тихого skip)
  - Verify: `pytest tests/test_ocr_verify_policy.py -q` — fail на новых тестах до VS-102
  - Files: `tests/test_ocr_verify_policy.py`

- [x] **VS-102:** GREEN — `_auto_image_decision` + проводка в 6 `should_run_*_verify`
  - Acceptance: S1–S6; порядок reason как в спеке; `min_short_side=0` не форсит; `None` → `auto_image_size_unknown`; `never` / `max_api_calls=1` важнее
  - Verify: `pytest tests/test_ocr_verify_policy.py -q`
  - Files: `core/ocr/verify_policy.py`, `tests/test_ocr_verify_policy.py`

- [x] **VS-103:** Settings + `.env.example`
  - Acceptance: `OCR_VERIFY_AUTO_MIN_SHORT_SIDE` default `1000`, `ge=0`; пайплайн позже читает поле
  - Verify: импорт `get_settings()` / существующие settings-тесты, если есть
  - Files: `core/config/settings.py`, `.env.example`

- [x] **VS-104:** `image_short_side_px`
  - Acceptance: PNG с известным size → `min(w,h)`; битые байты / PDF-like → `None`; файл не перезаписывается
  - Verify: `pytest tests/test_ocr_image_meta.py -q` (картинки через PIL в `tmp_path`, без фикстур в git)
  - Files: `core/ocr/image_meta.py` (новый), `tests/test_ocr_image_meta.py` (новый)

- [x] **VS-105:** Pipeline wiring
  - Acceptance: все 6 runners: `OcrVerifySettings(min_short_side=cfg.ocr_verify_auto_min_short_side)`, `short_side_px=image_short_side_px(image_path)`, лог `short_side_px`; не ресайзят изображение. Вынести сборку settings в один helper в `pipeline.py`, чтобы не забыть седьмой копии
  - Verify: `pytest tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py -q`; точечный тест, что plates-pipeline прокидывает short side (mock policy или fixture PNG)
  - Files: `core/ocr/pipeline.py`, существующие pipeline-тесты при необходимости

- [x] **VS-106:** Regression pack (S8)
  - Verify: `pytest tests/test_ocr_verify_policy.py tests/test_ocr_image_meta.py tests/test_recognition_pipeline.py tests/test_pile_ocr_pipeline.py tests/test_commercial_ocr_policy.py -q`
  - Files: нет, если зелёный

- [ ] **VS-107:** Manual S9 (после merge, не блокер кода)
  - Acceptance: 10–15 плохих скринов: `ocr_verify_applied_reason`, правки марок/строк vs baseline `auto`
  - Verify: вне CI; решение «нужен ли этап A»

---

## Verification checkpoints

| After | Command | Expected |
|-------|---------|----------|
| VS-101 | `pytest tests/test_ocr_verify_policy.py -q` | новые тесты RED |
| VS-102 | тот же | green |
| VS-104 | `pytest tests/test_ocr_image_meta.py -q` | green |
| VS-105 | pipeline tests | green, без live API |
| VS-106 | pack S8 | green |

---

## Risks

| Риск | Почему | Mitigation |
|------|--------|------------|
| Забыть один из 6 call sites | `pipeline.py` — шесть копий | helper сборки settings; обязательный `short_side_px` (забыл → TypeError в тестах/проде, не тихий skip) |
| Skip-тесты политики начнут ждать Verify | default `None` = unknown | во всех skip-кейсах явный `short_side_px=2000` |
| Default 1000 слишком агрессивен | Full HD скрины начнут всегда Verify | env; пилот S9; Ask first на смену default |
| PIL и PDF | upload допускает PDF | `None` → `auto_image_size_unknown` |
| Лишние ₽ / latency | каждый мелкий скрин = 2 вызова | объём 50–200/мес; `never` / `MAX_API_CALLS=1` остаются выключателем |

---

## Out of scope

Апскейл, DPI, кроп, SR, table-OCR, UI, разные пороги по `product_type`, live vision в CI, Telegram.
