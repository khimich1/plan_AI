# Plan: Срочное укрепление модуля ГСМ

**Created:** 2026-08-11  
**Status:** ✅ Done  
**Spec:** [`ai_docs/specs/gsm-urgent-hardening.md`](../../specs/gsm-urgent-hardening.md) (✅ approved)  
**Audit:** [`ai_docs/develop/audits/2026-08-11-gsm-module-audit.md`](../audits/2026-08-11-gsm-module-audit.md)

## Goal

Закрыть [S3] gitignore на весь `ГСМ/`, [S1][S2][S4] XSS в сгенерированной карте, пересобрать HTML offline. Без A1/A2/A3 и регеокода.

## Current state

| Компонент | Сейчас |
|-----------|--------|
| `.gitignore` | Нет правила на `ГСМ/` |
| `write_map_html` | `popupHtml` / `renderNearest` / `bindPopup(label)` без escape |
| Тесты | Нет XSS regression |
| HTML | `ГСМ/карта_маршрутов.html` собран до XSS-фиксов |

## Architecture decisions

1. **gitignore:** одна строка `ГСМ/` (весь каталог, как в спеке).
2. **XSS:** одна JS-функция `escapeHtml` внутри шаблона `write_map_html`; применить в popup маршрута, списке nearest, label маркера поиска. Не выносить JS во внешний файл в этом пакете.
3. **Тесты:** фикстура с payload в свойствах → assert экранирования в выходном HTML.
4. **Пересборка:** `build_gsm_routes_map.py --offline` после правок кода.

## Phases & Tasks

### Phase 1 — защита от коммита ПДн
- [x] T1. Добавить `ГСМ/` в `.gitignore`
  - Acceptance: `git check-ignore -v` игнорирует xlsx/html/geo_cache
  - Verify: `git check-ignore -v "ГСМ/пул_поездок.xlsx" "ГСМ/карта_маршрутов.html" "ГСМ/geo_cache/addresses.json"`
  - Files: `.gitignore`

### Phase 2 — XSS
- [x] T2. `escapeHtml` + применение в popup / nearest / search marker
  - Acceptance: все три места экранируют строки; обычные адреса отображаются читаемо
  - Verify: code review диффа + pytest T3
  - Files: `scripts/build_gsm_routes_map.py`
- [x] T3. XSS regression tests
  - Acceptance: payload в `адрес_A`/`машина` не попадает в HTML «сырым»; есть `escapeHtml` (или эквивалент)
  - Verify: `.venv/bin/python -m pytest tests/test_build_gsm_routes_map.py -q`
  - Files: `tests/test_build_gsm_routes_map.py`

### Phase 3 — артефакт
- [x] T4. Offline-пересборка HTML
  - Acceptance: `ГСМ/карта_маршрутов.html` обновлён; exit 0; в файле есть экранирование
  - Verify: `--offline` run + grep/escapeHtml в HTML
  - Files: `ГСМ/карта_маршрутов.html` (локально, в git не попадёт)

## Verification checkpoints

- После T1: gitignore работает
- После T3: все тесты routes_map зелёные
- После T4: карта открывается, поиск known address работает

## Out of scope

[A1] stale routes cache, [A2] god module, [A3] shared address package, полный регеокод, SRI CDN, trip_pool exit codes.
