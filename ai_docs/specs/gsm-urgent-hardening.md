# Spec: Срочное укрепление модуля ГСМ (gitignore + XSS + пересборка HTML)

> Источник: [`ai_docs/develop/audits/2026-08-11-gsm-module-audit.md`](../develop/audits/2026-08-11-gsm-module-audit.md)  
> Пакет: **рекомендуемый** (одобрен 2026-08-11)  
> Связано: [`gsm-routes-map.md`](./gsm-routes-map.md)

## Objective

Закрыть срочные риски личного offline-модуля ГСМ **до шаринга HTML и до случайного коммита данных**:

1. **[S3]** Не допускать попадания содержимого `ГСМ/` в git.
2. **[S1][S2][S4]** Устранить XSS в сгенерированной карте (popup маршрута, список ближайших, popup маркера поиска).
3. Пересобрать `ГСМ/карта_маршрутов.html` из текущего кэша **без** полного регеокода.

**Не входит:** разбиение god-модуля [A2], общий пакет адресов [A3], stale `routes.geojson` [A1], тесты trip_pool [Q7], SRI CDN [S5].

**Пользователь:** Роман (локальный ПК).  
**Успех:** `ГСМ/` игнорируется git; XSS-payload в адресе не исполняется в карте; HTML пересобран offline; тесты зелёные.

## Tech Stack

- Существующий: Python 3.12, `.venv`, `scripts/build_gsm_routes_map.py`, pytest
- Без новых зависимостей
- Экранирование: JS-функция `escapeHtml` в шаблоне карты (или эквивалент DOM/`textContent`)

## Commands

```bash
# Тесты
.venv/bin/python -m pytest tests/test_build_gsm_routes_map.py -q

# Проверка gitignore
git check-ignore -v "ГСМ/пул_поездок.xlsx" "ГСМ/карта_маршрутов.html" "ГСМ/geo_cache/addresses.json"

# Пересборка HTML без сети (кэш адресов/треков)
.venv/bin/python scripts/build_gsm_routes_map.py --offline --out "ГСМ/карта_маршрутов.html"
```

## Project Structure

```
.gitignore                              → правило на весь ГСМ/
scripts/build_gsm_routes_map.py         → escape в write_map_html (JS)
tests/test_build_gsm_routes_map.py      → XSS regression tests
ГСМ/карта_маршрутов.html                → пересборка (локально, в git не попадает)
ai_docs/specs/gsm-urgent-hardening.md   → эта спека
```

Путевые `.xls`, кэши и Excel-пул **остаются на диске**, но не коммитятся.

## Code Style

Экранирование всех пользовательских/файловых строк перед вставкой в HTML:

```javascript
function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
// popupHtml / renderNearest / bindPopup(label) — только через escapeHtml
```

В Python-тестах: фикстура GeoJSON/адреса с payload `<img src=x onerror=alert(1)>` / `"><script>` — в выходном HTML не должно быть «сырого» payload вне экранирования.

## Testing Strategy

- **pytest** в `tests/test_build_gsm_routes_map.py`
- Минимум:
  - [ ] `write_map_html` с вредоносными `адрес_A` / `машина` → в HTML есть `&lt;`/`&amp;` (или эквивалент), нет сырого `<script>` / `onerror=`
  - [ ] в сгенерированном JS присутствует `escapeHtml` (или аналог) и используется в popup/nearest/search
- Ручная: `git check-ignore`; открыть пересобранную карту, поиск не ломается на обычном адресе

## Boundaries

- **Always:** pytest зелёный перед «готово»; игнор всего `ГСМ/`; экранировать popup + nearest + search label
- **Ask first:** менять структуру `geo_cache`; трогать trip_pool; добавлять зависимости; коммитить что-либо из `ГСМ/`
- **Never:** полный регеокод 244 адресов в рамках этой задачи; рефакторинг god-модуля; удаление тестов; коммит ПДн

## Success Criteria

- [ ] В `.gitignore` есть правило, покрывающее весь каталог `ГСМ/`
- [ ] `git check-ignore` подтверждает игнор xlsx/html/geo_cache
- [ ] В шаблоне карты экранируются: popup маршрута, список nearest, label маркера поиска
- [ ] XSS-тесты в pytest проходят
- [ ] `build_gsm_routes_map.py --offline` успешно пишет `ГСМ/карта_маршрутов.html`
- [ ] Поведение карты для нормальных адресов сохраняется (фильтр машин, поиск known address)

## Open Questions

Закрыты выбором пакета:
- gitignore: **весь `ГСМ/`**
- XSS: **S1+S2+S4**
- Пересборка: **да, `--offline`**
- A1 stale cache: **не в этом пакете**
