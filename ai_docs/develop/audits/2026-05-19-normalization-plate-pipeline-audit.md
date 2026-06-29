# Консолидированный аудит: пайплайн «Нормализация / распознавание плит»

| Поле | Значение |
|------|----------|
| **Дата** | 2026-05-19 |
| **Область (Scope)** | Пайплайн «Нормализация / распознавание плит» |
| **Аудиторы** | senior-reviewer, security-auditor, reviewer (Composer 2.5) |
| **Health Score** | **2.0 / 10** |

### Файлы в scope

| Модуль | Путь |
|--------|------|
| Нормализация текста | `core/plate_text_normalizer.py` |
| Парсинг строки | `core/plate_line_parser.py` |
| Сервисный пайплайн | `app/services/plate_parser_service.py` |
| Legacy-парсинг списков | `core/parsing/plate_lists.py` |
| Константы и конфиг | `core/config_and_data.py` |
| Парсинг в БД КП | `core/kp_db.py` (участки парсинга) |
| OCR GPT Vision | `core/ocr_gpt.py` |
| Telegram OCR/ввод | `bot/handlers/commercial.py` |
| Тесты | `tests/test_recognition_pipeline.py`, `tests/test_plate_normalizer.py` |

---

## 1. Executive Summary

Пайплайн нормализации и распознавания плит — критический путь коммерческого бота и веб-API: от сырого текста/OCR до структурированного заказа с распределением по ширинам, нагрузкам и номенклатуре. Аудит выявил **системную архитектурную фрагментацию**: два (фактически три) независимых контура парсинга с **разной бизнес-семантикой**, что напрямую влияет на корректность заказов, КП и расчётов.

**Health Score: 2.0 / 10** — состояние требует срочной консолидации перед добавлением новых форматов плит или расширением OCR.

### Сводка по серьёзности

| Категория | Critical | High | Medium | Low | Всего |
|-----------|----------|------|--------|-----|-------|
| Architecture (A*) | 2 | 4 | 6 | 4 | 16 |
| Security (S*) | 0 | 2 | 6 | 3 | 11 |
| Code Quality (Q*) | 0 | 5 | 11 | 8 | 24 |
| **Итого** | **2** | **11** | **23** | **15** | **51** |

**Расчёт Health Score:**

```
10 − min(2×2, 6) − min(11×0.5, 3) − min(23×0.1, 1) = 10 − 4 − 3 − 1 = 2.0
```

| Метрика | Значение |
|---------|----------|
| Критические находки | 2 |
| Высокий приоритет | 11 |
| Средний приоритет | 23 |
| Низкий приоритет | 15 |
| **Всего находок** | **51** |

---

## 2. Рекомендация

**Единый use-case «ParsePlateOrder»** должен стать единственной точкой входа для текста и OCR-результата: `normalize → parse_line → distribute_widths → fill_nomenclature`. Legacy-функции в `plate_lists.py` и дублирующая логика в `PlateParserService.add_items` следует свести к одной реализации распределения по ширинам (включая правило 1.5 м → 1.2 + 0.3), вынести математику размеров в `core/domain/plate_dimensions.py`, а оркестрацию из `commercial.py` — в сервисный слой. Параллельно закрыть пробелы безопасности Telegram-пути (rate limit, лимит размера изображения, санитизация ошибок). Без этой консолидации любое исправление в одном контуре будет **молча расходиться** с другим — риск неверных заказов выше, чем риск падения приложения.

---

## 3. Critical Issues

### [A1] Два параллельных пайплайна парсинга — расхождение заказов

| | |
|---|---|
| **Категория** | Architecture — Critical |
| **Локация** | `core/parsing/plate_lists.py` (`parse_text_to_plate_lists`, `add_items`) vs `app/services/plate_parser_service.py` (`PlateParserService.parse_order_text`) |
| **Суть** | Бот и `kp_db` по-прежнему используют legacy-путь через `plate_lists`; API и новый код — `PlateParserService`. Оба вызывают `parse_line`, но **последующая бизнес-логика различается** (распределение ширин, fallback, учёт нагрузок, wide lines). |
| **Impact** | Один и тот же текст заказа даёт **разный состав плит** в Telegram vs веб vs сохранённом КП; невоспроизводимые баги, потеря доверия к системе, сложность поддержки. |
| **Fix** | 1) Выделить единый `PlateOrderParser` (domain service). 2) Перевести `plate_lists`, `kp_db`, `commercial.py` на него. 3) Deprecate дублирующие `add_items`. 4) Добавить golden-тесты: один вход → один выход для всех потребителей. |

### [A2] Разная семантика распределения по ширинам (legacy split 1.5 м vs API fallback в 1_2)

| | |
|---|---|
| **Категория** | Architecture — Critical |
| **Локация** | `core/parsing/plate_lists.py:155–184` (1.45–1.55 м → `plates_1_2` + `plates_0_32`) vs `app/services/plate_parser_service.py:111–113` (неизвестная ширина → fallback `plates_1_2` без split) |
| **Суть** | Legacy **явно раскладывает** плиту 1.5 м на 1.2 + 0.3. Сервисный путь **не содержит** ветки 1.5 м и при нераспознанной ширине молча кладёт всё в `plates_1_2`. |
| **Impact** | Плиты 1.5 м в API/новом пути попадают в одну корзину 1.2 м — **неверный заказ, цена и производство**. Обратная ситуация: legacy даёт две позиции, сервис — одну. |
| **Fix** | Единая функция `distribute_plate_by_width(width_m, length_m, qty, …)` с таблицей правил из `config_and_data` / domain constants; unit-тесты на границы 1.45–1.55, 0.88, «меньший рез». Удалить fallback «всё в 1_2» или сделать его явным с логированием и `unrecognized_lines`. |

---

## 4. High Priority

### Архитектура

| ID | Проблема | Локация | Fix |
|----|----------|---------|-----|
| **A3** | Дублирование математики размеров | `core/plate_line_parser.py` (WxL, дм→м) vs `core/config_and_data.py` / константы | Вынести в `core/domain/plate_dimensions.py`; импортировать из одного модуля |
| **A4** | Оркестрация в handler, не use-case | `bot/handlers/commercial.py` (~2000+ строк): normalize, OCR, parse, save | `CommercialOrderUseCase` / `PlateRecognitionWorkflow` в `app/services/` |
| **A5** | Два доменных `PlateOrder` | `app/domain/` vs `core/domain/` + runtime-структуры в `plate_lists` | Один канонический DTO; мапперы на границах legacy |
| **A6** | Два механизма номенклатуры | DI filler в app vs прямой `sqlite3` в `PlateParserService` | Единый `NomenclatureRepository` через DI |

### Безопасность

| ID | Проблема | Локация | Fix |
|----|----------|---------|-----|
| **S1** | Нет rate limiting OCR в Telegram | `bot/handlers/commercial.py` (фото); веб: ~10/час в API | Общий `OcrRateLimiter` (Redis/DB), тот же лимит для бота |
| **S2** | Нет лимита размера изображения в боте | Загрузка фото → RAM → `core/ocr_gpt.py` | `MAX_OCR_IMAGE_BYTES` до `read()`; отказ с UX-сообщением |

### Качество кода

| ID | Проблема | Локация | Fix |
|----|----------|---------|-----|
| **Q1** | Два пайплайна с разным поведением | См. [A1] | Консолидация + regression suite |
| **Q2** | Недостижимые ветки ширины в `add_items` | `plate_parser_service.py:98–109` (`0.74`, `0.34` перекрываются / недостижимы при текущих диапазонах) | Удалить мёртвые ветки или выровнять с `plate_lists` |
| **Q3** | Расхождение WxL: `parse_line` vs `get_wide_plate_lines` | `plate_line_parser.py` vs `plate_text_normalizer.py` | Один парсер геометрии; wide lines — фильтр над результатом |
| **Q4** | Тройное дублирование математики размеров | `plate_line_parser`, `plate_text_normalizer`, `plate_lists` | См. [A3] |
| **Q5** | Молчаливый fallback `load=8.0` в `parse_catalog_mark` | `plate_line_parser.py` / связанные хелперы | Явный `LoadCode.UNKNOWN`; warning в лог; не подставлять 8.0 без маркировки |

---

## 5. Medium Priority (по темам)

### Пайплайн и архитектура

| ID | Описание | Локация |
|----|----------|---------|
| **A7** | `get_wide_plate_lines` — третья копия детекции форматов | `core/plate_text_normalizer.py` |
| **A8** | OCR не в доменном контракте; двойная нормализация (OCR → normalize → parse снова) | `ocr_gpt.py`, `commercial.py`, `plate_parser_service.py` |
| **A9** | `normalize` не зафиксирован как обязательный pre-step для `parse_line` | Документация + assert/валидатор на входе парсера |
| **A10** | `PlateParserService` — god-method (`parse_order_text` + inline `add_items` + sqlite) | `app/services/plate_parser_service.py` |
| **A12** | Двойной парсинг при текстовом вводе в боте | `bot/handlers/commercial.py` |

### OCR и безопасность

| ID | Описание | Локация |
|----|----------|---------|
| **S3** | PII на фото уходит в OpenAI | `core/ocr_gpt.py` |
| **S4** | `str(e)` пользователю в Telegram | `commercial.py:292, 1361, 1747` |
| **S5** | Debug agent-log на диск | `core/kp_db.py`, `core/config_and_data.py` |
| **S6** | Нет лимита длины текстового ввода | `commercial.py`, handlers ввода |
| **S7** | Prompt injection через OCR (LLM integrity) | Промпт + пост-валидация `parse_line` |
| **S8** | `image_path` без валидации | `ocr_gpt.recognize_with_gpt_vision` |

### Тесты

| ID | Описание | Локация |
|----|----------|---------|
| **A11** | Слабые/разрозненные тесты; нет паритета legacy vs service | `tests/test_recognition_pipeline.py`, `tests/test_plate_normalizer.py` |

### Качество кода (Q6–Q16)

| ID | Описание | Локация |
|----|----------|---------|
| **Q6** | Монолит `get_wide_plate_lines` | `plate_text_normalizer.py` |
| **Q7** | `unrecognized_lines` всегда `[]` | `plate_parser_service.py` |
| **Q8** | Broad `except` без типов | `plate_lists.py`, handlers |
| **Q9** | Мёртвая переменная `parsed` в `plate_lists` | `plate_lists.py` |
| **Q10** | Слабые/неполные аннотации типов | Несколько модулей scope |
| **Q11** | Дубли нормализации в цепочке | `commercial.py` + service |
| **Q12** | Длинный handler (>2000 строк) | `commercial.py` |
| **Q13** | Agent log в production-пути | `kp_db.py`, `config_and_data.py` |
| **Q14** | `print` / traceback в OCR | `core/ocr_gpt.py` |
| **Q15** | Хрупкий JSON regex для ответа GPT | `ocr_gpt.py` |
| **Q16** | MIME `image/jpeg` всегда, без определения формата | `ocr_gpt.py` |

---

## 6. Low Priority / Suggestions

### Архитектура

| ID | Описание | Локация |
|----|----------|---------|
| **A13** | Lazy import + тихий `except` | `plate_lists.py` |
| **A14** | `unrecognized_lines` всегда пустой (не заполняется) | `plate_parser_service.py` |
| **A15** | Мёртвая ветка hybrid в OCR | `ocr_gpt.py` |
| **A16** | `kp_db` зависит от legacy `config_and_data` для имён | `core/kp_db.py` |

### Безопасность

| ID | Описание | Локация |
|----|----------|---------|
| **S9** | Временные фото OCR в боте не удаляются | `commercial.py` |
| **S10** | `print` / traceback в production | `ocr_gpt.py` |
| **S11** | SQL `LIKE` без escape `%` / `_` | `kp_db.py` (поиск) |

### Качество кода (Q17–Q24)

| ID | Тема |
|----|------|
| **Q17** | Магические числа ширины вместо именованных констант |
| **Q18** | Дублирование regex-паттернов между normalizer и parser |
| **Q19** | Отсутствие `__all__` / публичного API модуля normalizer |
| **Q20** | Несогласованные имена списков (`PLATES_1_2` vs `plates_1_2`) |
| **Q21** | Логирование на русском и английском вперемешку |
| **Q22** | Отсутствие метрик (счётчики parse fail / OCR fail) |
| **Q23** | CLI / `__main__` в OCR рядом с production-кодом |
| **Q24** | Комментарии-дубли кода в `plate_lists` |

---

## 7. Priority Matrix

| ID | Issue | Severity | Effort | Priority |
|----|-------|----------|--------|----------|
| A1 | Два параллельных пайплайна парсинга | Critical | L | P0 |
| A2 | Разная семантика ширин (1.5 split vs fallback 1_2) | Critical | M | P0 |
| A3 | Дублирование математики размеров | High | M | P1 |
| A4 | Оркестрация в commercial handler | High | L | P1 |
| A5 | Два доменных PlateOrder | High | M | P1 |
| A6 | Два механизма номенклатуры | High | M | P1 |
| S1 | Нет rate limit OCR в Telegram | High | S | P1 |
| S2 | Нет лимита размера изображения в боте | High | S | P1 |
| Q1 | Два пайплайна — разное поведение | High | L | P0 |
| Q2 | Недостижимые ветки ширины | High | S | P2 |
| Q3 | Расхождение WxL parse vs wide | High | M | P1 |
| Q4 | Тройное дублирование математики | High | M | P1 |
| Q5 | Молчаливый load=8.0 fallback | High | S | P1 |
| A7 | Третья копия детекции форматов | Medium | M | P2 |
| A8 | OCR вне доменного контракта | Medium | M | P2 |
| A9 | normalize не обязателен | Medium | S | P2 |
| A10 | God-method PlateParserService | Medium | M | P2 |
| A11 | Нет паритета тестов legacy vs service | Medium | M | P2 |
| A12 | Двойной парсинг в боте | Medium | S | P2 |
| S3–S8 | OCR/PII/ошибки/лимиты/prompt injection | Medium | M | P2 |
| Q6–Q16 | Качество кода, логи, handler | Medium | S–M | P3 |
| A13–A16, S9–S11, Q17–Q24 | Low / suggestions | Low | S | P4 |

*Effort: S = small (часы), M = medium (1–3 дня), L = large (спринт+)*

---

## 8. Next Steps

### Immediate (0–3 дня) — P0

1. **Зафиксировать расхождение A2** — добавить тест-кейсы на плиту 1.5 м для обоих путей; временно документировать «источник истины» = legacy `plate_lists` до миграции.
2. **Запретить silent fallback в `PlateParserService`** (строки 111–113) — логировать и возвращать в `unrecognized_lines`.
3. **S2**: лимит байт изображения в боте до вызова OCR.
4. **S4**: заменить `str(e)` на безопасные коды ошибок для пользователя.

### Sprint (1–2 недели) — P1

1. **Консолидация A1/Q1**: единый `PlateOrderParser`, миграция `commercial.py` и `kp_db`.
2. **A3/Q4**: модуль `plate_dimensions`, удаление дублей.
3. **S1**: rate limiting OCR для Telegram = веб-политика.
4. **Golden tests**: 20–30 реальных строк заказов → expected `PlateOrder` (legacy = service).

### Backlog — P2–P4

- Рефакторинг `commercial.py` → use-cases ([A4], [Q12])
- Порт OCR + политики PII ([A8], [S3], [S7])
- Очистка agent-log, print, временных файлов ([S5], [S9], [S10], [Q13], [Q14])
- ADR: контракт `normalize → parse → distribute`
- Метрики и observability ([Q22])

---

## 9. Позитивные находки

| Область | Находка |
|---------|---------|
| **SQL** | Запросы в `kp_db` используют параметризацию (`?`) — снижен риск SQL-инъекций |
| **Секреты** | Ключи OpenAI и прочие секреты загружаются из окружения, не захардкожены в scope-модулях |
| **Веб OCR** | API-путь (`commercial_workflow_service`) имеет rate limiting и валидацию загрузки — образец для унификации бота |
| **Парсер строки** | `parse_line` выделен отдельно и покрыт unit-тестами (`test_plate_normalizer.py`) |
| **Защита от OCR-ошибок** | В legacy `plate_lists` есть проверки адекватности размеров и лимит qty ≤ 500 |
| **Нормализация** | `normalize_order_text` централизует предобработку текста — хорошая база для единого pre-step |

---

## 10. Связанные команды Cursor

| Команда | Когда использовать |
|---------|-------------------|
| [`/refactor`](.cursor/commands/refactor.md) | Консолидация пайплайнов, вынос `add_items`, разбиение `commercial.py` и `PlateParserService` |
| [`/orchestrate`](.cursor/commands/orchestrate.md) | Полный цикл: план → единый parser → тесты паритета → миграция потребителей |
| [`/implement`](.cursor/commands/implement.md) | Точечные фиксы: rate limit, лимит фото, безопасные ошибки, тесты на 1.5 м |

### Связанные аудиты

- [Аудит `core/ocr_gpt.py`](2026-05-19-core-ocr-gpt-audit.md) — углублённый разбор OCR-модуля (Health Score 6.0/10)

---

*Отчёт сформирован documenter-агентом на основе агрегированных результатов senior-reviewer, security-auditor и reviewer.*
