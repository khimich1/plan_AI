# GigaChat Vision OCR для списков плит

> **Спека:** [`ai_docs/specs/gigachat-vision-ocr.md`](../specs/gigachat-vision-ocr.md)  
> **План:** [`ai_docs/develop/plans/2026-07-08-gigachat-vision-ocr.md`](../develop/plans/2026-07-08-gigachat-vision-ocr.md)

## Problem Statement

**Как нам распознавать списки ЖБ-плит с фото максимально точно (разные форматы записи, запятые, нагрузки), оплачивая API как физлицо в рублях и не тратя лишние вызовы Vision на простые чёткие снимки?**

## Recommended Direction

Заменить GPT-4o Vision на **GigaChat 2 Max** с **адаптивным пайплайном**: число API-вызовов регулируется правилами и настройками, а не фиксированным «всегда два раза».

### Почему GigaChat 2 Max

- Лучше следует инструкциям (IFEval-ru 0,83) — критично для JSON-контракта и правил «66,2 ≠ 66».
- Vision из коробки, русский контекст, оплата в рублях, данные в РФ.
- Для физлица: freemium + пакеты без договора с юрлицом.

### Почему адаптивное число вызовов

При 50–200 фото/мес лишний verify на каждом снимке удваивает расход токенов (~2–4 ₽ → ~4–8 ₽ за фото). Часть фото — короткая печатная таблица с телефона: одного Extract часто достаточно. Verify нужен на длинных, размытых или «сомнительных» результатах.

### Архитектура пайплайна

```
Фото → [опционально: препроцессинг] → Extract (Max)
     → Parser Gate (plate_line_parser, бесплатно)
     → Решение: нужен ли Verify?
         ├─ нет  → результат (1 вызов API)
         └─ да   → Verify (Max) → Parser Gate → результат (2 вызова API)
```

**Parser Gate** — локальная проверка каждой строки через `core.plate_line_parser.parse_line()`. Не тратит токены. Понижает `confidence` и добавляет `issues: ["parser_rejected"]` при несовпадении с форматами парсера.

### Режимы регулировки вызовов

| Режим (`OCR_VERIFY_MODE`) | Поведение |
|---------------------------|-----------|
| `never` | Всегда 1 вызов (только Extract) |
| `always` | Всегда 2 вызова (Extract + Verify) |
| `auto` *(рекомендуется)* | Verify только при срабатывании эвристик |

Жёсткий потолок: `OCR_MAX_API_CALLS` (1 или 2) — даже в `auto` не больше заданного числа.

### Эвристики для `auto` (когда пропустить Verify)

Verify **не вызывается**, если **все** условия выполнены:

1. **Размер файла** ≤ `OCR_VERIFY_AUTO_MAX_BYTES` (по умолчанию 800 КБ) — типичное чёткое фото с телефона.
2. **Число строк** ≤ `OCR_VERIFY_AUTO_MAX_ROWS` (по умолчанию 15).
3. **Уверенность**: у каждой строки `confidence` ≥ `OCR_VERIFY_AUTO_MIN_CONFIDENCE` (по умолчанию 0,92).
4. **Parser Gate**: все строки успешно парсятся (`parsed == true`).
5. **Issues**: пустой `issues` у всех строк.

Verify **всегда вызывается**, если хотя бы одно:

- строк больше порога;
- есть `parser_rejected` или `confidence` ниже порога;
- файл больше порога (много деталей / высокое разрешение);
- Extract вернул пустой или обрезанный список;
- принудительно: `OCR_VERIFY_MODE=always`.

Опционально (фаза 2): метрика резкости изображения (Laplacian variance) — если ниже порога, всегда Verify.

### Провайдер и откат

- `OCR_PROVIDER=gigachat` — основной путь.
- `OCR_PROVIDER=openai` — fallback на текущий GPT-4o без переписывания веб-флоу.
- Публичный API для сервисов: `recognize_text_smart()` в `core/ocr/` (обратная совместимость с `commercial_draft_service`).

### Стоимость для физлица

Тарифы GigaChat API ([individual-tariffs](https://developers.sber.ru/docs/ru/gigachat/tariffs/individual-tariffs)):

| | За 1 000 токенов | Пакет |
|--|------------------|-------|
| GigaChat 2 Max | 0,65 ₽ | 3 млн токенов = **1 950 ₽** (12 мес.) |
| Freemium | — | **50 000** токенов Max бесплатно на 12 мес. |

Оценка на одно фото (~4 000–6 000 токенов на Extract):

| Сценарий | Токенов | Стоимость (Max) |
|----------|---------|-----------------|
| 1 вызов (`never` / `auto` пропустил verify) | ~4 000–6 000 | **2,6–4 ₽** |
| 2 вызова (`always` / `auto` с verify) | ~8 000–12 000 | **5–8 ₽** |

При 50–200 фото/мес и доле verify ~30–40% (типично для `auto`):

| Объём | Только 1 вызов | Адаптивно (~35% verify) | Всегда 2 вызова |
|-------|----------------|---------------------------|-----------------|
| 50 фото/мес | ~130–200 ₽ | ~200–350 ₽ | ~250–400 ₽ |
| 200 фото/мес | ~520–800 ₽ | ~800–1 400 ₽ | ~1 000–1 600 ₽ |

Пакет **1 950 ₽ / 3 млн токенов** хватит примерно на **250–750 фото** в зависимости от доли двухэтапных распознаваний.

## Key Assumptions to Validate

- [ ] **GigaChat Max на реальных фото КП не хуже GPT-4o** — пилот на 30–50 снимках из архива менеджеров; метрика: число corrections и parser_rejected.
- [ ] **Эвристики `auto` отделяют «простые» фото** — на пилоте замерить: при каких порогах verify ловит ошибки, которые пропустил один Extract.
- [ ] **Физлицо достаточно для продакшена** — лимиты freemium/пакетов и ToS Сбера покрывают коммерческое использование веб-приложения завода (уточнить в ЛК Studio).
- [ ] **Синхронный SDK GigaChat приемлем по latency** — обёртка `asyncio.to_thread`; целевое время ответа < 15 с на фото.
- [ ] **Менеджеру достаточно UI corrections** — при редких ошибках не нужен отдельный экран «сравнение с фото».

## MVP Scope

### В scope

1. **`core/ocr/`** — провайдер GigaChat Max, OpenAI fallback, общий пайплайн.
2. **Адаптивный verify** — `OCR_VERIFY_MODE=auto|always|never`, пороги в env/settings.
3. **Parser Gate** после Extract и после Verify.
4. **Настройки** в `core/config/settings.py` + `.env.example`.
5. **Метаданные в ответе**: `ocr_api_calls`, `ocr_verify_skipped_reason`, `ocr_cost_rub`, `ocr_method`.
6. **Тесты** с моками SDK (без реальных API-вызовов в CI).
7. **Пилот** на freemium 50k токенов, затем пакет Max при необходимости.

### Переменные окружения (черновик)

```env
OCR_PROVIDER=gigachat
OCR_EXTERNAL_ENABLED=true
GIGACHAT_CREDENTIALS=...
GIGACHAT_MODEL=GigaChat-2-Max
GIGACHAT_SCOPE=GIGACHAT_API_PERS

OCR_VERIFY_MODE=auto          # auto | always | never
OCR_MAX_API_CALLS=2           # жёсткий потолок: 1 или 2
OCR_VERIFY_AUTO_MAX_ROWS=15
OCR_VERIFY_AUTO_MIN_CONFIDENCE=0.92
OCR_VERIFY_AUTO_MAX_BYTES=819200   # 800 КБ
```

### Вне scope MVP

- Препроцессинг изображений (обрезка, deskew, резкость).
- Pro + Max hybrid (разные модели на Extract/Verify).
- Асинхронный режим GigaChat (дешевле, но медленнее — не для интерактивного КП).
- Telegram-бот (архив `bot_archived` — отдельная задача).
- Автоматический A/B GPT vs GigaChat в проде.

## Not Doing (and Why)

- **GigaChat Lite для OCR** — слишком слаб для посимвольной точности на марках плит.
- **Три и более LLM-вызова на фото** — дорого и медленно; parser gate закрывает часть кейсов бесплатно.
- **Юрлицо / договор в MVP** — выбрано физлицо; миграция на B2B — отдельное решение при росте объёма.
- **Полный отказ от ручной проверки менеджером** — цель «почти 0 ошибок», не полная автономность.
- **Собственный OCR (Tesseract/EasyOCR) как основной путь** — уже есть `core/ocr_recognition.py`, но не покрывает структуру таблицы и 4 формата; остаётся legacy.

## Open Questions

- Достаточно ли freemium/пакета физлица для **коммерческого** использования внутри завода (не личный эксперимент)?
- Нужны ли **разные пороги `auto` по окружению** (staging: `always`, prod: `auto`)?
- Сохранять ли **статистику** (доля 1 vs 2 вызовов, средний cost_rub) в логах для калибровки порогов?
- При `OCR_MAX_API_CALLS=1` и сомнительном Extract — показывать предупреждение менеджеру или молча отдавать черновик?

## Implementation Sketch

```
core/ocr/
  __init__.py              # recognize_text_smart, apply_plates_with_ai
  pipeline.py              # extract → parser_gate → decide_verify → verify?
  verify_policy.py         # should_run_verify(mode, metrics, settings)
  parser_gate.py           # apply_parser_gate(plates)
  prompts.py               # system + verify prompts (из plate_format_prompt)
  parsing.py               # parse_gpt_response, parse_verify_response
  providers/
    base.py                # Protocol
    gigachat.py            # GigaChat Vision (asyncio.to_thread)
    openai.py              # текущий GPT-4o
```

**`verify_policy.should_run_verify()`** — единая точка для регулировки числа вызовов; пороги читаются из settings; возвращает `(run: bool, reason: str)` для метаданных и логов.

**Порядок внедрения:** провайдер GigaChat → parser gate → verify policy → восстановить verify в пайплайне → пилот → калибровка порогов `auto`.

## Success Criteria

- На пилоте 30–50 фото: **≥ 90%** строк без ручной правки марки/qty (менеджер подтверждает).
- В режиме `auto`: **≥ 50%** фото обрабатываются **1 вызовом** без роста ошибок vs `always`.
- Средняя стоимость фото в `auto` **ниже на 30%+**, чем при `always`, при сопоставимой точности.
- Все существующие тесты `test_recognition_pipeline.py` проходят (с моками провайдера).
