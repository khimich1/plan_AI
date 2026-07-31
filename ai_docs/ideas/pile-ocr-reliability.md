# Надёжное OCR свай (фото → состав КП)

## Problem Statement

**How might we** сделать так, чтобы типовое фото списка свай (3–20 строк) попадало в состав КП **без ручной правки текста в ≥95% случаев**, не ухудшив UX по сравнению с плитами (сейчас 30 с – 1 мин на фото)?

## Recommended Direction

**Direction A+C:** отдельный pile OCR pipeline (как у плит, но с `pile_format_prompt` + `pile_line_parser`) **плюс** детерминированный post-OCR normalizer для типичных ошибок GPT («Сваи 90.30-11 189» → «С90.30-11 189»).

Verify — режим **`auto`**: второй GPT-вызов только если parser gate отклонил строки, низкий confidence или слишком много строк. На чистых таблицах (как на пилотном скрине) — один вызов (~10–15 с).

**Scope дополнительно:** ускорение OCR плит — добиться, чтобы `auto` чаще пропускал verify на типовых коротких таблицах (метрика: `ocr_api_calls=1`).

## Key Assumptions to Validate

- [ ] Марки с фото (`С90.30-11` и т.п.) есть в `pile_prices` — проверить `pb.db` до IMPLEMENT
- [ ] GPT с pile prompt стабильно отделяет слово «Сваи» от марки `С90.30-11`
- [ ] Auto-skip verify не пропускает критичные ошибки qty на «грязных» WhatsApp-фото
- [ ] Текущий провайдер prod — OpenAI или GigaChat; оба должны поддерживать pile extract

## MVP Scope

**In:**
- `product_type=piles` → pile OCR pipeline на create/update draft с фото
- `apply_pile_parser_gate`, `should_run_pile_verify` (auto)
- Normalizer: strip «Сваи»/«Свай», prepend `С` к `\d+\.\d+-\d+`, strip «шт»
- Prompt: few-shot «Сваи С90.30-11», «189 шт»
- Тесты: регрессия по строкам со скрина; mock OCR integration test
- Plate: тест/док что auto-skip срабатывает на 3-row fixture; при необходимости tune verify thresholds

**Out:**
- Tesseract / local OCR
- Fuzzy match марок
- Отдельный verify prompt только для свай (reuse структуры plate verify с pile JSON)
- UI-кнопки auto-fix (batch-review достаточно)
- Сжатие изображений перед OCR

## Not Doing (and Why)

- **Handwriting models** — mixed sources, но без отдельного ML; batch-review как fallback
- **Изменение strict match цен** — по спеке КП на сваи
- **Telegram bot OCR** — web-only
- **Schema changes** — не нужны для OCR

## Open Questions

- Есть ли `С90.30-11` / `С110.30-13` / `С120.30-12` в production `pb.db`?
- Класть ли PNG скрина в `tests/fixtures/` как регрессионный кейс?
