# Сопоставление двух аудитов (2026-08-02)

**Цель:** честный apples-to-apples diff между первым и вторым прогоном `/audit`, без двойного счёта и с учётом исправлений в коде между прогонами.

| | Первый прогон | Второй прогон |
|---|---------------|---------------|
| **Строк в отчёте** | 46 находок | 63 находки |
| **Уникальных проблем** | ~28 | ~30 |
| **Scope** | `app/`, `core/`, `frontend/src/` | + `bot/`, `viz_modules/` |
| **Health Score** | 4.0/10 | 0.0/10 (артефакт caps при 4 Critical) |

---

## Сводка (уникальные пункты)

| Статус | Кол-во | Смысл |
|--------|:------:|-------|
| ✅ **Закрыто** | 5 | Было в первом, в коде исправлено |
| 🔄 **Открыто без изменений** | 18 | Та же проблема, тот же или близкий severity |
| ⬆️ **Переклассificировано** | 1 | Та же проблема, severity выше |
| 🆕 **Новое во 2-м** | 8 | Первый прогон не нашёл |
| 🔁 **Дубль в отчёте №2** | 1 | A1 + S1 — одна проблема, два ID |
| 📉 **Пропущено во 2-м** | 6 | Было в первом, во втором не упомянуто явно |
| 🔀 **Разбито / объединено** | ~5 | Одна тема → несколько ID или наоборот |

**Реальных P0-блокеров сейчас: 3** (финансы для logistics, datetime import, gantt import).

---

## 1. ✅ Закрыто с первого аудита

| Было (прогон 1) | Severity | Что изменилось |
|-----------------|----------|----------------|
| **A1** `core` импортирует `app.domain.enums` | Critical | Enum-ы в `core/domain/enums.py`; grep `from app.` в `core/` — пусто |
| **A6** Дублирование enum core ↔ app | High | Следствие A1; SSOT в `core/domain/enums.py` |
| **Q1** (v1) `except Exception: pass` после freeze при move_to_production | High | Оба сервиса вызывают `commit_move_to_production`; ошибка пробрасывается (`ArchiveError` / `ValueError`) |
| **Q2** (v1) Нет атомарности «перевести в производство» | High | `core/kp/offers_write.commit_move_to_production` — одна транзакция; отмечено как сильная сторона во 2-м аудите |
| **Q3** (v1) Неактивные перевозчики исчезают из UI | Medium | `useCarriersQuery` default `activeOnly: false`; `CarriersView` не передаёт `activeOnly: true` |

> **Важно:** во 2-м аудите ID **Q1/Q2 переиспользованы** под другие баги (missing imports). Ниже v1 = прогон 1, v2 = прогон 2.

---

## 2. ⬆️ Переклассificировано

| Проблема | Прогон 1 | Прогон 2 | Комментарий |
|----------|----------|----------|-------------|
| Утечка финансовых полей КП для роли logistics (`/archive/search`) | **S1 High** | **A1 + S1 Critical** | Та же уязвимость; выше severity + дубль ID |

---

## 3. 🔁 Дубль во втором отчёте

| ID в прогоне 2 | Суть | Действие при подсчёте |
|----------------|------|------------------------|
| A1 + S1 | Logistics видит `subtotal`, `total_amount`, `discount_percent` | Считать как **1 уникальную** проблему |

Без дубля: Critical **3**, не 4. Health Score при той же формуле: `10 − 6 − 3 − 1 = 0.0` (caps не меняются).

---

## 4. 🆕 Новое во втором прогоне (реально новые находки)

| ID (v2) | Severity | Суть | Проверка в коде |
|---------|----------|------|-----------------|
| **Q1** (v2) | Critical | `datetime` без import в `offers_service.py:145,169` | ✅ подтверждено |
| **Q2** (v2) | Critical | `create_gantt_excel` без import в `archive_service.py:352` | ✅ подтверждено |
| **S4** | High | OCR шлёт документы во внешний LLM | Новый security-угол |
| **Q3** (v2) | High | Дублирование `order_data` vs `core/kp_order_data.py` | Новый code-quality |
| **Q5** (v2) | High | Fallback `m=x` маскирует битые данные | Новый |
| **Q9** (v2) | High | `complete()` игнорирует free-only рейсы | Новый |
| **A7** (v2) | High | Partial DI (только AuthService через Depends) | Было A15 Low → поднято |
| **A20, S14** | Low | Bot decommissioned / `bot_archived` DB access | Расширен scope |

**Итого genuinely new blockers: +2 Critical** (imports), **+0 закрытых Critical**.

---

## 5. 🔄 Открыто в обоих прогонах (маппинг ID)

Единая таблица **~18 устойчивых** проблем:

| Уникальная тема | Прогон 1 | Прогон 2 | Severity |
|-----------------|----------|----------|----------|
| God ShipmentService | A2 | A2 | High |
| God CommercialWorkflowService + Service Locator | A3 | A3 | High |
| Сервисы возвращают Pydantic-схемы API | A4 | A4 | High |
| SQL в сервисах без repository | A5 | A5 | High |
| Lazy-import / DIP violation → SgpService | A8 | A6 | High |
| Rate limiting in-process | S2 | S2 | High |
| SQLite-ошибки клиенту | S4 | S3 (+ A22 Low) | High / Low |
| SGP API под `/production` | A7 | A8 | Medium |
| freeze в read-path | Q5 | Q4, A9 | High / Medium |
| Расхождение N/M archive vs SGP | Q4 | A10 | Medium |
| Cross-feature React Query invalidation | A11 | A11 | Medium |
| Ручное дублирование API types | A12 | A12 | Medium |
| God ShipmentItemsSection | A10 | A13 | Medium |
| DraftStore file-based | A13 | A14 | Medium |
| Mutable global plate runtime | A9 | A21 | Medium / Low |
| CSP Report-Only | S3 | S5 | Medium |
| No rate limit mutating API | S6 | S6 | Medium |
| SQLite unencrypted at rest | S7 | S7 | Medium |
| Legacy `/web/login` без CSRF | S5 | S8 | Medium |
| Дублирование move_to_production archive/offers | Q6 | Q6 | High |
| Нет тестов kp_db_shipments | Q8 | Q7 | Medium → High |
| OffersService мало тестов | Q9 | Q8 | Medium → High |
| KpRepository() на каждый вызов | A15 | A17 | Low |
| Legacy routes bypass DI | A16 | A18 | Low |
| Dependency scanning в CI | S10 | S15 | Low |

---

## 6. 📉 Было в прогоне 1, явно не попало в прогон 2

| ID (v1) | Тема | Вероятная причина |
|---------|------|-------------------|
| Q7 (v1) | Magic string `ValueError` в offers | Переименовано → Q19 (v2) Low |
| Q10 (v1) | Нет тестов CommercialExportService | Поглощено общим «test gaps» |
| Q11 (v1) | Нет тестов CreateShipmentDialog | Частично устарело — `CreateShipmentDialog.test.tsx` появился |
| Q12 (v1) | Тихий сбой даты в KpReadinessService | Не повторено (возможно oversight) |
| Q13–Q19 (v1) | DRY freeze, draftWeight, plan_aggregation… | Часть → Q13–Q25 (v2), часть потеряна |
| S8–S11 (v1) | sessionStorage, debug logs, password policy в ответах | S9–S13 (v2) частично; S11 (v1) → не явно |

Не значит «исправлено» — скорее **нepолнота второго прогона** по low-priority.

---

## 7. Матрица «что реально изменилось в проекте»

```
                    Прогон 1          Между прогонами        Прогон 2
                    ────────          ─────────────────        ────────
Critical            1 (core→app)    FIXED enums            3 unique (finance×1, imports×2)
High                9               −3 fixed, +1 reclass,   16 unique (+7 new/split)
                                      +5 new
Medium              22              ~stable, re-ID          24
Low                 14              +scope bot/viz          18

Сырой total         46              —                       63
Уникальный total    ~28             net −5 closed,          ~30
                                      +7 new, +1 dup
```

---

## 8. Рекомендуемый «честный» Health Score

| Метод | Score | Комментарий |
|-------|:-----:|-------------|
| Прогон 1 (46 items) | 4.0 | 1 Crit, 9 High |
| Прогон 2 raw (63 items) | 0.0 | Caps при 4 Crit |
| **Скорректированный** | **~3.5–4.5** | 3 Crit (−6 cap), ~14 unique High (−3 cap), ~22 Med (−1 cap) → **~0–1** по формуле ИЛИ пересчёт без dup: 3 Crit → 4.0 после −2−3−1 |

Практичнее смотреть на **P0/P1**, а не на абсолютный score:

| Приоритет | Уникальные пункты |
|-----------|-------------------|
| **P0** | Financial leak (1), datetime import (1), gantt import (1) |
| **P1** | God-services, repository layer, rate limit, SQLite errors, move_to_production DRY, test gaps (~10) |
| **P2+** | Medium/Low из §5–6 |

---

## 9. Вывод для команды

1. **46 → 63 — не деградация качества кода**, а сумма: переиспользование ID, дубль A1/S1, расширенный scope, более мелкая нарезка, 2 новых runtime-бага.
2. **Между прогонами закрыто ~5 пунктов**, включая бывший единственный Critical (core→app).
3. **Net-new critical risk:** 2 missing imports + escalation financial leak High→Critical.
4. Для tracking использовать **уникальные темы** (столбец в §5), а не сырые ID A*/S*/Q*.

---

*Сопоставление составлено 2026-08-02 на основе первого отчёта (46 находок) и второго (`2026-08-02-full-project-audit.md`, 63 находки) с верификацией по текущему коду.*
