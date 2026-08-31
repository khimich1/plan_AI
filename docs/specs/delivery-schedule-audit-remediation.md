# Spec: Remediation аудита «График поставки» (P0 + «скоро»)

Дата: 2026-08-11. Статус: spec accepted (scope P0+скоро); plan на ревью.
Источник: `ai_docs/develop/audits/2026-08-11-delivery-schedule-audit.md`
Базовая спека модуля: `docs/specs/delivery-schedule.md` (не заменяет её — точечные исправления).

## Objective

Закрыть срочные и близкие по смыслу находки аудита, чтобы модуль можно было
безопасно отдавать нескольким менеджерам и опираться на светофор в переговорах
с клиентом — без ложного green и без скрытых дыр доступа.

**В scope (P0 + «скоро»):**

| ID | Что делаем |
|----|------------|
| A1 / S1 / S2 | AuthZ: `assert_offer_read_access` на GET и POST `/import` (+ тест 403) |
| S7 | После AuthZ: чужой КП → 403, свой без графика → 404 (не раскрывать чужие) |
| A3 | Read-only просмотр графика у КП не «в работе»/«На СГП» (кнопка не disabled) |
| A5 | Produced считать по `plate_id` / без завышения остатка (нет ложного green) |
| Q4 | Импорт: конфликт дат у строк с одним именем партии → unmatched / warning, не тихий merge |
| Q1 / A10 / S9 | Сузить `except` в светофоре; сигнал деградации (`traffic_light_degraded`) |

**Вне scope (явно):**

- God-service refactor (A2/Q6), N+1 / multi-conn (A8/A9)
- Drift констант ёмкости plan_calendar (A4) — отдельная задача после CP2
- Warning `produce_by ≤ deliver_from` (Q5) — soft UX, не блокер
- Upload magic / formula injection / PUT limits (S3–S5)
- Бейдж списка / kp_files (A6/A7)
- AuthZ на пустой `/template` (S10) — опционально заодно с P0, не обязательно

## Tech Stack

Без новых зависимостей. Существующий стек: FastAPI, Pydantic v2, SQLite,
`assert_offer_read_access` / `assert_offer_write_access` из `app.security.offer_access`,
React feature `delivery-schedule`, pytest + vitest.

## Commands

```bash
# Backend
.venv/bin/python -m pytest tests/test_delivery_schedule_*.py tests/test_archive_authorization.py -q
.venv/bin/python -m pytest tests/test_delivery_schedule_endpoints.py tests/test_delivery_schedule_service.py tests/test_delivery_schedule_xlsx.py -q

# Frontend
cd frontend && npm run typecheck && npm run test -- --run
```

## Project Structure (затрагиваемые файлы)

```
app/api/v1/endpoints/delivery_schedule.py     # прокинуть user; маппинг 403
app/services/delivery_schedule_service.py     # get/import AuthZ; produced; except; degraded flag
app/schemas/delivery_schedule.py              # traffic_light_degraded (+ warnings import optional)
core/delivery_schedule_xlsx.py                # конфликт дат при merge партий
frontend/.../OfferDetailsDrawer.tsx           # кнопка: enable + readOnly
frontend/.../types/deliverySchedule.ts        # поле degraded
tests/test_delivery_schedule_endpoints.py     # IDOR 403
tests/test_delivery_schedule_service.py       # produced, degraded
tests/test_delivery_schedule_xlsx.py          # conflicting dates
frontend/.../OfferDetailsDrawer.test.tsx      # кнопка не disabled для архива (если есть)
```

## Code Style

Паттерн AuthZ как в `archive_service` / `generate_document`:

```python
def get(self, kp_id: int, *, user: dict, today: str | None = None) -> DeliveryScheduleView:
    offer = self._fetch_offer(kp_id)
    if not offer:
        raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
    assert_offer_read_access(user, offer)  # → 403 через существующий маппинг endpoint
    # ... load schedule or NotFound "график не найден"
```

В endpoint: `user` не `_user`; ловить отказ доступа так же, как archive
(проверить, какой exception бросает `assert_offer_*` — PermissionError /
кастом — и маппить в 403).

## Testing Strategy

| Область | Уровень | Что покрыть |
|---------|---------|-------------|
| AuthZ GET/import | endpoint + auth fixtures | manager_b не читает график manager_a → **403**; свой без графика → **404**; admin OK |
| Produced | service unit | две строки `kp_plates` с одной identity; `completed` меньше суммы qty → остаток/статус не «полностью покрыто» ложно |
| Import dates | xlsx unit | две строки одно имя партии, разные `deliver_from` → unmatched `conflicting batch dates` (или аналог), партии не молча мержатся |
| Degraded | service unit | monkeypatch calendar/readiness → exception ожидаемого типа → view с `traffic_light_degraded=True`, status=None; неожиданный баг (например TypeError в check) — **не** глотать (пробрасывается или 500) |
| Read-only UI | component/manual | у КП «в архиве»/«выполнено» кнопка активна, диалог `readOnly`, Save/Import скрыты |

Регрессия: существующие `test_delivery_schedule_*.py` остаются зелёными.

## Boundaries

- **Always:** тесты AuthZ зелёные перед merge; минимальный diff; не менять семантику PUT/write; не трогать Not Doing базовой спеки.
- **Ask first:** изменение схемы БД; рефакторинг сервиса сверх точечных правок; смена контракта API кроме добавления `traffic_light_degraded` (и опционально `warnings` на import).
- **Never:** ослаблять `assert_offer_write_access` на PUT; «чинить» IDOR только на фронте; коммитить секреты/живые БД.

## Success Criteria

1. **AuthZ:** manager не может GET/import чужой `kp_id` → HTTP **403**. Свой КП без графика → **404**. Admin и владелец — как сейчас (200 при наличии).
2. **Read-only:** для КП со статусом не из `{в работе, На СГП}` кнопка «График поставки» **не** `disabled`; открывается диалог с `readOnly=true`; PUT с UI недоступен; backend PUT по-прежнему 422/ошибка валидации статуса.
3. **Produced:** при двух `kp_plates` с одинаковой маркой/размерами суммарный вычет produced по партии **не превышает** фактический `on_sgp` / completed для этой identity (нет «каждый plate_id получил полный on_sgp»). Светофор не показывает green из-за двойного вычета.
4. **Import:** конфликт дат у одного имени партии → строка(и) в `unmatched_rows` с явной причиной; черновик не содержит «тихо» даты первой строки + чужие qty.
5. **Degrade:** при недоступности readiness/calendar GET не 500; в ответе `traffic_light_degraded: true`; UI может показать предупреждение (минимум — поле в API; Alert на фронте — желательно в том же PR).
6. Все релевантные pytest/vitest из Commands — зелёные.

## Open Questions

Закрыты выбором scope «P0 + скоро» (2026-08-11):
- God-service / A4 constants / Q5 produce_by warning / upload hardening — **вне** этой спеки.
- `/template` AuthZ (S10): не блокирует; можно добавить в P0 за +15 мин — **по желанию implementer**, не обязательный критерий.

Остаётся на реализацию:
- Точный exception type от `assert_offer_read_access` для маппинга 403 (смотреть `app/security/offer_access.py` и archive endpoints).
- Стратегия A5: пропорциональное распределение `on_sgp` vs прямой SQL по `completed_plates` — выбрать минимальный корректный вариант в Plan.

## Связь с аудитом

После реализации — обновить секцию Remediation в
`ai_docs/develop/audits/2026-08-11-delivery-schedule-audit.md`.
