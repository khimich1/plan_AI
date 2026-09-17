# Spec: Свободное имя клиента, контрагент 1С — на производстве

> **Источник идеи:** [`ai_docs/ideas/kp-counterparty-bind-on-production.md`](../ideas/kp-counterparty-bind-on-production.md)  
> **План:** [`ai_docs/develop/plans/2026-09-17-kp-counterparty-bind-on-production.md`](../develop/plans/2026-09-17-kp-counterparty-bind-on-production.md)  
> **Фаза SDD:** SPECIFY ✅ → PLAN ✅ → TASKS ✅ → IMPLEMENT ✅  
> **Статус:** реализовано (2026-09-17)  
> **Дата:** 2026-09-17  
> **Связанные модули:** шаг «Клиент» мастера КП, `save_offer`, архив, `move_to_production`, справочник `counterparties`.  
> **Рядом, не это:** импорт контрагентов из 1С, `POST /counterparties`, живой 1С, два имени в КП.

---

## Assumptions I'm Making

1. **«КП не уходит» = перевод в производство**, не запрет PDF/XLSX из архива. Файлы с рабочим именем можно отдавать клиенту.
2. **После привязки** `customer_name` / ИНН / КПП **перезаписываются** из карточки 1С. Уже скачанные файлы в MVP не пересобираем.
3. Клиента нет в списке, потому что **карточки ещё нет в 1С**, а не потому что поиск сломан.
4. Привязывать может любой, у кого уже есть **запись КП** (`assert_offer_write_access`).
5. **Пустое имя нельзя.** Нужна непустая строка «клиент», даже без `counterparty_id`.
6. **Append / resume** не меняем: шапка липкая. Привязка 1С — в архиве, не повторный шаг «Клиент».
7. Параллельный `POST /api/v1/offers` (`CreateOfferRequest`) **выравниваем**: без карточки можно только в архив; `save_mode=work` по-прежнему требует клиента 1С.
8. Схему SQLite **не трогаем**: `KP_offers.counterparty_id` уже nullable.

→ Если что-то из этого неверно — править spec до кода.

---

## Decisions locked (ideation 2026-09-17)

| # | Тема | Решение |
|---|------|---------|
| D1 | Поле | Одно: имя + опциональный `counterparty_id` |
| D2 | Шаг «Клиент» | Автокомплит как сейчас. Не нашёл — любое непустое имя, расчёт разрешён |
| D3 | Создание в справочнике | Из мастера убрать. `POST /counterparties` не удаляем, UI не зовёт |
| D4 | Архив | Сохранение без карточки: `customer_name` = ввод, `counterparty_id` / ИНН / КПП = null |
| D5 | Снапшот при выборе | Если id есть — имя и реквизиты **только** из справочника (как сейчас) |
| D6 | Производство | `move_to_production` → `require_active_client`. Кнопка неактивна без id |
| D7 | Привязка | `PATCH …/archive/{kp_id}/counterparty` в статусе «в архиве». Перезапись снапшота |
| D8 | Перепривязка | Пока «в архиве» — можно сменить карточку. Отвязать нельзя |
| D9 | Файлы | На bind не регенерируем. Имя в PDF обновится при следующем сохранении из конструктора |
| D10 | Список архива | Бейдж «нет 1С». Фильтра очереди в MVP нет |
| D11 | Skip шага | Только `counterparty_id` / append / resume. Голое имя шаг не пропускает |

---

## Objective

Менеджер считает КП и кладёт в архив, не дожидаясь карточки 1С. В производство КП не выходит, пока не привязан активный клиент из справочника.

### User stories

| # | Как… | Я хочу… | Чтобы… |
|---|------|---------|--------|
| US-1 | менеджер | ввести любое имя, если клиента нет в списке | посчитать КП сегодня |
| US-2 | менеджер | выбрать клиента из справочника, если он есть | не набивать имя руками и сразу иметь ИНН/КПП |
| US-3 | менеджер | сохранить в архив без кода 1С | отдать черновик файлов заказчику |
| US-4 | офис | видеть «нет 1С» в архиве | не отправить в завод безымянную карточку |
| US-5 | офис | занести контрагента в карточке архива | открыть «В производство» |
| US-6 | система | не принимать `move_to_production` без клиента | нельзя обойти UI |

### Reframed success criteria

| Требование | Критерий |
|------------|----------|
| Расчёт без 1С | Шаг «Клиент»: непустой `clientName`, `counterpartyId` может быть null → «Рассчитать КП» |
| Архив без 1С | `save_offer` с именем и без id → 200, строка КП с `counterparty_id IS NULL` |
| Архив с 1С | С id → имя/ИНН/КПП из справочника, не из формы |
| Гейт | `POST …/move-to-production` без активного клиента → 400, КП остаётся «в архиве» |
| Привязка | PATCH валидного клиента → снапшот обновлён, кнопка производства активна |
| Мастер | Нет «Добавить контрагента» / вызова `createCounterparty` |
| Skip | Новое КП с одним именем всё ещё останавливается на шаге «Клиент» |

---

## Tech Stack

Тот же: FastAPI + Pydantic v2, SQLite `plita.db`, React 19 / TypeScript, vitest + pytest. Новых зависимостей нет.

## Commands

```bash
# backend
pytest tests/test_offer_counterparty_validation.py tests/test_kp_counterparty_snapshot.py \
  tests/test_archive_service.py tests/test_archive_endpoints.py \
  tests/test_commercial_draft_append.py -q

# frontend
cd frontend && npx vitest run \
  src/features/commercial-offer/components/steps/ClientConditionsStep.test.tsx \
  src/features/commercial-offer/components/CounterpartyAutocomplete.test.tsx \
  src/features/commercial-offer/lib/wizardStepOrder.test.ts \
  src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx \
  src/features/commercial-archive/components/ArchiveOfferList.tsx
```

Точные имена новых тестов — в плане (CBP-*).

## Project Structure

```
app/services/commercial_draft_lifecycle.py   # save_offer: имя без id
app/services/archive_service.py              # move_to_production + bind
app/services/offers_service.py               # параллельный create_offer
app/api/v1/endpoints/archive.py              # PATCH counterparty
app/schemas/archive.py / offers.py
frontend/.../ClientConditionsStep.tsx
frontend/.../CounterpartyAutocomplete.tsx    # убрать NewCounterpartyDialog
frontend/.../OfferDetailsDrawer.tsx          # бейдж, bind, гейт кнопки
frontend/.../ArchiveOfferList.tsx            # бейдж в списке
```

## Code Style

Гейт производства — тот же сервис, что карточка:

```python
def move_to_production(self, kp_id: int, terms_input: str, *, user: dict) -> ArchiveOfferDetails:
    raw = self.repository.get_by_id(kp_id)
    # ... status «в архиве», write access ...
    CounterpartiesService(db_path=self.repository.db_path).require_active_client(
        raw.get("counterparty_id")
    )
```

Bind не пишет свободное имя в `counterparties`. Только FK + снапшот с карточки.

Сообщения пользователю — по-русски, в том же тоне, что «Контрагент не найден в справочнике».

## Testing Strategy

| Слой | Где | Что |
|------|-----|-----|
| Backend unit | `tests/test_offer_counterparty_validation.py`, новый `test_archive_counterparty_bind.py` | save без id; save с id; move без id; PATCH bind / supplier / inactive / не архив |
| Wizard skip | уже есть `should_skip_client_step` / `wizardStepOrder.test.ts` | голое имя ≠ skip |
| FE wizard | `ClientConditionsStep.test.tsx`, autocomplete | submit с текстом без id; нет кнопки «Добавить контрагента» |
| FE archive | `OfferDetailsDrawer.test.tsx`, list | бейдж; disabled «В производство»; bind вызывает PATCH |

Не тащить живую 1С и не расширять импорт справочника.

## Boundaries

- **Always:** тесты на оба пути (с карточкой / без); `require_active_client` на любом входе в «в работе»; не создавать строки в `counterparties` из мастера.
- **Ask first:** фильтр архива «без 1С»; авто-перегенерация PDF на bind; закрыть `POST /counterparties` для manager; второе поле «имя в КП».
- **Never:** живой API 1С; фейковый `code_1c`; снимать гейт производства «на потом»; менять skip append/resume.

## Success Criteria

- [x] Новое КП: нет в справочнике → имя → расчёт → «В архив» → статус «в архиве», `counterparty_id` null
- [x] То же КП: «В производство» недоступно (UI + 400 API)
- [x] После PATCH карточки 1С: имя/ИНН/КПП из справочника, производство открыто
- [x] Клиент уже в справочнике: путь как до изменения (снапшот, skip шага при id)
- [x] В мастере нельзя завести контрагента
- [x] Регрессии: скидка/логистика архива, resume, append партий, простые типы с «В производство / скоро»

## Out of scope

- Два отображаемых имени (рабочее vs юрлицо)
- Фильтр/очередь «без 1С»
- Реген файлов на bind
- Импорт/синхронизация справочника 1С
- Смена контрагента у КП уже «в работе» / «выполнено»

## Open Questions

Нет блокирующих. Открытые из идеи закрыты D9–D11. Если после приёмки файлы с новым юрлицом нужны сразу — отдельная задача на реген, не этот MVP.
