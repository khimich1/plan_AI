# Runbook: `allow_cross_kp` при списании плит (S4)

**Audit ID:** S4 (controlled)  
**Связано:** `find_kp_plate_row`, `PlateCompletionService.move_plates_to_completed`

## Политика по умолчанию

- `allow_cross_kp=False` — единственный режим в production и в обычных bot/web flow.
- Списание ищет строку только в рамках `prefer_kp_id` (и шагов matching без глобального скана).

## Когда допустимо `allow_cross_kp=True`

Только при **явном** операционном решении:

1. Восстановление после сбоя плана, когда плиты одного `plan_id` числятся на другом `kp_id`.
2. Ручная коррекция данных под контролем администратора (бот/web admin, остановленный бот при массовых правках).

**Кто может включать:** роль `admin` (Telegram) или авторизованный admin API (FastAPI). Не передавать флаг из пользовательского ввода без проверки роли.

## Обязательные меры при включении

1. Записать в audit: `plate_status_log` через `audit_append` / `PlateAuditRepository` (actor = telegram user id или web user id).
2. Зафиксировать в логе приложения: `kp_id`, `plan_ids`, количество списанных плит, причина.
3. Не включать `KP_DB_AGENT_DEBUG` / `APP_DEBUG` в production одновременно с cross-KP (риск PII в `debug_logs/`).

## Запрещено

- Глобальный default `allow_cross_kp=True` в коде или конфиге.
- Включение в commercial/KP save path, offers, archive.
- Автоматическое включение в оптимизаторе или `plan_commit` без review.

## Проверка после операции

```bash
pytest tests/test_kp_db_find_matching_rests.py tests/test_plate_completion_service.py -q
```

## Ссылки в коде

- Matching: `core/domain/plate_completion_matching.py` — docstring `find_kp_plate_row`
- Orchestration: `core/plate_completion_service.py`
