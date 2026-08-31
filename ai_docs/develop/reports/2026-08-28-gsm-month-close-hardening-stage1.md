# Report: ГСМ — hardening закрытия месяца, этап 1

**Date:** 2026-08-28  
**Status:** этап 1 сделан (тесты). Коммита нет. Live UI не гоняли.  
**Spec:** [`../../specs/gsm-month-close-hardening.md`](../../specs/gsm-month-close-hardening.md)  
**Plan:** [`../plans/2026-08-28-gsm-month-close-hardening.md`](../plans/2026-08-28-gsm-month-close-hardening.md)

## Summary

Сервер повторяет замки комплекта (хвост чужого месяца, `chain_broken`, красные) на `POST /report/usage` и `POST /waybills/export` **до** soffice и **до** `exported`. Generate августа при открытом июле — отказ; при только разрыве цепи — идёт. Бейдж `has_red_days` важнее «нужна генерация». «Экспорт» в строке прыгает на хвост; текущий период идёт через `planKit`; 4xx — ошибка, не «скачан zip».

## Tasks

| Task | Содержание | Статус |
|------|------------|--------|
| T1 | `gsm_kit_gate` + unit pytest | ✅ |
| T2 | Гейт в usage zip и прямой export zip | ✅ |
| T3 | Generate / bulk: хвост стоп, цепь нет | ✅ |
| T4 | `_status_of` + `handleExportKit` | ✅ |

## Files

**Созданы**

- `app/services/gsm_kit_gate.py`
- `tests/test_gsm_kit_gate.py`

**Изменены**

- `app/services/gsm_report_service.py`
- `app/services/gsm_export_service.py`
- `app/services/gsm_generation_service.py`
- `app/services/gsm_overview_service.py`
- `tests/test_gsm_usage_report.py`
- `tests/test_gsm_export.py`
- `tests/test_gsm_generation_api.py`
- `tests/test_gsm_generate_bulk_api.py` (фикстура старта: предыдущий месяц `exported`, иначе гейт хвоста ломал generate апреля)
- `tests/test_gsm_overview_api.py`
- `frontend/src/features/gsm/components/FleetOverviewView.tsx`
- `frontend/src/features/gsm/components/FleetOverviewView.test.tsx`
- `ai_docs/develop/plans/2026-08-28-gsm-month-close-hardening.md`

## Commands

```
.venv/bin/python -m pytest tests/test_gsm_kit_gate.py tests/test_gsm_usage_report.py tests/test_gsm_export.py tests/test_gsm_generation_api.py tests/test_gsm_overview_api.py tests/test_gsm_generator.py -q
# 137 passed

cd frontend && node node_modules/vitest/vitest.mjs run src/features/gsm/components/FleetOverviewView.test.tsx src/features/gsm/lib/exportGate.test.ts
# 58 passed (2 files)
```

Дополнительно: `tests/test_gsm_generate_bulk_api.py` зелёный после смены фикстуры старта на `exported`.

## Review (этап 1)

Вердикт: **approve** (нет блокеров этапа 1).

- Гейт читает `fleet_overview` + `_chain_broken` (порог 0.01 л), коды `gsm_kit_tail` / `gsm_kit_chain` / `gsm_kit_red`.
- Смесь → 200, zip только чистых; одна плохая / ноль прошедших → 4xx, статусы не меняются.
- Skip-список в HTTP не добавляли.
- **Nit:** `generate()` дергает полный overview на каждую машину (bulk = N одинаковых SQL). Не чинили — минимальный diff.
- **Nit:** прыжок «Экспорт» на чужой месяц по-прежнему без клиентского `planKit` июля (нет данных июля на экране августа); замок — сервер.

## Не делали

- Этапы 2–4 (норма/round, импорт, soffice/job)
- `core/gsm/generator.py`
- Live `plita.db`
- Коммит
- Live UI / browser
- `исключения.txt`
- Полный аудит КП

## Checkpoint

В плане отмечены T1–T4 acceptance и тестовый checkpoint этапа 1. Live UI и коммит — открыты, ждут просьбы.
