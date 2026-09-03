# КП: очередь исходных фото на шаге 1 (Drawer)

**Статус**: IDEATE ✅ · SPECIFY ✅ · PLAN ✅ · IMPLEMENT ✅  
**Дата**: 2026-09-02  
**Spec**: [../specs/kp-source-image-queue-drawer.md](../specs/kp-source-image-queue-drawer.md)  
**Plan**: [../develop/plans/2026-09-02-kp-source-image-queue-drawer.md](../develop/plans/2026-09-02-kp-source-image-queue-drawer.md)

## Problem Statement

How might we дать менеджеру на **шаге 1** после сверки снова открыть **очередь исходных фото** текущего распознавания (1 или N страниц) в боковом Drawer — пока не начат новый круг набора и не сохранено в архив — без постоянного блока в UI и без серверного хранения?

## Recommended Direction

**Sticky session gallery + side Drawer (FE-only).**

После OCR страницы (одна или несколько) остаются в очереди с blob URL; после «Список верен» URL **не revoke** сразу. На шаге 1 кнопка «Исходные фото (N)» открывает Drawer с просмотром и листанием. Очередь живёт до `start-append-cycle` / нового круга, успешного «В архив» / «Сохранить изменения», или «Создать новое КП». Без upload на сервер, без фото на клиенте/Result, без resume из архива.

## Key Assumptions to Validate

- [x] Кнопка только на шаге 1 (ввод/предпросмотр после сверки)
- [x] Очередь = 1..N страниц текущего захода
- [x] UX = выезжающий Drawer, не split-view
- [x] Clear на новый круг и archive save
- [x] Браузер-only (blob), без сервера

## MVP Scope

**In:** очередь blob URL; кнопка на шаге 1; Drawer + листание; clear на новый круг / archive save / create new  
**Out:** IndexedDB/F5; server storage; Result/client; archive resume с фото; вечный split-view

## Not Doing (and Why)

- Серверное хранение OCR-файлов — политика/объём; «на всякий случай» в сессии хватает
- Фото на Result / клиенте — не просили
- Persist после F5 — later; MVP = вкладка
- Постоянный блок «Исходное фото» на предпросмотре — шум; Drawer по запросу

## Open Questions

_Нет — locked 2026-09-02._
