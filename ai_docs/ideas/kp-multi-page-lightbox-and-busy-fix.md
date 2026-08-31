# КП multi-page: lightbox до OCR + busy false positive

> Ideation: 2026-08-31  
> Parent: [`kp-multi-page-screenshots.md`](./kp-multi-page-screenshots.md)  
> Спека: [`../specs/kp-multi-page-screenshots.md`](../specs/kp-multi-page-screenshots.md) (R8/R9)

## Problem Statement

**How might we** убрать ложный «busy» после добавления фото (до «Распознать») и дать менеджеру крупный просмотр скринов **до** OCR — не ломая прогрессивную сверку после старта?

## Recommended Direction

1. **Busy только после старта OCR:** `isRecognizing` = очередь реально бежит (`isRunning` или `hasStarted && hasPendingOrRunning`). До `start` кнопка «Распознать фото», enabled при наличии страниц.
2. **Lightbox до OCR:** клик по thumbnail до `hasStarted` открывает оверлей с крупным превью; Esc / backdrop / ✕ закрывают; ←/→ листают страницы. ✕ на thumbnail по-прежнему удаляет (stopPropagation).
3. **После `hasStarted`:** клик по thumbnail остаётся select для review (PageReviewNav без изменений). Progressive review as designed.

## Key Assumptions

- [x] Пользователь подтвердил: busy-fix + lightbox до OCR; progressive review не трогаем
- [x] Placeholder text рядом с фото — ок
- [ ] Lightbox достаточно лёгкий (маленький компонент), не постоянный split-pane

## MVP Scope

**In:** R8 busy fix; R9 lightbox (open/close/nav) до `hasStarted`; тесты unit/RTL; дописки в спеку/план.

**Out:** redesign progressive review; server job Phase B; постоянный split-preview pane до OCR; commit.

## Not Doing (and Why)

- **Переделка progressive review** — уже работает; delta только pre-OCR UX
- **Server job** — отдельная фаза B
- **Постоянная панель превью рядом с текстом до OCR** — overkill; lightbox по клику достаточен
