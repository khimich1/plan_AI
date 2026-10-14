# Spec: PDF спецификации отдельной кнопкой

> **Продолжает:** [`specification-print-form.md`](specification-print-form.md). Бланк № 650, абзацы, ворота производства и `GET .../specification/file` без параметра не меняются.  
> **Фаза SDD:** PLAN + IMPLEMENTATION. Пользователь одобрил план и реализацию 2026-10-01.  
> **Дата:** 2026-10-01

Кнопка «Скачать PDF» стоит в панели спецификации архива справа от «Скачать Excel». Файл — PDF той же книги, что уже отдаёт Excel.

---

## Assumptions I'm Making

1. Нужен PDF бланка № 650, а не второй макет. Книга собирается как сейчас (`build_specification_xlsx`), затем LibreOffice (`soffice --convert-to pdf`) переводит её в PDF. Reportlab и отдельная вёрстка PDF не используются.
2. Кнопка только в панели спецификации, в том месте, куда указывает стрелка на скриншоте: после «Скачать Excel». Подпись: «Скачать PDF». Поведение то же, что у Excel: черновик проверяется, сохраняется, затем скачивается.
3. Полка «Документы» карточки КП не меняется. «Скачать спецификацию» по-прежнему отдаёт xlsx. Кнопки «PDF» и «XLSX» на этой полке — коммерческое предложение, не спецификация.
4. `GET /api/v1/commercial/archive/{kp_id}/specification/file` без `format` остаётся xlsx. PDF — тот же путь с `format=pdf`, по образцу скачивания договора.
5. Нет `soffice`, конвертация не удалась или истёк таймаут — ответ 503 с русским текстом. Тело ответа не является xlsx.
6. В образе бэкенда сейчас только `libreoffice-writer-nogui`. Для xlsx нужен Calc: в `docker/backend/Dockerfile` добавляется `libreoffice-calc-nogui`. Образ в этой задаче на машине разработчика не пересобирается, если сборка Docker не нужна для тестов.
7. Имя файла то же, что у Excel, с расширением `.pdf`: `Спецификация КП {id}.pdf` или `Спецификация по счету {номер}.pdf`.
8. Живой `soffice` в pytest не вызывается. Конвертер подменяется.

→ Если что-то из этого неверно — править spec до плана.

---

## Objective

Менеджер в архиве, на статусе «на согласовании», сохраняет спецификацию и скачивает её PDF рядом с Excel. PDF — печатный вид той же книги: рамка, таблица, три абзаца, подписи.

### Пользователь

Менеджер КП. Уже умеет скачивать спецификацию в Excel и сверять её с бланком № 650.

### Сценарии

1. **Обе кнопки.** В панели видны «Сохранить», «Скачать Excel» и «Скачать PDF». PDF не заменяет Excel.
2. **Скачать PDF.** Менеджер нажимает «Скачать PDF». Выбор сохраняется, как при Excel, затем браузер получает PDF. В файле те же позиции, итог и абзацы, что попали бы в xlsx этой сохранённой спецификации.
3. **Скачать Excel.** Кнопка и `GET` без `format` ведут себя как до этой задачи.
4. **Нет LibreOffice.** Запрос `format=pdf` отвечает 503. Клиент показывает текст ошибки. Файл xlsx при этом скачивается.
5. **Спецификация не сохранена.** Как и для Excel: пустое хранилище не отдаёт файл. Кнопка в панели сначала сохраняет черновик.
6. **Полка документов.** «Скачать спецификацию» по-прежнему качает xlsx уже сохранённой спецификации и не предлагает PDF.

---

## Commands

```bash
pytest tests/test_specification_archive.py tests/test_archive_endpoints.py -q
cd frontend && npm run test -- src/features/commercial-archive/components/SpecificationPanel.test.tsx src/features/commercial-archive/components/OfferDetailsDrawer.test.tsx
```

Полный прогон по желанию: `pytest` из корня; `cd frontend && npm run typecheck`.

## Project Structure

```
core/specification_xlsx.py                 книга не меняется
app/services/archive_service.py            download_specification(format)
app/api/v1/endpoints/archive.py            query format=xlsx|pdf
docker/backend/Dockerfile                  libreoffice-calc-nogui
frontend/.../SpecificationPanel.tsx        кнопка «Скачать PDF»
frontend/.../archiveApi.ts                 download с format=pdf
tests/                                     конвертер подменён, 503 без soffice
```

Конвертацию xlsx → pdf держать рядом с уже существующим `convert_docx_to_pdf`, без второго способа вызывать `soffice`. Договорный docx этот путь не открывает.

## Code Style

Эндпоинт по-прежнему отдаёт байты. Формат — явный параметр, по умолчанию книга.

```python
@router.get("/{kp_id}/specification/file")
def download_archive_specification(
    kp_id: int,
    file_format: Literal["xlsx", "pdf"] = Query(default="xlsx", alias="format"),
) -> Response:
    downloaded = service.download_specification(kp_id, user=user, file_format=file_format)
```

Кнопка на панели вызывает тот же `submit`, что и Excel, с другим колбэком скачивания.

## Testing Strategy

| Слой | Где | Что |
|---|---|---|
| Сервис | `tests/test_specification_archive.py` | `format=pdf` вызывает конвертер с байтами xlsx и возвращает PDF; имя с `.pdf`; без сохранённой спецификации файл не создаётся |
| HTTP | `tests/test_archive_endpoints.py` | без `format` — xlsx и прежний Content-Type; `format=pdf` — `application/pdf`; нет `soffice` — 503, тело не xlsx |
| Панель | `SpecificationPanel.test.tsx` | три кнопки; «Скачать PDF» передаёт тот же черновик, что и Excel |
| Полка | `OfferDetailsDrawer.test.tsx` | «Скачать спецификацию» по-прежнему xlsx |

Живой LibreOffice в тестах не запускается. Сверка PDF с бланком № 650 на экране — ручная, после реализации.

## Boundaries

- **Always:** PDF из только что собранной книги; `format` по умолчанию `xlsx`; 503 без рабочего `soffice`; кнопка в панели спецификации; байты не пишутся в папку обмена; тесты не вызывают живой `soffice`.
- **Ask first:** PDF на полке «Документы»; отдельный макет PDF; хранить PDF в базе; менять бланк `templates/specification.xlsx`.
- **Never:** подменять «Скачать Excel» на PDF; отдавать xlsx с заголовком PDF или наоборот; коммитить чужой заполненный № 650; рисовать бланк заново в Reportlab.

## Success Criteria

- [x] В панели спецификации рядом с «Скачать Excel» есть «Скачать PDF»
- [x] Нажатие сохраняет текущий выбор и скачивает PDF
- [x] PDF получен конвертацией той же книги, что отдаёт Excel для этой сохранённой спецификации
- [x] `GET .../specification/file` без `format` по-прежнему отдаёт xlsx
- [x] `GET .../specification/file?format=pdf` отдаёт `application/pdf` и имя с `.pdf`
- [x] Нет `soffice` или сбой конвертации — 503, тело не xlsx
- [x] В `docker/backend/Dockerfile` есть `libreoffice-calc-nogui`
- [x] Полка «Документы» и кнопка «Скачать спецификацию» качают xlsx, как раньше

## Open Questions

Блокирующих нет, если допущения 1–3 верны. Единственное, что стоит подтвердить вслух: PDF нужен только в панели, а полка «Скачать спецификацию» остаётся xlsx.
