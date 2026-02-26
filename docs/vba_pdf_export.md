## Export procedures found

Поиск по `extracted_vba/modules/*.bas` и `extracted_vba/classes/*.cls` не нашёл вызовов:

- `ExportAsFixedFormat`
- `PrintOut`, `PrintPreview`
- `PublishObjects.Add`

Итого: **в текущем извлечённом VBA нет процедур экспорта PDF/печати**.

## Export scope (sheet vs workbook)

Не применимо — экспорт не реализован в VBA.

## Output naming/path logic

Не применимо — логики формирования пути/имени PDF в VBA нет.

## Risks for automation

- **Нельзя “просто запустить макрос и получить PDF”**: runner должен либо
  - использовать внешнюю логику экспорта (через Excel COM: `Workbook.ExportAsFixedFormat` / `Worksheet.ExportAsFixedFormat`), либо
  - добавить отдельный VBA‑макрос экспорта (вне рамок текущего аудита).

