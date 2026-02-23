## Excel VBA extractor (для Cursor)

Cursor не умеет читать `.xlsm` напрямую. Этот инструмент извлекает VBA‑модули в текстовые файлы (`.bas/.cls/.frm`) и строит карту книги (листы + именованные диапазоны, если получается).

### Установка зависимостей

Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r tools\excel_extract\requirements.txt
```

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r tools/excel_extract/requirements.txt
```

### Запуск

Пример (как в задаче):

```bash
python tools/excel_extract/extract_vba.py --input "шаблон 12.09.xlsm" --out extracted_vba
```

### Где лежат результаты

Скрипт создаёт/перезаписывает папку вывода (например `extracted_vba/`) и генерирует:

- `extracted_vba/modules/*.bas` — обычные модули
- `extracted_vba/classes/*.cls` — классы (в т.ч. `ThisWorkbook`, модули листов и т.п.)
- `extracted_vba/forms/*.frm` — UserForms (если есть)
- `extracted_vba/meta/olevba_report.txt` — краткий отчёт по извлечению/анализу
- `extracted_vba/workbook_map.json` — машина‑читаемая карта книги
- `extracted_vba/workbook_map.md` — человеко‑читаемое описание

Эти результаты можно коммитить в репозиторий, и Cursor сможет анализировать VBA как обычный код.

### Частые ошибки

- **Ошибка про шифрование / пароль** (`encrypted`, `password`, и т.п.)
  - Книга защищена. Нужна **незашифрованная/без пароля** копия `.xlsm`.
- **`no VBA found`**
  - В книге нет VBA (либо VBA удалён/отключён), либо макросы не сохранены в файле.

