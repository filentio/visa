## Windows Excel Runner (Python + pywin32)

Этот runner автоматизирует Excel через COM:

- копирует `.xlsm` в `work_dir` (чтобы не портить оригинал),
- заполняет лист `ввод данных`,
- прописывает абсолютные пути к PNG на листе `компании`,
- запускает VBA‑макрос `Module1.UpdateAllStampsFromConfig` (вставка печатей/подписей/логотипа),
- экспортирует указанные листы в PDF через `ExportAsFixedFormat` (без VBA),
- корректно закрывает Excel и чистит COM.

### Требования

- Windows 10/11 или Windows Server
- Установлен Microsoft Excel
- Python 3.10+

### Установка

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r excel_runner\requirements.txt
```

### Запуск

Payload из файла:

```powershell
python excel_runner\runner.py --payload excel_runner\config.example.json
```

Payload из stdin:

```powershell
type excel_runner\config.example.json | python excel_runner\runner.py --payload -
```

### Формат payload

См. `excel_runner/config.example.json`.

Ключевые поля:

- `template_path`: путь к исходному `.xlsm`
- `work_dir`: рабочая папка джобы (будут созданы `template.xlsm`, `assets/`, `output/`)
- `company.selected_company_name`: имя компании (ставится в `ввод данных`!`B11` и `C11`)
- `company.assets.*_png`: пути к исходным PNG (runner копирует их в `work_dir/assets/`)
- `export.*_template`: точные имена листов для экспорта PDF (должны совпадать с именами вкладок)

### Что делает runner по шагам (коротко)

1) `work_dir/`
   - копирует шаблон в `work_dir/template.xlsm`
   - создаёт `work_dir/output/` и `work_dir/assets/`
   - копирует PNG в `work_dir/assets/` (`logo.png`, `seal.png`, `director_sign.png`, `client_sign.png`)
2) Excel COM (headless):
   - `Visible=False`, `DisplayAlerts=False`, `ScreenUpdating=False`, `EnableEvents=False`, `AskToUpdateLinks=False`
3) Заполнение `ввод данных`:
   - `C2,C3,C4,C5,C6,C7,B8,C11,C15,C16,C17,C20` (и `C18`, если это не формула)
   - дополнительно выставляет `B11` = выбранная компания (это читает VBA)
4) Лист `компании`:
   - находит строку по совпадению названия компании в колонке A
   - записывает пути (как в извлечённом VBA `Module1.bas`):
     - D (4) = **seal/stamp**
     - E (5) = **logo**
     - F (6) = director sign
     - G (7) = client sign
5) Запускает макрос:
   - `excel.Run("Module1.UpdateAllStampsFromConfig")`
6) Экспорт PDF:
   - `Worksheet.ExportAsFixedFormat(Type=0, Filename=...)` для каждого листа из payload

### Важно про безопасность макросов

Если Excel настроен блокировать макросы (Mark of the Web / Trust Center), COM‑автоматизация не сможет “обойти” это надёжно.
Рекомендуется:

- хранить шаблон в **trusted location** Excel,
- или снять блокировку файла (Properties → Unblock),
- или настроить политику макросов для этого окружения runner’а.

### Формат ответа (stdout)

Успех:

```json
{
  "status": "ok",
  "output_dir": "C:\\\\temp\\\\visa_jobs\\\\job_123\\\\output",
  "pdf_files": ["Contract.pdf", "Bank_Statement_6m.pdf", "Insurance.pdf", "Salary_Certificate.pdf"]
}
```

Ошибка:

```json
{
  "status": "error",
  "message": "human readable",
  "details": "stacktrace/exception short"
}
```

