## Windows worker (Redis consumer) — queue → Excel runner → S3 → job status

Этот воркер запускается **на Windows**, где установлен Excel, и делает полный цикл:

1) Берёт сообщение из Redis очереди `visa_jobs` (минимум данных: `job_id`, `package_id`)
2) Запрашивает у backend internal API полный payload: `GET /internal/jobs/{job_id}/payload`
3) Скачивает из S3/MinIO:
   - шаблон `template_key` → `C:\temp\visa_jobs\<job_id>\template.xlsm`
   - PNG ассеты компании → `...\assets\*.png`
4) Формирует JSON payload для `excel_runner/runner.py`
5) Запускает runner как subprocess
6) Загружает PDFs и `bundle.zip` в storage
7) Отмечает job `done` через `POST /internal/jobs/{job_id}/complete` (или `fail`)

### Установка

В одном окружении нужно поставить зависимости worker + runner:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r windows_worker\requirements.txt
pip install -r excel_runner\requirements.txt
```

### Переменные окружения

- `BACKEND_BASE_URL` (пример: `http://127.0.0.1:8000`)
- `INTERNAL_API_KEY` (совпадает с `INTERNAL_API_KEY` backend)
- `REDIS_URL` (пример: `redis://127.0.0.1:6379/0`)
- `S3_ENDPOINT_URL` (MinIO: `http://127.0.0.1:9000`)
- `S3_BUCKET` (пример: `visa`)
- `S3_ACCESS_KEY_ID`
- `S3_SECRET_ACCESS_KEY`
- `S3_REGION` (обычно `us-east-1`)
- `WORK_ROOT` (опционально, default: `C:\temp\visa_jobs`)
- `RUNNER_SCRIPT` (опционально, default: `excel_runner\runner.py` относительно корня репо)

### Запуск

```powershell
python windows_worker\worker.py
```

### Важно (Excel Trust Center / макросы)

Excel может блокировать макросы (Trust Center / Mark-of-the-Web). Для стабильной автоматизации:

- используйте **Trusted Location** для папки с шаблонами/`work_dir`, или
- снимайте блокировку файла `.xlsm`, или
- настройте политику макросов в окружении worker’а.

### Логи и персональные данные

Worker **не должен логировать** паспорт/ФИО. В логах печатаются только `job_id`, `package_id` и технические ошибки.

