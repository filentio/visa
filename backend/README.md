## Backend API (FastAPI) — Packages + Jobs + Storage links + Redis queue

Этот backend делает:

- `POST /packages/generate` → создаёт client+package+job и кладёт сообщение в Redis очередь `visa_jobs`
- `GET /jobs/{job_id}` → статус job
- `GET /packages/{package_id}` → мета пакета + документы
- `GET /packages/{package_id}/download` → presigned URL на ZIP (bundle)
- внутренние эндпоинты для Windows worker (по API key):
  - `GET /internal/jobs/{job_id}/payload`
  - `POST /internal/jobs/{job_id}/complete`
  - `POST /internal/jobs/{job_id}/fail`

### Переменные окружения

- **Database**
  - `DATABASE_URL` (пример: `postgresql+psycopg2://postgres:postgres@localhost:5432/visa`)
- **Redis**
  - `REDIS_URL` (пример: `redis://localhost:6379/0`)
- **S3/MinIO**
  - `S3_ENDPOINT_URL` (пример MinIO: `http://localhost:9000`)
  - `S3_BUCKET` (пример: `visa`)
  - `S3_ACCESS_KEY_ID`
  - `S3_SECRET_ACCESS_KEY`
  - `S3_REGION` (опционально; для MinIO можно `us-east-1`)
- **Internal API auth**
  - `INTERNAL_API_KEY` (worker передаёт в заголовке `X-Internal-Api-Key`)

### Запуск (локально)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r backend/requirements.txt

export DATABASE_URL="postgresql+psycopg2://postgres:postgres@localhost:5432/visa"
export REDIS_URL="redis://localhost:6379/0"
export S3_ENDPOINT_URL="http://localhost:9000"
export S3_BUCKET="visa"
export S3_ACCESS_KEY_ID="minioadmin"
export S3_SECRET_ACCESS_KEY="minioadmin"
export INTERNAL_API_KEY="change-me"

uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
```

### Storage keys (конвенция)

- шаблон: `templates/template.xlsm`
- ассеты компаний: ключи лежат в таблице `companies` (`logo_key`, `seal_key`, `director_sign_key`, `client_sign_key`)
- результаты:
  - `packages/{package_id}/Contract_v1.pdf`
  - `packages/{package_id}/Bank_Statement_6m_v1.pdf`
  - `packages/{package_id}/Insurance_v1.pdf`
  - `packages/{package_id}/Salary_Certificate_v1.pdf`
  - `packages/{package_id}/bundle_v1.zip`

