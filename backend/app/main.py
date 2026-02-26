from __future__ import annotations

from datetime import date
from typing import Any, Dict, Optional

import requests
from fastapi import Depends, FastAPI, HTTPException, Query, status
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import or_, select
from sqlalchemy.sql import func
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from . import models, schemas
from .db import get_db, init_db
from .queue import enqueue_job
from .security import require_internal_key
from .settings import settings
from .storage import presigned_get_url
from .utils import (
    country_display_from_issuing,
    generate_contract_number,
    issuing_country_from_mrz,
    new_id,
    random_start_date_within_last_6_months,
)


app = FastAPI(title="visa-backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def _startup() -> None:
    init_db()


def _fetch_cbr_rate(currency: str) -> float:
    """
    Best-effort FX source for MVP.
    Uses CBR daily JSON (if available). If it fails, caller must handle.
    """
    # CBR provides rates vs RUB.
    url = "https://www.cbr-xml-daily.ru/daily_json.js"
    r = requests.get(url, timeout=10)
    r.raise_for_status()
    data = r.json()
    valute = (data.get("Valute") or {}).get(currency)
    if not valute:
        raise RuntimeError(f"CBR rate not found for {currency}")
    value = float(valute["Value"])
    nominal = float(valute.get("Nominal", 1))
    return value / nominal


@app.post("/packages/generate", response_model=schemas.GeneratePackageOut)
def generate_package(body: schemas.GeneratePackageIn, db: Session = Depends(get_db)) -> schemas.GeneratePackageOut:
    company = db.execute(select(models.Company).where(models.Company.id == body.company_id)).scalar_one_or_none()
    if not company:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="company_id not found")

    mrz = "\n".join([ln for ln in [body.client.mrz_line1, body.client.mrz_line2] if ln and ln.strip()]) or None

    issuing = body.client.issuing_country
    if not issuing and mrz:
        issuing = issuing_country_from_mrz(mrz)
    country_display = country_display_from_issuing(issuing)

    # Address is not in the current request spec; generate placeholder.
    address = "RUSSIA, Moscow, 119087, Akademik Tupolev str.14 apt.430"

    # FX rate
    if body.job.fx_source == "manual":
        if body.job.fx_rate is None:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="fx_rate is required for manual fx")
        fx_rate = float(body.job.fx_rate)
    else:
        try:
            fx_rate = _fetch_cbr_rate(body.job.currency)
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=f"Failed to fetch CBR fx rate for {body.job.currency}: {e}",
            ) from e

    # Upsert client by passport_no + dob
    client = db.execute(
        select(models.Client).where(
            models.Client.passport_no == body.client.passport_no, models.Client.dob == body.client.dob
        )
    ).scalar_one_or_none()
    if client is None:
        client = models.Client(
            id=new_id(),
            full_name=body.client.full_name,
            passport_no=body.client.passport_no,
            dob=body.client.dob,
            mrz=mrz,
            issuing_country=issuing,
        )
        db.add(client)
    else:
        client.full_name = body.client.full_name
        client.mrz = mrz
        client.issuing_country = issuing

    # Package fields
    start_date = random_start_date_within_last_6_months()
    contract_start_date = start_date

    # Contract number must be unique.
    contract_number = None
    for _ in range(15):
        candidate = generate_contract_number()
        exists = db.execute(select(models.Package.id).where(models.Package.contract_number == candidate)).scalar_one_or_none()
        if not exists:
            contract_number = candidate
            break
    if not contract_number:
        raise HTTPException(status_code=500, detail="Failed to allocate unique contract number")

    package = models.Package(
        id=new_id(),
        status=models.PackageStatus.created,
        version_counter=0,
        client_id=client.id,
        company_id=company.id,
        currency=models.Currency(body.job.currency),
        fx_source=models.FxSource(body.job.fx_source),
        fx_rate=fx_rate,
        salary_rub=body.job.salary_rub,
        position=body.job.position,
        start_date=start_date,
        contract_start_date=contract_start_date,
        contract_number=contract_number,
        contract_template=body.templates.contract_template,
        insurance_template=body.templates.insurance_template,
        country_display=country_display,
        address=address,
    )
    db.add(package)

    # Allocate v1 for the initial generation.
    package.version_counter = 1
    job = models.Job(
        id=new_id(),
        package_id=package.id,
        status=models.JobStatus.queued,
        version=1,
        error_message=None,
        started_at=None,
        finished_at=None,
    )
    db.add(job)

    try:
        db.commit()
    except IntegrityError as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"DB integrity error: {e}") from e

    enqueue_job(job_id=job.id, package_id=package.id)
    return schemas.GeneratePackageOut(job_id=job.id, package_id=package.id)


@app.post("/packages/{package_id}/regenerate", response_model=schemas.RegenerateOut)
def regenerate_package(package_id: str, db: Session = Depends(get_db)) -> schemas.RegenerateOut:
    pkg = db.execute(select(models.Package).where(models.Package.id == package_id)).scalar_one_or_none()
    if not pkg:
        raise HTTPException(status_code=404, detail="package not found")

    # Allocate next version.
    next_version = int(getattr(pkg, "version_counter", 0)) + 1
    pkg.version_counter = next_version

    job = models.Job(
        id=new_id(),
        package_id=pkg.id,
        status=models.JobStatus.queued,
        version=next_version,
        error_message=None,
        started_at=None,
        finished_at=None,
    )
    db.add(job)
    db.commit()

    enqueue_job(job_id=job.id, package_id=pkg.id)
    return schemas.RegenerateOut(job_id=job.id, package_id=pkg.id, version=next_version)


@app.get("/companies", response_model=list[schemas.CompanyPublicOut])
def list_companies(db: Session = Depends(get_db)) -> list[schemas.CompanyPublicOut]:
    companies = db.execute(select(models.Company).order_by(models.Company.created_at.asc())).scalars().all()
    return [
        schemas.CompanyPublicOut(
            id=c.id,
            name=c.name,
        )
        for c in companies
    ]


@app.post("/internal/companies", dependencies=[Depends(require_internal_key)], response_model=schemas.CompanyOut)
def internal_create_company(body: schemas.CompanyCreateIn, db: Session = Depends(get_db)) -> schemas.CompanyOut:
    company = models.Company(
        id=new_id(),
        name=body.name,
        seal_key=body.seal_key,
        logo_key=body.logo_key,
        director_sign_key=body.director_sign_key,
        client_sign_key=body.client_sign_key,
    )
    db.add(company)
    try:
        db.commit()
    except IntegrityError as e:
        db.rollback()
        raise HTTPException(status_code=400, detail=f"Company create failed: {e}") from e
    return schemas.CompanyOut(
        company_id=company.id,
        name=company.name,
        seal_key=company.seal_key,
        logo_key=company.logo_key,
        director_sign_key=company.director_sign_key,
        client_sign_key=company.client_sign_key,
    )


@app.get("/jobs/{job_id}", response_model=schemas.JobStatusOut)
def get_job(job_id: str, db: Session = Depends(get_db)) -> schemas.JobStatusOut:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    return schemas.JobStatusOut(
        job_id=job.id,
        status=job.status.value,
        error_message=job.error_message,
        started_at=job.started_at.isoformat() if getattr(job, "started_at", None) else None,
        finished_at=job.finished_at.isoformat() if getattr(job, "finished_at", None) else None,
        version=int(getattr(job, "version", 1)),
    )


@app.get("/packages/{package_id}", response_model=schemas.PackageOut)
def get_package(package_id: str, db: Session = Depends(get_db)) -> schemas.PackageOut:
    pkg = db.execute(select(models.Package).where(models.Package.id == package_id)).scalar_one_or_none()
    if not pkg:
        raise HTTPException(status_code=404, detail="package not found")

    docs = (
        db.execute(select(models.Document).where(models.Document.package_id == pkg.id).order_by(models.Document.created_at.asc()))
        .scalars()
        .all()
    )
    client = db.execute(select(models.Client).where(models.Client.id == pkg.client_id)).scalar_one()
    company = db.execute(select(models.Company).where(models.Company.id == pkg.company_id)).scalar_one()

    def _mask_passport(p: str) -> str:
        s = str(p)
        digits = [i for i, ch in enumerate(s) if ch.isalnum()]
        if len(digits) <= 2:
            return "*" * len(s)
        keep = set(digits[-2:])
        out = []
        for i, ch in enumerate(s):
            if i in keep:
                out.append(ch)
            elif ch.isalnum():
                out.append("*")
            else:
                out.append(ch)
        return "".join(out)

    def _pkg_status() -> str:
        if pkg.status == models.PackageStatus.created:
            return "draft"
        if pkg.status == models.PackageStatus.generated:
            return "generated"
        return "error"

    def _doc_type(v: str) -> str:
        if v == "bank":
            return "bank_statement"
        if v in ("contract", "insurance", "salary", "bundle"):
            return v
        return "other"

    doc_out = []
    for d in docs:
        try:
            url = presigned_get_url(key=d.storage_key, expires_seconds=3600)
        except Exception:
            url = None
        doc_out.append(
            schemas.PackageDocumentOut(
                doc_type=_doc_type(d.doc_type.value),
                version=d.version,
                file_key=d.storage_key,
                created_at=d.created_at.isoformat(),
                presigned_url=url,
            )
        )

    return schemas.PackageOut(
        package_id=pkg.id,
        status=_pkg_status(),
        version_counter=int(getattr(pkg, "version_counter", 0)),
        client=schemas.PackageClientOut(
            client_id=client.id,
            full_name=client.full_name,
            passport_masked=_mask_passport(client.passport_no),
            dob=client.dob.isoformat(),
            issuing_country=client.issuing_country,
        ),
        company=schemas.PackageCompanyOut(company_id=company.id, name=company.name),
        documents=doc_out,
    )


@app.get("/packages/{package_id}/download", response_model=schemas.DownloadOut)
def download_package_zip(package_id: str, db: Session = Depends(get_db)) -> Dict[str, Any]:
    """
    Return a presigned URL for the latest bundle ZIP.
    """
    pkg = db.execute(select(models.Package).where(models.Package.id == package_id)).scalar_one_or_none()
    if not pkg:
        raise HTTPException(status_code=404, detail="package not found")

    # Prefer bundle with version == package.version_counter, fallback to max bundle version.
    bundle = None
    try:
        vc = int(getattr(pkg, "version_counter", 0))
    except Exception:
        vc = 0
    if vc > 0:
        bundle = (
            db.execute(
                select(models.Document).where(
                    models.Document.package_id == package_id,
                    models.Document.doc_type == models.DocumentType.bundle,
                    models.Document.version == vc,
                )
            )
            .scalars()
            .first()
        )
    if bundle is None:
        bundle = (
            db.execute(
                select(models.Document)
                .where(models.Document.package_id == package_id, models.Document.doc_type == models.DocumentType.bundle)
                .order_by(models.Document.version.desc())
            )
            .scalars()
            .first()
        )
    if not bundle:
        raise HTTPException(status_code=404, detail="bundle not found")
    url = presigned_get_url(key=bundle.storage_key, expires_seconds=3600)
    return {"url": url}


@app.get("/clients", response_model=list[schemas.ClientSearchItem])
def search_clients(
    query: str | None = Query(default=None),
    limit: int = Query(default=30, ge=1, le=100),
    db: Session = Depends(get_db),
) -> list[schemas.ClientSearchItem]:
    q = (query or "").strip()
    stmt = select(models.Client).order_by(models.Client.created_at.desc()).limit(limit)
    if q:
        stmt = (
            select(models.Client)
            .where(or_(models.Client.full_name.ilike(f"%{q}%"), models.Client.passport_no.like(f"%{q}%")))
            .order_by(models.Client.created_at.desc())
            .limit(limit)
        )
    clients = db.execute(stmt).scalars().all()

    def _mask_passport(p: str) -> str:
        s = str(p)
        digits = [i for i, ch in enumerate(s) if ch.isalnum()]
        if len(digits) <= 2:
            return "*" * len(s)
        keep = set(digits[-2:])
        out = []
        for i, ch in enumerate(s):
            if i in keep:
                out.append(ch)
            elif ch.isalnum():
                out.append("*")
            else:
                out.append(ch)
        return "".join(out)

    return [
        schemas.ClientSearchItem(
            client_id=c.id,
            full_name=c.full_name,
            passport_masked=_mask_passport(c.passport_no),
            dob=c.dob.isoformat(),
            issuing_country=c.issuing_country,
            created_at=c.created_at.isoformat(),
        )
        for c in clients
    ]


@app.get("/clients/{client_id}", response_model=schemas.ClientDetail)
def get_client(client_id: str, db: Session = Depends(get_db)) -> schemas.ClientDetail:
    c = db.execute(select(models.Client).where(models.Client.id == client_id)).scalar_one_or_none()
    if not c:
        raise HTTPException(status_code=404, detail="client not found")

    def _mask_passport(p: str) -> str:
        s = str(p)
        digits = [i for i, ch in enumerate(s) if ch.isalnum()]
        if len(digits) <= 2:
            return "*" * len(s)
        keep = set(digits[-2:])
        out = []
        for i, ch in enumerate(s):
            if i in keep:
                out.append(ch)
            elif ch.isalnum():
                out.append("*")
            else:
                out.append(ch)
        return "".join(out)

    return schemas.ClientDetail(
        client_id=c.id,
        full_name=c.full_name,
        passport_masked=_mask_passport(c.passport_no),
        dob=c.dob.isoformat(),
        issuing_country=c.issuing_country,
        country_display=country_display_from_issuing(c.issuing_country),
        created_at=c.created_at.isoformat(),
    )


@app.get("/clients/{client_id}/packages", response_model=list[schemas.ClientPackageItem])
def get_client_packages(client_id: str, db: Session = Depends(get_db)) -> list[schemas.ClientPackageItem]:
    # Ensure client exists
    c = db.execute(select(models.Client.id).where(models.Client.id == client_id)).scalar_one_or_none()
    if not c:
        raise HTTPException(status_code=404, detail="client not found")

    pkgs = (
        db.execute(select(models.Package).where(models.Package.client_id == client_id).order_by(models.Package.created_at.desc()))
        .scalars()
        .all()
    )

    company_ids = {p.company_id for p in pkgs}
    companies = (
        db.execute(select(models.Company).where(models.Company.id.in_(company_ids))).scalars().all() if company_ids else []
    )
    company_map = {co.id: co for co in companies}

    def _pkg_status(p: models.Package) -> str:
        if p.status == models.PackageStatus.created:
            return "draft"
        if p.status == models.PackageStatus.generated:
            return "generated"
        return "error"

    out: list[schemas.ClientPackageItem] = []
    for p in pkgs:
        co = company_map.get(p.company_id)
        out.append(
            schemas.ClientPackageItem(
                package_id=p.id,
                status=_pkg_status(p),  # type: ignore[arg-type]
                version_counter=int(getattr(p, "version_counter", 0)),
                company=schemas.PackageCompanyOut(company_id=p.company_id, name=co.name if co else p.company_id),
                created_at=p.created_at.isoformat(),
                updated_at=p.updated_at.isoformat(),
            )
        )
    return out


@app.get("/files/presign", response_model=schemas.DownloadOut)
def presign_file(key: str) -> Dict[str, Any]:
    url = presigned_get_url(key=key, expires_seconds=3600)
    return {"url": url}


def _default_bank_template(currency: models.Currency) -> str:
    # Based on workbook sheet names extracted in `extracted_vba/workbook_map.md`
    if currency == models.Currency.USD:
        return "т-банк 2 (6 мес) $"
    return "т-банк 2 (6 мес)"


def _default_currency_symbol(currency: models.Currency) -> str:
    if currency == models.Currency.USD:
        return "$"
    if currency == models.Currency.AED:
        return "AED"
    return "$"


@app.get("/internal/jobs/{job_id}/payload", dependencies=[Depends(require_internal_key)], response_model=schemas.InternalJobPayloadOut)
def internal_job_payload(job_id: str, db: Session = Depends(get_db)) -> schemas.InternalJobPayloadOut:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    pkg = db.execute(select(models.Package).where(models.Package.id == job.package_id)).scalar_one()
    client = db.execute(select(models.Client).where(models.Client.id == pkg.client_id)).scalar_one()
    company = db.execute(select(models.Company).where(models.Company.id == pkg.company_id)).scalar_one()

    currency_symbol = _default_currency_symbol(pkg.currency)
    bank_template = _default_bank_template(pkg.currency)
    salary_template = "Salary упрошенная"

    # Worker builds runner payload with local paths; backend gives only storage keys and values.
    return schemas.InternalJobPayloadOut(
        job_id=job.id,
        package_id=pkg.id,
        template_key=settings.template_key,
        company={
            "company_id": company.id,
            "selected_company_name": company.name,
            "assets": {
                "logo_key": company.logo_key,
                "seal_key": company.seal_key,
                "director_sign_key": company.director_sign_key,
                "client_sign_key": company.client_sign_key,
            },
        },
        client={
            "full_name": client.full_name,
            "passport_no": client.passport_no,
            "dob": client.dob.isoformat(),
            "address": pkg.address,
            "country_display": pkg.country_display,
        },
        job={
            "version": int(getattr(job, "version", 1)),
            "currency_symbol": currency_symbol,
            "fx_rate": float(pkg.fx_rate),
            "salary_rub": float(pkg.salary_rub),
            "position": pkg.position,
            "contract_start_date": pkg.contract_start_date.isoformat(),
            "contract_number": pkg.contract_number,
        },
        export={
            "contract_template": pkg.contract_template,
            "bank_template": bank_template,
            "insurance_template": pkg.insurance_template,
            "salary_template": salary_template,
            "output_files": {
                "contract": "Contract.pdf",
                "bank": "Bank_Statement_6m.pdf",
                "insurance": "Insurance.pdf",
                "salary": "Salary_Certificate.pdf",
            },
        },
    )


@app.post("/internal/jobs/{job_id}/start", dependencies=[Depends(require_internal_key)])
def internal_job_start(job_id: str, db: Session = Depends(get_db)) -> Dict[str, Any]:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")

    if job.status != models.JobStatus.queued:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="job is not queued")

    job.status = models.JobStatus.running
    job.started_at = func.now()
    job.error_message = None
    db.commit()
    return {"status": "ok", "job_id": job.id}


@app.post("/internal/jobs/{job_id}/complete", dependencies=[Depends(require_internal_key)])
def internal_job_complete(job_id: str, body: schemas.InternalCompleteIn, db: Session = Depends(get_db)) -> Dict[str, Any]:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    if job.status != models.JobStatus.running:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="job is not running")

    pkg = db.execute(select(models.Package).where(models.Package.id == job.package_id)).scalar_one()

    expected_version = int(getattr(job, "version", 1))

    # Create document rows (versioned).
    for f in body.files:
        doc_type = f.get("doc_type")
        storage_key = f.get("storage_key")
        filename = f.get("filename")
        content_type = f.get("content_type") or "application/octet-stream"
        if not doc_type or not storage_key or not filename:
            raise HTTPException(status_code=400, detail="Invalid file item")

        try:
            doc_type_enum = models.DocumentType(str(doc_type))
        except Exception:
            raise HTTPException(status_code=400, detail=f"Invalid doc_type: {doc_type!r}")

        file_version = f.get("version")
        if file_version is not None and int(file_version) != expected_version:
            raise HTTPException(status_code=400, detail="version mismatch")

        doc = models.Document(
            id=new_id(),
            package_id=pkg.id,
            doc_type=doc_type_enum,
            version=expected_version,
            filename=str(filename),
            storage_key=str(storage_key),
            content_type=str(content_type),
        )
        db.add(doc)

    job.status = models.JobStatus.done
    job.error_message = None
    job.finished_at = func.now()
    pkg.status = models.PackageStatus.generated

    db.commit()
    return {"status": "ok", "job_id": job.id, "package_id": pkg.id}


@app.post("/internal/jobs/{job_id}/fail", dependencies=[Depends(require_internal_key)])
def internal_job_fail(job_id: str, body: schemas.InternalFailIn, db: Session = Depends(get_db)) -> Dict[str, Any]:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    if job.status != models.JobStatus.running:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="job is not running")

    pkg = db.execute(select(models.Package).where(models.Package.id == job.package_id)).scalar_one()

    job.status = models.JobStatus.error
    job.error_message = body.error_message[:2000]
    job.finished_at = func.now()
    pkg.status = models.PackageStatus.error
    db.commit()
    return {"status": "ok", "job_id": job.id}

