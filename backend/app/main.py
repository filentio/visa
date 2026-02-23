from __future__ import annotations

from datetime import date
from typing import Any, Dict, Optional

import requests
from fastapi import Depends, FastAPI, HTTPException, status
from sqlalchemy import select
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

    issuing = body.client.issuing_country
    if not issuing and body.client.mrz:
        issuing = issuing_country_from_mrz(body.client.mrz)
    country_display = country_display_from_issuing(issuing)

    # Address is not in the current request spec; generate placeholder if not provided.
    address = body.address or "RUSSIA, Moscow, 119087, Akademik Tupolev str.14 apt.430"

    # FX rate
    if body.fx_source == "manual":
        if body.fx_rate is None:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="fx_rate is required for manual fx")
        fx_rate = float(body.fx_rate)
    else:
        try:
            fx_rate = _fetch_cbr_rate(body.currency)
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=f"Failed to fetch CBR fx rate for {body.currency}: {e}",
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
            mrz=body.client.mrz,
            issuing_country=issuing,
        )
        db.add(client)
    else:
        client.full_name = body.client.full_name
        client.mrz = body.client.mrz
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
        client_id=client.id,
        company_id=company.id,
        currency=models.Currency(body.currency),
        fx_source=models.FxSource(body.fx_source),
        fx_rate=fx_rate,
        salary_rub=body.salary_rub,
        position=body.position,
        start_date=start_date,
        contract_start_date=contract_start_date,
        contract_number=contract_number,
        contract_template=body.contract_template,
        insurance_template=body.insurance_template,
        country_display=country_display,
        address=address,
    )
    db.add(package)

    job = models.Job(id=new_id(), package_id=package.id, status=models.JobStatus.queued, error_message=None)
    db.add(job)

    try:
        db.commit()
    except IntegrityError as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"DB integrity error: {e}") from e

    enqueue_job(job_id=job.id, package_id=package.id)
    return schemas.GeneratePackageOut(job_id=job.id, package_id=package.id)


@app.get("/companies", response_model=list[schemas.CompanyOut])
def list_companies(db: Session = Depends(get_db)) -> list[schemas.CompanyOut]:
    companies = db.execute(select(models.Company).order_by(models.Company.created_at.asc())).scalars().all()
    return [
        schemas.CompanyOut(
            company_id=c.id,
            name=c.name,
            seal_key=c.seal_key,
            logo_key=c.logo_key,
            director_sign_key=c.director_sign_key,
            client_sign_key=c.client_sign_key,
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
    return schemas.JobStatusOut(job_id=job.id, status=job.status.value, error_message=job.error_message)


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
    return schemas.PackageOut(
        package_id=pkg.id,
        status=pkg.status.value,
        client_id=pkg.client_id,
        company_id=pkg.company_id,
        currency=pkg.currency.value,
        fx_source=pkg.fx_source.value,
        fx_rate=float(pkg.fx_rate),
        start_date=pkg.start_date,
        contract_start_date=pkg.contract_start_date,
        contract_number=pkg.contract_number,
        contract_template=pkg.contract_template,
        insurance_template=pkg.insurance_template,
        country_display=pkg.country_display,
        address=pkg.address,
        documents=[
            schemas.DocumentOut(
                doc_type=d.doc_type.value,
                version=d.version,
                filename=d.filename,
                storage_key=d.storage_key,
            )
            for d in docs
        ],
    )


@app.get("/packages/{package_id}/download")
def download_package_zip(package_id: str, db: Session = Depends(get_db)) -> Dict[str, Any]:
    """
    Return a presigned URL for the latest bundle ZIP.
    """
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
    return {"package_id": package_id, "bundle_key": bundle.storage_key, "url": url}


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


@app.post("/internal/jobs/{job_id}/complete", dependencies=[Depends(require_internal_key)])
def internal_job_complete(job_id: str, body: schemas.InternalCompleteIn, db: Session = Depends(get_db)) -> Dict[str, Any]:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")

    pkg = db.execute(select(models.Package).where(models.Package.id == job.package_id)).scalar_one()

    # Create document rows (version=1 MVP).
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

        doc = models.Document(
            id=new_id(),
            package_id=pkg.id,
            doc_type=doc_type_enum,
            version=int(f.get("version") or 1),
            filename=str(filename),
            storage_key=str(storage_key),
            content_type=str(content_type),
        )
        db.add(doc)

    job.status = models.JobStatus.done
    job.error_message = None
    pkg.status = models.PackageStatus.generated

    db.commit()
    return {"status": "ok", "job_id": job.id, "package_id": pkg.id}


@app.post("/internal/jobs/{job_id}/fail", dependencies=[Depends(require_internal_key)])
def internal_job_fail(job_id: str, body: schemas.InternalFailIn, db: Session = Depends(get_db)) -> Dict[str, Any]:
    job = db.execute(select(models.Job).where(models.Job.id == job_id)).scalar_one_or_none()
    if not job:
        raise HTTPException(status_code=404, detail="job not found")

    pkg = db.execute(select(models.Package).where(models.Package.id == job.package_id)).scalar_one()

    job.status = models.JobStatus.error
    job.error_message = body.error_message[:2000]
    pkg.status = models.PackageStatus.error
    db.commit()
    return {"status": "ok", "job_id": job.id}

