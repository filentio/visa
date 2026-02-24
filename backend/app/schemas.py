from __future__ import annotations

from datetime import date
from typing import Literal, Optional

from pydantic import BaseModel, Field


class ClientIn(BaseModel):
    full_name: str
    passport_no: str
    dob: date
    mrz: Optional[str] = None
    issuing_country: Optional[str] = None


class GeneratePackageIn(BaseModel):
    client: ClientIn

    company_id: str

    salary_rub: float
    position: str

    currency: Literal["USD", "AED"]
    fx_source: Literal["manual", "cbr"]
    fx_rate: Optional[float] = None

    contract_template: Literal["договор", "договор2"]
    insurance_template: Literal["страховка", "РГС"]

    # Optional: allow passing address explicitly; otherwise backend generates a placeholder.
    address: Optional[str] = None


class GeneratePackageOut(BaseModel):
    job_id: str
    package_id: str
    version: int | None = None


class RegenerateOut(BaseModel):
    job_id: str
    package_id: str
    version: int


class JobStatusOut(BaseModel):
    job_id: str
    status: str
    error_message: Optional[str] = None


class DocumentOut(BaseModel):
    doc_type: str
    version: int
    filename: str
    storage_key: str


class PackageOut(BaseModel):
    package_id: str
    status: str
    client_id: str
    company_id: str

    currency: str
    fx_source: str
    fx_rate: float

    start_date: date
    contract_start_date: date
    contract_number: str

    contract_template: str
    insurance_template: str

    country_display: str
    address: str

    documents: list[DocumentOut] = Field(default_factory=list)


class InternalJobPayloadOut(BaseModel):
    job_id: str
    package_id: str

    template_key: str

    company: dict
    client: dict
    job: dict
    export: dict


class InternalCompleteIn(BaseModel):
    files: list[dict]


class InternalFailIn(BaseModel):
    error_message: str


class CompanyCreateIn(BaseModel):
    name: str
    seal_key: str
    logo_key: str
    director_sign_key: str
    client_sign_key: str


class CompanyOut(BaseModel):
    company_id: str
    name: str
    seal_key: str
    logo_key: str
    director_sign_key: str
    client_sign_key: str

