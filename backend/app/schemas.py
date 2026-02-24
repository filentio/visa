from __future__ import annotations

from datetime import date
from typing import Literal, Optional

from pydantic import BaseModel, Field


class GenerateClientIn(BaseModel):
    full_name: str
    passport_no: str
    dob: date
    mrz_line1: Optional[str] = None
    mrz_line2: Optional[str] = None
    issuing_country: Optional[str] = None


class GenerateJobIn(BaseModel):
    position: str
    salary_rub: float
    currency: Literal["USD", "AED"]
    fx_source: Literal["manual", "cbr"]
    fx_rate: Optional[float] = None


class GenerateTemplatesIn(BaseModel):
    contract_template: Literal["договор", "договор2"]
    insurance_template: Literal["страховка", "РГС"]
    bank_template: Literal["т-банк 2 (6 мес) $"]
    salary_template: Literal["Salary упрошенная"]


class GeneratePackageIn(BaseModel):
    client: GenerateClientIn
    company_id: str
    job: GenerateJobIn
    templates: GenerateTemplatesIn


class GeneratePackageOut(BaseModel):
    job_id: str
    package_id: str


class RegenerateOut(BaseModel):
    job_id: str
    package_id: str
    version: int


class JobStatusOut(BaseModel):
    job_id: str
    status: Literal["queued", "running", "done", "error"]
    error_message: Optional[str] = None
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    version: Optional[int] = None


class CompanyPublicOut(BaseModel):
    id: str
    name: str


class PackageDocumentOut(BaseModel):
    doc_type: Literal["contract", "bank_statement", "insurance", "salary", "bundle", "other"]
    version: int
    file_key: str
    created_at: str
    presigned_url: Optional[str] = None


class PackageClientOut(BaseModel):
    client_id: str
    full_name: str
    passport_masked: str
    dob: str
    issuing_country: Optional[str] = None


class PackageCompanyOut(BaseModel):
    company_id: str
    name: str


class PackageOut(BaseModel):
    package_id: str
    status: Literal["draft", "generated", "error"]
    version_counter: int
    client: PackageClientOut
    company: PackageCompanyOut
    documents: list[PackageDocumentOut] = Field(default_factory=list)


class DownloadOut(BaseModel):
    url: str


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

