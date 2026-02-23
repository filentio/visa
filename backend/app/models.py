from __future__ import annotations

import enum
from datetime import date, datetime

from sqlalchemy import Date, DateTime, Enum, ForeignKey, Integer, Numeric, String, Text, UniqueConstraint, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .db import Base


class JobStatus(str, enum.Enum):
    queued = "queued"
    running = "running"
    done = "done"
    error = "error"


class PackageStatus(str, enum.Enum):
    created = "created"
    generated = "generated"
    error = "error"


class FxSource(str, enum.Enum):
    manual = "manual"
    cbr = "cbr"


class Currency(str, enum.Enum):
    USD = "USD"
    AED = "AED"


class DocumentType(str, enum.Enum):
    contract = "contract"
    bank = "bank"
    insurance = "insurance"
    salary = "salary"
    bundle = "bundle"


class Client(Base):
    __tablename__ = "clients"

    id: Mapped[str] = mapped_column(String(36), primary_key=True)
    full_name: Mapped[str] = mapped_column(String(256), nullable=False)
    passport_no: Mapped[str] = mapped_column(String(64), nullable=False)
    dob: Mapped[date] = mapped_column(Date, nullable=False)
    mrz: Mapped[str | None] = mapped_column(Text, nullable=True)
    issuing_country: Mapped[str | None] = mapped_column(String(64), nullable=True)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False
    )

    __table_args__ = (UniqueConstraint("passport_no", "dob", name="uq_clients_passport_dob"),)


class Company(Base):
    __tablename__ = "companies"

    id: Mapped[str] = mapped_column(String(36), primary_key=True)
    name: Mapped[str] = mapped_column(String(256), nullable=False, unique=True)

    seal_key: Mapped[str] = mapped_column(String(512), nullable=False)
    logo_key: Mapped[str] = mapped_column(String(512), nullable=False)
    director_sign_key: Mapped[str] = mapped_column(String(512), nullable=False)
    client_sign_key: Mapped[str] = mapped_column(String(512), nullable=False)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False
    )


class Package(Base):
    __tablename__ = "packages"

    id: Mapped[str] = mapped_column(String(36), primary_key=True)
    status: Mapped[PackageStatus] = mapped_column(Enum(PackageStatus), nullable=False, default=PackageStatus.created)

    client_id: Mapped[str] = mapped_column(ForeignKey("clients.id"), nullable=False)
    company_id: Mapped[str] = mapped_column(ForeignKey("companies.id"), nullable=False)

    currency: Mapped[Currency] = mapped_column(Enum(Currency), nullable=False)
    fx_source: Mapped[FxSource] = mapped_column(Enum(FxSource), nullable=False)
    fx_rate: Mapped[float] = mapped_column(Numeric(18, 6), nullable=False)

    salary_rub: Mapped[float] = mapped_column(Numeric(18, 2), nullable=False)
    position: Mapped[str] = mapped_column(String(256), nullable=False)

    start_date: Mapped[date] = mapped_column(Date, nullable=False)
    contract_start_date: Mapped[date] = mapped_column(Date, nullable=False)
    contract_number: Mapped[str] = mapped_column(String(64), nullable=False, unique=True)

    contract_template: Mapped[str] = mapped_column(String(128), nullable=False)
    insurance_template: Mapped[str] = mapped_column(String(128), nullable=False)

    country_display: Mapped[str] = mapped_column(String(256), nullable=False)
    address: Mapped[str] = mapped_column(Text, nullable=False)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False
    )

    client: Mapped[Client] = relationship("Client")
    company: Mapped[Company] = relationship("Company")
    jobs: Mapped[list["Job"]] = relationship("Job", back_populates="package")
    documents: Mapped[list["Document"]] = relationship("Document", back_populates="package")


class Job(Base):
    __tablename__ = "jobs"

    id: Mapped[str] = mapped_column(String(36), primary_key=True)
    package_id: Mapped[str] = mapped_column(ForeignKey("packages.id"), nullable=False)
    status: Mapped[JobStatus] = mapped_column(Enum(JobStatus), nullable=False, default=JobStatus.queued)
    error_message: Mapped[str | None] = mapped_column(Text, nullable=True)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False
    )

    package: Mapped[Package] = relationship("Package", back_populates="jobs")


class Document(Base):
    __tablename__ = "documents"

    id: Mapped[str] = mapped_column(String(36), primary_key=True)
    package_id: Mapped[str] = mapped_column(ForeignKey("packages.id"), nullable=False)

    doc_type: Mapped[DocumentType] = mapped_column(Enum(DocumentType), nullable=False)
    version: Mapped[int] = mapped_column(Integer, nullable=False, default=1)

    filename: Mapped[str] = mapped_column(String(256), nullable=False)
    storage_key: Mapped[str] = mapped_column(String(512), nullable=False)
    content_type: Mapped[str] = mapped_column(String(128), nullable=False)

    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    package: Mapped[Package] = relationship("Package", back_populates="documents")

    __table_args__ = (UniqueConstraint("package_id", "doc_type", "version", name="uq_documents_pkg_type_ver"),)

