from __future__ import annotations

from contextlib import contextmanager
from typing import Iterator

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import Session, declarative_base, sessionmaker

from .settings import settings


engine = create_engine(settings.database_url, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False)

Base = declarative_base()


def get_db() -> Iterator[Session]:
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    # MVP: create tables automatically.
    from . import models  # noqa: F401

    Base.metadata.create_all(bind=engine)
    _ensure_schema()


def _ensure_schema() -> None:
    """
    Minimal "migration" for MVP without Alembic.
    Adds new columns and enum values if the DB already exists.
    """
    insp = inspect(engine)
    if not insp.has_table("jobs") or not insp.has_table("packages"):
        return

    jobs_cols = {c["name"] for c in insp.get_columns("jobs")}
    packages_cols = {c["name"] for c in insp.get_columns("packages")}

    with engine.begin() as conn:
        # jobs: started_at, finished_at, version
        if "started_at" not in jobs_cols:
            conn.execute(text('ALTER TABLE "jobs" ADD COLUMN "started_at" TIMESTAMPTZ NULL'))
        if "finished_at" not in jobs_cols:
            conn.execute(text('ALTER TABLE "jobs" ADD COLUMN "finished_at" TIMESTAMPTZ NULL'))
        if "version" not in jobs_cols:
            conn.execute(text('ALTER TABLE "jobs" ADD COLUMN "version" INTEGER NOT NULL DEFAULT 1'))

        # packages: version_counter
        if "version_counter" not in packages_cols:
            conn.execute(text('ALTER TABLE "packages" ADD COLUMN "version_counter" INTEGER NOT NULL DEFAULT 0'))

        # Ensure enum value 'running' exists for job status (Postgres).
        if engine.dialect.name == "postgresql":
            for type_name in ("jobstatus", "jobstatus_enum", "job_status"):
                # Never let enum alteration abort the surrounding transaction.
                conn.execute(
                    text(
                        f"""
DO $$
BEGIN
  ALTER TYPE {type_name} ADD VALUE 'running';
EXCEPTION
  WHEN duplicate_object THEN NULL;
  WHEN undefined_object THEN NULL;
END $$;
"""
                    )
                )

