from __future__ import annotations

import argparse
import re
from pathlib import Path

import boto3
from sqlalchemy import select

from backend.app.db import SessionLocal, init_db
from backend.app.models import Company
from backend.app.settings import settings
from backend.app.utils import new_id


def slugify(name: str) -> str:
    s = name.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    s = s.strip("-")
    return s or "company"


def s3_client():
    return boto3.client(
        "s3",
        endpoint_url=settings.s3_endpoint_url,
        aws_access_key_id=settings.s3_access_key_id,
        aws_secret_access_key=settings.s3_secret_access_key,
        region_name=settings.s3_region,
    )


def ensure_bucket(s3, bucket: str) -> None:
    try:
        s3.head_bucket(Bucket=bucket)
    except Exception:
        try:
            s3.create_bucket(Bucket=bucket)
        except Exception:
            pass


def upload_file(s3, *, bucket: str, key: str, path: Path, content_type: str) -> None:
    s3.upload_file(str(path), bucket, key, ExtraArgs={"ContentType": content_type})


def main() -> int:
    p = argparse.ArgumentParser(description="Bootstrap demo storage + company record for visa backend.")
    p.add_argument("--template", default=str(Path("шаблон 12.09.xlsm")), help="Path to template .xlsm")
    p.add_argument("--company-name", default="STARLINK LLC", help="Company name")
    p.add_argument("--logo", help="Path to logo PNG")
    p.add_argument("--seal", help="Path to seal/stamp PNG")
    p.add_argument("--director-sign", help="Path to director signature PNG")
    p.add_argument("--client-sign", help="Path to client signature PNG")
    args = p.parse_args()

    init_db()

    template_path = Path(args.template)
    if not template_path.exists():
        raise SystemExit(f"Template not found: {template_path}")

    s3 = s3_client()
    ensure_bucket(s3, settings.s3_bucket)

    # Upload template to the key expected by workers.
    upload_file(
        s3,
        bucket=settings.s3_bucket,
        key=settings.template_key,
        path=template_path,
        content_type="application/vnd.ms-excel.sheet.macroEnabled.12",
    )

    prefix = f"companies/{slugify(args.company_name)}"

    def _upload_png(local: str | None, key_name: str) -> str:
        if not local:
            # Allow creating company without uploading assets (worker will fail if keys missing).
            return f"{prefix}/{key_name}.png"
        pth = Path(local)
        if not pth.exists():
            raise SystemExit(f"Asset not found: {pth}")
        key = f"{prefix}/{key_name}.png"
        upload_file(s3, bucket=settings.s3_bucket, key=key, path=pth, content_type="image/png")
        return key

    logo_key = _upload_png(args.logo, "logo")
    seal_key = _upload_png(args.seal, "seal")
    director_key = _upload_png(args.director_sign, "director_sign")
    client_key = _upload_png(args.client_sign, "client_sign")

    with SessionLocal() as db:
        existing = db.execute(select(Company).where(Company.name == args.company_name)).scalar_one_or_none()
        if existing:
            existing.logo_key = logo_key
            existing.seal_key = seal_key
            existing.director_sign_key = director_key
            existing.client_sign_key = client_key
            company = existing
        else:
            company = Company(
                id=new_id(),
                name=args.company_name,
                logo_key=logo_key,
                seal_key=seal_key,
                director_sign_key=director_key,
                client_sign_key=client_key,
            )
            db.add(company)
        db.commit()

    print("OK")
    print(f"- template_key: {settings.template_key}")
    print(f"- company_id: {company.id}")
    print(f"- assets:")
    print(f"  - logo_key: {logo_key}")
    print(f"  - seal_key: {seal_key}")
    print(f"  - director_sign_key: {director_key}")
    print(f"  - client_sign_key: {client_key}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

