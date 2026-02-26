from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import boto3

from .settings import settings


@dataclass(frozen=True)
class Storage:
    client: object
    bucket: str


def get_storage() -> Storage:
    s3 = boto3.client(
        "s3",
        endpoint_url=settings.s3_endpoint_url,
        aws_access_key_id=settings.s3_access_key_id,
        aws_secret_access_key=settings.s3_secret_access_key,
        region_name=settings.s3_region,
    )
    return Storage(client=s3, bucket=settings.s3_bucket)


def presigned_get_url(*, key: str, expires_seconds: int = 3600) -> str:
    st = get_storage()
    return st.client.generate_presigned_url(
        "get_object",
        Params={"Bucket": st.bucket, "Key": key},
        ExpiresIn=expires_seconds,
    )

