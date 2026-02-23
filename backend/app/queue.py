from __future__ import annotations

import json
from typing import Any, Dict

import redis

from .settings import settings


QUEUE_NAME = "visa_jobs"


def get_redis() -> redis.Redis:
    return redis.Redis.from_url(settings.redis_url, decode_responses=True)


def enqueue_job(*, job_id: str, package_id: str) -> None:
    r = get_redis()
    msg = {"job_id": job_id, "package_id": package_id}
    r.rpush(QUEUE_NAME, json.dumps(msg, ensure_ascii=False))

