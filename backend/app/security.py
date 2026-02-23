from __future__ import annotations

from fastapi import Header, HTTPException, status

from .settings import settings


def require_internal_key(x_internal_api_key: str | None = Header(default=None)) -> None:
    if not x_internal_api_key or x_internal_api_key != settings.internal_api_key:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Unauthorized")

