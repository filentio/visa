"""Единый лимитер откликов — общий для дашборда и агентов отправки.

Лимиты живут в окружении (config/.env), значения по умолчанию совпадают с тем,
что дашборд показывал до вынесения кода сюда:

    OUTREACH_PER_HOUR  (по умолчанию 3)
    OUTREACH_PER_DAY   (по умолчанию 15)

Считаются ВСЕ отправленные отклики — строки applications с непустым sent_at,
независимо от канала. hh, tg и остальные каналы делят одну квоту: лимит
защищает аккаунты и репутацию в целом, а не каждый канал по отдельности.

Единственная точка правды: и плашка «за час / за сутки» в дашборде,
и HHApplyAgent спрашивают квоту здесь.
"""
from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone

from sqlalchemy import func

from .db import Application, get_session_factory

DEFAULT_PER_HOUR = 3
DEFAULT_PER_DAY = 15

_sf = None


def _session():
    global _sf
    if _sf is None:
        _sf = get_session_factory()
    return _sf()


def per_hour() -> int:
    return int(os.environ.get("OUTREACH_PER_HOUR", str(DEFAULT_PER_HOUR)))


def per_day() -> int:
    return int(os.environ.get("OUTREACH_PER_DAY", str(DEFAULT_PER_DAY)))


def rate_status(session=None) -> dict:
    """Текущее состояние квоты.

    session — необязателен: если не передан, открываем свою и закрываем.
    Возвращает те же ключи, что раньше собирал дашборд, плюс next_slot_at.
    """
    own = session is None
    session = session or _session()
    try:
        now = datetime.now(timezone.utc)
        hour_ago = now - timedelta(hours=1)
        day_ago = now - timedelta(days=1)

        sent_hour = (
            session.query(func.count(Application.id))
            .filter(Application.sent_at >= hour_ago)
            .scalar()
        ) or 0
        sent_day = (
            session.query(func.count(Application.id))
            .filter(Application.sent_at >= day_ago)
            .scalar()
        ) or 0

        ph, pd = per_hour(), per_day()
        allowed = max(0, min(ph - sent_hour, pd - sent_day))

        # Когда освободится слот: самый старый отклик внутри исчерпанного окна
        # плюс длина окна. Часовое окно освобождается раньше суточного.
        next_slot_at = None
        if allowed == 0:
            window = hour_ago if sent_hour >= ph else day_ago
            length = timedelta(hours=1) if sent_hour >= ph else timedelta(days=1)
            oldest = (
                session.query(func.min(Application.sent_at))
                .filter(Application.sent_at >= window)
                .scalar()
            )
            if oldest is not None:
                if oldest.tzinfo is None:
                    oldest = oldest.replace(tzinfo=timezone.utc)
                next_slot_at = oldest + length

        return {
            "sent_hour": sent_hour,
            "sent_day": sent_day,
            "per_hour": ph,
            "per_day": pd,
            "allowed_now": allowed,
            "next_slot": next_slot_at.strftime("%H:%M") if next_slot_at else "",
            "next_slot_at": next_slot_at.isoformat() if next_slot_at else None,
        }
    finally:
        if own:
            session.close()
