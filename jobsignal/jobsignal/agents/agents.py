"""Агенты конвейера. Этап 0 — заглушки с описанием будущей логики.

Каждый класс реализуется на своём этапе (см. README, раздел «Дорожная карта»).
"""
from __future__ import annotations

from .base import BaseAgent


class OutreachAgent(BaseAgent):
    """Этап 5. SEMI_AUTO: ничего не шлёт сам, ждёт кнопки в дашборде.
    FULL_AUTO (опц.): отправляет черновики с rate-limit. Ставит sent_at, статус APPLIED."""

    name = "outreach"

    def run(self) -> dict:
        return self._stub()


class ReplyMonitorAgent(BaseAgent):
    """Этап 6. Слушает входящие в TG-аккаунте, матчит ответ к Application,
    создаёт Reply, шлёт уведомление тебе, статус REPLIED."""

    name = "reply_monitor"

    def run(self) -> dict:
        return self._stub()
