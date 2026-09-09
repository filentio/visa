"""Оркестратор конвейера.

Гоняет агентов по порядку. На Этапе 0 каждый агент — заглушка, но сам конвейер,
логирование и порядок шагов уже рабочие. Запускается вручную или по расписанию.

Порядок: collector → raw_dedup → parser → deduplicator → matcher → composer → outreach.
ReplyMonitor живёт отдельным долгоживущим процессом (слушатель входящих), не в этом цикле.
"""
from __future__ import annotations

import logging

from .agents import (
    CollectorAgent,
    DeduplicatorAgent,
    MatcherAgent,
    ParserAgent,
    RawDeduplicatorAgent,
)
from .config import AppConfig, get_config

log = logging.getLogger("jobsignal")

# Периодический конвейер (по таймеру на сервере): только сбор и оценка.
# Composer/Outreach — semi-auto, запускаются из дашборда по кнопке, не по расписанию.
PIPELINE = [
    CollectorAgent,
    RawDeduplicatorAgent,   # грубый дедуп по хешу текста — до платного разбора
    ParserAgent,
    DeduplicatorAgent,
    MatcherAgent,
]


class Orchestrator:
    def __init__(self, config: AppConfig | None = None) -> None:
        self.config = config or get_config()
        self.agents = [cls(self.config) for cls in PIPELINE]

    def run_once(self) -> list[dict]:
        """Один полный проход конвейера."""
        log.info("=== Запуск конвейера (режим аутрича: %s) ===",
                 self.config.settings.outreach_mode.value)
        results = []
        for agent in self.agents:
            try:
                stats = agent.run()
            except Exception as exc:  # noqa: BLE001 — на этапе 0 логируем и идём дальше
                log.exception("[%s] ошибка: %s", agent.name, exc)
                stats = {"agent": agent.name, "status": "error", "error": str(exc)}
            results.append(stats)
        log.info("=== Конвейер завершён ===")
        return results
