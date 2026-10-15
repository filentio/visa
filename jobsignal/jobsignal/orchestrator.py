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
        from .llm import LLMUnavailable

        results = []
        for agent in self.agents:
            try:
                stats = agent.run()
            except LLMUnavailable:
                # Отказ провайдера — не сбой одного агента, а остановка всей
                # системы, и дальше по цепочке будет то же самое. Наверх, в
                # run.py: там причина пишется человеческим языком и прогон
                # выходит с кодом 1, чтобы сработал OnFailure. Раньше общий
                # except ниже гасил и это: конвейер выходил с кодом 0, тревога
                # не уходила — ровно простой 14.09, от которого он защищал.
                raise
            except Exception as exc:  # noqa: BLE001 — один агент не роняет остальных
                log.exception("[%s] ошибка: %s", agent.name, exc)
                stats = {"agent": agent.name, "status": "error", "error": str(exc)}
            results.append(stats)
        log.info("=== Конвейер завершён ===")
        return results

    @staticmethod
    def failed(results: list[dict]) -> list[str]:
        """Агенты, упавшие в прогоне. Прогон доводится до конца, но падение
        не должно выглядеть успехом: run.py по этому списку выходит с кодом 1."""
        return [r.get("agent", "?") for r in results
                if isinstance(r, dict) and r.get("status") == "error"]
