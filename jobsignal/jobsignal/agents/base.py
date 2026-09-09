"""Базовый класс агента конвейера.

Каждый агент — отдельная ответственность. На Этапе 0 все агенты — заглушки:
метод run() логирует, что он вызван, и возвращает пустую статистику.
Реальную логику добавляем поэтапно (Этапы 1–6).
"""
from __future__ import annotations

import logging
from abc import ABC, abstractmethod

from ..config import AppConfig

log = logging.getLogger("jobsignal")


class BaseAgent(ABC):
    name: str = "base"

    def __init__(self, config: AppConfig) -> None:
        self.config = config

    @abstractmethod
    def run(self) -> dict:
        """Выполняет шаг конвейера. Возвращает статистику, напр. {'processed': 10}."""
        ...

    def _stub(self) -> dict:
        log.info("[%s] заглушка — логика будет добавлена на следующих этапах", self.name)
        return {"agent": self.name, "status": "stub"}
