"""Конфигурация проекта.

Секреты (ключи, телефон, сессия) читаются из .env.
Настраиваемые параметры и список каналов — из config/channels.yaml.
"""
from __future__ import annotations

import enum
from functools import lru_cache
from pathlib import Path

import yaml
from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)


class OutreachMode(str, enum.Enum):
    """Режим отправки сообщений рекрутёрам.

    SEMI_AUTO — система готовит черновик, отправка только по кнопке в дашборде (дефолт).
    FULL_AUTO — отправка без подтверждения (включать осознанно, после обкатки).
    """

    SEMI_AUTO = "semi_auto"
    FULL_AUTO = "full_auto"


class Settings(BaseSettings):
    """Секреты и подключения. Берутся из .env (см. config/.env.example)."""

    model_config = SettingsConfigDict(
        env_file=str(CONFIG_DIR / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # --- Telegram (user-аккаунт-сборщик, MTProto через Telethon) ---
    tg_api_id: int = Field(default=0)
    tg_api_hash: str = Field(default="")
    tg_phone: str = Field(default="")
    tg_session: str = Field(default="collector")  # имя файла сессии Telethon

    # --- Telegram-бот для уведомлений ЛИЧНО тебе (Bot API, отдельная сущность) ---
    notify_bot_token: str = Field(default="")
    notify_chat_id: str = Field(default="")

    # --- Anthropic ---
    anthropic_api_key: str = Field(default="")
    # модель по умолчанию (матчинг/генерация — где важнее качество)
    anthropic_model: str = Field(default="claude-sonnet-4-6")
    # парсер: дёшево и достаточно (классификация + извлечение полей)
    parser_model: str = Field(default="claude-haiku-4-5")
    # сколько постов парсить за один прогон (контроль расхода)
    parser_batch_limit: int = Field(default=200)
    # сколько вакансий оценивать матчером за один прогон
    match_batch_limit: int = Field(default=200)
    # сколько черновиков генерить за один батч (CLI compose)
    compose_batch_limit: int = Field(default=20)

    # --- Выбор LLM-провайдера ---
    # anthropic — облако Anthropic; ollama — локально; gigachat — Сбер (рубли, из РФ)
    llm_provider: str = Field(default="anthropic")
    ollama_base_url: str = Field(default="http://localhost:11434")
    ollama_model: str = Field(default="qwen2.5:3b")

    # --- GigaChat (Сбер) ---
    gigachat_auth_key: str = Field(default="")  # «Авторизационные данные» (base64) из кабинета
    gigachat_scope: str = Field(default="GIGACHAT_API_PERS")  # PERS — физлица
    gigachat_model: str = Field(default="GigaChat")  # GigaChat | GigaChat-Pro | GigaChat-Max
    gigachat_verify_ssl: bool = Field(default=False)  # сертификаты НУЦ Минцифры не в trust store

    # --- Поведение конвейера ---
    match_threshold: int = Field(default=70)  # минимальный % соответствия для показа
    outreach_mode: OutreachMode = Field(default=OutreachMode.SEMI_AUTO)

    # --- Хранилище ---
    database_url: str = Field(default=f"sqlite:///{DATA_DIR / 'jobsignal.db'}")

    # --- Дашборд (продакшн на сервере) ---
    dashboard_port: int = Field(default=5000)
    dashboard_user: str = Field(default="admin")
    dashboard_pass: str = Field(default="")  # пусто = без пароля (локально); на сервере задать!

    @field_validator("tg_api_id", mode="before")
    @classmethod
    def _empty_to_zero(cls, v):
        # пустой TG_API_ID в .env => 0 (на этапе 0 Telegram ещё не нужен)
        if v in ("", None):
            return 0
        return v


class AppConfig:
    """Объединяет секреты (Settings) и настройки из channels.yaml."""

    def __init__(self) -> None:
        self.settings = Settings()
        self._yaml = self._load_yaml()

    @staticmethod
    def _load_yaml() -> dict:
        path = CONFIG_DIR / "channels.yaml"
        if not path.exists():
            return {"channels": [], "cv_path": None}
        with path.open(encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    @property
    def channels(self) -> list[dict]:
        """Список каналов для мониторинга: [{handle, title, niche}, ...]."""
        return self._yaml.get("channels", [])

    @property
    def cv_path(self) -> str | None:
        """Путь к твоему резюме (txt/md/pdf) для скоринга."""
        return self._yaml.get("cv_path")

    @property
    def profiles(self) -> list[dict]:
        """Карьерные профили для мульти-таргет скоринга: [{name, cv_text}, ...].
        Читает config/profiles.yaml (name + cv_path), подгружает текст резюме."""
        import logging

        log = logging.getLogger("jobsignal")
        path = CONFIG_DIR / "profiles.yaml"
        if not path.exists():
            return []
        with path.open(encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        out: list[dict] = []
        for p in data.get("profiles", []):
            name = p.get("name")
            cv_path = p.get("cv_path")
            if not name or not cv_path:
                continue
            cv_file = Path(cv_path) if Path(cv_path).is_absolute() else (ROOT / cv_path)
            if not cv_file.exists():
                log.warning("[profiles] резюме не найдено: %s (профиль '%s') — пропускаю",
                            cv_file, name)
                continue
            text = cv_file.read_text(encoding="utf-8").strip()
            if not text:
                log.warning("[profiles] резюме пустое: %s — пропускаю", cv_file)
                continue
            # готовый PDF для отправки рекрутёру (опционально)
            resume_path = None
            rp = p.get("resume_path")
            if rp:
                rf = Path(rp) if Path(rp).is_absolute() else (ROOT / rp)
                if rf.exists():
                    resume_path = str(rf)
            out.append({"name": name, "cv_text": text, "resume_path": resume_path})
        return out


@lru_cache(maxsize=1)
def get_config() -> AppConfig:
    return AppConfig()
