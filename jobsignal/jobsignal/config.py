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
    # anthropic — облако Anthropic; openai — любой OpenAI-совместимый (YandexGPT,
    # DeepSeek, OpenAI…); gigachat — Сбер (рубли, из РФ); ollama — локально
    llm_provider: str = Field(default="anthropic")
    ollama_base_url: str = Field(default="http://localhost:11434")
    ollama_model: str = Field(default="qwen2.5:3b")

    # --- GigaChat (Сбер) ---
    gigachat_auth_key: str = Field(default="")  # «Авторизационные данные» (base64) из кабинета
    gigachat_scope: str = Field(default="GIGACHAT_API_PERS")  # PERS — физлица
    gigachat_model: str = Field(default="GigaChat")  # GigaChat | GigaChat-Pro | GigaChat-Max
    gigachat_verify_ssl: bool = Field(default=False)  # сертификаты НУЦ Минцифры не в trust store
    # Лёгкая модель GigaChat для разбора постов; пусто — та же, что GIGACHAT_MODEL.
    gigachat_model_small: str = Field(default="")

    # --- OpenAI-совместимый провайдер (LLM_PROVIDER=openai) ---
    # Через этот интерфейс работают YandexGPT (Yandex AI Studio), DeepSeek,
    # OpenAI и большинство других. Примеры — в config/.env.example.
    openai_base_url: str = Field(default="")
    openai_api_key: str = Field(default="")
    openai_model: str = Field(default="")         # сильная: оценка вакансий, письма
    openai_model_small: str = Field(default="")   # лёгкая: разбор постов; пусто — та же
    # Схема заголовка авторизации: Bearer (почти все) или Api-Key (Yandex Cloud).
    openai_auth_scheme: str = Field(default="Bearer")
    # Дополнительные заголовки JSON-объектом, например {"x-folder-id": "b1g..."}.
    openai_extra_headers: str = Field(default="")
    # response_format=json_object для ответов-JSON; не все провайдеры его знают.
    openai_json_mode: bool = Field(default=False)

    # --- DeepSeek (LLM_PROVIDER=deepseek или маршрут задачи) ---
    # Отдельное подключение рядом с OPENAI_*: так Яндекс и DeepSeek работают
    # одновременно, каждый на своих задачах (LLM_ROUTE_*).
    deepseek_api_key: str = Field(default="")
    deepseek_base_url: str = Field(default="https://api.deepseek.com")
    deepseek_model: str = Field(default="deepseek-chat")
    deepseek_model_small: str = Field(default="")

    # Маршруты задач: какой провайдер делает разбор постов, оценку вакансий,
    # письма и проверку фактов. Пусто — LLM_PROVIDER.
    llm_route_parse: str = Field(default="")
    llm_route_match: str = Field(default="")
    llm_route_letter: str = Field(default="")
    llm_route_facts: str = Field(default="")

    # Запасной провайдер: если основной закрыт (ключ, оплата, верификация),
    # запросы идут сюда. 03.11 Anthropic закрыл доступ до верификации
    # организации, и конвейер встал целиком. Пусто — без запасного.
    llm_fallback_provider: str = Field(default="")

    # --- Поведение конвейера ---
    match_threshold: int = Field(default=70)  # минимальный % соответствия для показа
    outreach_mode: OutreachMode = Field(default=OutreachMode.SEMI_AUTO)
    # Порог для автоматических отправок по таймеру (run.py hh-apply --auto).
    # Выше match_threshold сознательно: то, что между порогами, видно в
    # дашборде и уходит только с решения человека. Ручной прогон и кнопка в
    # дашборде этого порога не знают — они и есть присмотр.
    hh_auto_threshold: int = Field(default=80)

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
        """Профили с текстом для оценки: [{name, key, cv_text, resume_path}, ...].

        Источник — реестр jobsignal/profiles.py (config/profiles.yaml). Текст
        для оценки — выжимка cv_path, а если её нет — текст PDF резюме.
        Профиль без текста пропускаем громко: оценка по пустому резюме дала
        бы нули по всем вакансиям.
        """
        import logging

        from jobsignal import profiles as registry

        log = logging.getLogger("jobsignal")
        out: list[dict] = []
        for p in registry.all_profiles():
            text = p.cv_text()
            if not text:
                log.warning("[profiles] у профиля «%s» нет текста резюме: ни %s, ни "
                            "PDF (config/resumes/%s.pdf) — пропускаю",
                            p.name, p.cv_path or "cv_path", p.key)
                continue
            pdf = p.resume_pdf()
            out.append({"name": p.name, "key": p.key, "cv_text": text,
                        "resume_path": str(pdf) if pdf else None})
        return out


@lru_cache(maxsize=1)
def get_config() -> AppConfig:
    return AppConfig()
