"""Обёртка над LLM для всех агентов (parser/matcher/composer).

Провайдер выбирается в config/.env через LLM_PROVIDER:
  - anthropic — облако Anthropic (ключ + оплата)
  - ollama    — локальная модель на Mac (бесплатно, без сети)
  - gigachat  — Сбер GigaChat (из РФ, рубли, есть бесплатный лимит)

Логика агентов от провайдера не зависит — переключение одной строкой в .env.
"""
from __future__ import annotations

import json
import logging
import re
import time
import uuid

import requests

from .config import get_config

log = logging.getLogger("jobsignal")

_client = None
_giga_token = {"value": None, "exp": 0.0}


class LLMError(Exception):
    """Транзиентная ошибка (сеть/лимит/сервер/токен) — стоит ретраить."""


# --- Anthropic ---

def _get_client():
    global _client
    if _client is None:
        try:
            from anthropic import Anthropic
        except ImportError as exc:  # pragma: no cover
            raise LLMError("пакет anthropic не установлен: pip install anthropic") from exc
        key = get_config().settings.anthropic_api_key
        if not key:
            raise LLMError("ANTHROPIC_API_KEY пуст — впиши его в config/.env")
        _client = Anthropic(api_key=key)
    return _client


def _log_usage(model, usage, tag=""):
    """Пишет фактический расход токенов по ответу.

    Без этих строк экономию не с чем сравнить: input_tokens — вход по полной цене,
    cache_creation_input_tokens — запись в кэш (1.25x), cache_read_input_tokens —
    чтение из кэша (0.1x), output_tokens — выход. Поля могут прийти None.
    """
    def _n(name):
        return getattr(usage, name, 0) or 0

    log.info("[usage]%s model=%s in=%d cache_w=%d cache_r=%d out=%d",
             f" {tag}" if tag else "", model,
             _n("input_tokens"), _n("cache_creation_input_tokens"),
             _n("cache_read_input_tokens"), _n("output_tokens"))


def _anthropic_text(system, user, model, max_tokens, tag="", cache_system=False):
    """cache_system=True — ставит точку кэширования на системный промпт.

    Порядок сборки промпта — tools -> system -> messages, поэтому кэшируется
    именно префикс: система должна быть байт-в-байт одинаковой внутри прогона,
    а меняющийся текст вакансии идёт в messages, уже после точки кэширования.
    Минимальный кэшируемый префикс у claude-sonnet-5 — 1024 токена; короче
    кэш не создаётся молча, без ошибки. TTL — 5 минут от начала запроса.
    """
    system_param = system
    if cache_system:
        system_param = [{
            "type": "text",
            "text": system,
            "cache_control": {"type": "ephemeral"},
        }]
    try:
        resp = _get_client().messages.create(
            model=model, max_tokens=max_tokens, system=system_param,
            messages=[{"role": "user", "content": user}],
        )
    except Exception as exc:  # noqa: BLE001
        raise LLMError(str(exc)) from exc
    _log_usage(model, resp.usage, tag)
    return "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")


# --- Ollama (локально) ---

def _ollama_chat(system, user, max_tokens, as_json):
    cfg = get_config().settings
    url = cfg.ollama_base_url.rstrip("/") + "/api/chat"
    payload = {
        "model": cfg.ollama_model,
        "messages": [{"role": "system", "content": system},
                     {"role": "user", "content": user}],
        "stream": False,
        "options": {"temperature": 0, "num_predict": max_tokens},
    }
    if as_json:
        payload["format"] = "json"
    try:
        r = requests.post(url, json=payload, timeout=300)
    except requests.RequestException as exc:
        raise LLMError(f"Ollama недоступен ({exc}). `ollama serve` запущен? "
                       f"модель скачана: `ollama pull {cfg.ollama_model}`?") from exc
    if r.status_code != 200:
        raise LLMError(f"Ollama HTTP {r.status_code}: {r.text[:200]}")
    return r.json().get("message", {}).get("content", "")


# --- GigaChat (Сбер) ---

GIGA_OAUTH = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
GIGA_CHAT = "https://gigachat.devices.sberbank.ru/api/v1/chat/completions"


def _gigachat_token():
    cfg = get_config().settings
    now = time.time()
    if _giga_token["value"] and now < _giga_token["exp"] - 60:
        return _giga_token["value"]
    if not cfg.gigachat_auth_key:
        raise LLMError("GIGACHAT_AUTH_KEY пуст — впиши ключ авторизации в config/.env")
    try:
        r = requests.post(
            GIGA_OAUTH,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "RqUID": str(uuid.uuid4()),
                "Authorization": f"Basic {cfg.gigachat_auth_key}",
            },
            data={"scope": cfg.gigachat_scope},
            timeout=30,
            verify=cfg.gigachat_verify_ssl,
        )
    except requests.RequestException as exc:
        raise LLMError(f"GigaChat OAuth недоступен: {exc}") from exc
    if r.status_code != 200:
        raise LLMError(f"GigaChat OAuth HTTP {r.status_code}: {r.text[:200]}")
    data = r.json()
    exp = data.get("expires_at", 0) or 0
    if exp > 1e12:        # иногда приходит в миллисекундах
        exp /= 1000.0
    _giga_token["value"] = data["access_token"]
    _giga_token["exp"] = exp or (now + 1500)
    return _giga_token["value"]


def _gigachat_chat(system, user, max_tokens):
    cfg = get_config().settings
    token = _gigachat_token()
    try:
        r = requests.post(
            GIGA_CHAT,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "model": cfg.gigachat_model,
                "messages": [{"role": "system", "content": system},
                             {"role": "user", "content": user}],
                "temperature": 0.1,
                "max_tokens": max_tokens,
            },
            timeout=120,
            verify=cfg.gigachat_verify_ssl,
        )
    except requests.RequestException as exc:
        raise LLMError(f"GigaChat недоступен: {exc}") from exc
    if r.status_code == 401:
        _giga_token["value"] = None  # токен протух — сбросим, ретрай перелогинится
        raise LLMError("GigaChat 401 (токен истёк) — повтор")
    if r.status_code != 200:
        raise LLMError(f"GigaChat HTTP {r.status_code}: {r.text[:200]}")
    return r.json()["choices"][0]["message"]["content"]


# --- Общий интерфейс ---

def _strip_fences(text: str) -> str:
    t = text.strip()
    if t.startswith("```"):
        t = t.split("\n", 1)[1] if "\n" in t else t
        if t.endswith("```"):
            t = t[: t.rfind("```")]
    return t.strip()


def _repair_json(t: str) -> str:
    """Чинит типичные огрехи моделей: умные кавычки и висячие запятые."""
    # умные/типографские кавычки -> обычные
    for ch in ("\u201c", "\u201d", "\u00ab", "\u00bb"):
        t = t.replace(ch, '"')
    t = t.replace("\u2018", "'").replace("\u2019", "'")
    # висячие запятые перед } или ]
    t = re.sub(r",\s*([}\]])", r"\1", t)
    return t


def _close_truncated(t: str) -> str:
    """Достраивает оборванный по лимиту токенов JSON: закрывает строку и скобки.
    Спасает частичный ответ (например, успевшие прийти профили) вместо потери всего."""
    # отрезаем висящий хвост после последнего полного объекта/значения
    t = t.rstrip().rstrip(",")
    # баланс кавычек: если нечётно — закрываем строку
    if t.count('"') % 2 == 1:
        t += '"'
    # докрываем скобки в правильном порядке по стеку
    stack = []
    in_str = False
    esc = False
    for ch in t:
        if esc:
            esc = False
            continue
        if ch == "\\":
            esc = True
            continue
        if ch == '"':
            in_str = not in_str
            continue
        if in_str:
            continue
        if ch in "{[":
            stack.append(ch)
        elif ch == "}" and stack and stack[-1] == "{":
            stack.pop()
        elif ch == "]" and stack and stack[-1] == "[":
            stack.pop()
    for opener in reversed(stack):
        t += "}" if opener == "{" else "]"
    return t


def _loads_lenient(text: str) -> dict:
    """Парсит JSON, терпя мусор вокруг, типичные огрехи и обрезку по лимиту токенов."""
    t = _strip_fences(text)
    i = t.find("{")
    if i == -1:
        raise json.JSONDecodeError("нет JSON-объекта", t, 0)
    tail = t[i:]                      # от первой { до конца (для достройки обрезанного)
    j = tail.rfind("}")
    block = tail[: j + 1] if j > 0 else tail  # до последней } (для корректного с мусором)

    for candidate in (block, _repair_json(block), _repair_json(_close_truncated(tail))):
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue
    raise json.JSONDecodeError("не удалось распарсить", t, 0)


def _raw(system, user, model, max_tokens, as_json, tag="", cache_system=False):
    provider = get_config().settings.llm_provider
    if provider == "ollama":
        return _ollama_chat(system, user, max_tokens, as_json=as_json)
    if provider == "gigachat":
        return _gigachat_chat(system, user, max_tokens)
    return _anthropic_text(system, user, model, max_tokens, tag=tag,
                           cache_system=cache_system)


def complete_text(system: str, user: str, model: str, max_tokens: int = 1024,
                  tag: str = "", cache_system: bool = False) -> str:
    return _raw(system, user, model, max_tokens, as_json=False, tag=tag,
                cache_system=cache_system)


def complete_json(system: str, user: str, model: str, max_tokens: int = 1024,
                  tag: str = "", cache_system: bool = False) -> dict:
    """JSON-ответ. LLMError — транзиентный сбой (ретрай); ValueError — кривой JSON."""
    raw = _raw(system, user, model, max_tokens, as_json=True, tag=tag,
               cache_system=cache_system)
    try:
        return _loads_lenient(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"модель вернула невалидный JSON: {raw[:200]}") from exc
