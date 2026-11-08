"""Обёртка над LLM для всех агентов (parser/matcher/composer).

Провайдер выбирается в config/.env через LLM_PROVIDER:
  - anthropic — облако Anthropic (ключ + оплата)
  - openai    — любой OpenAI-совместимый API: YandexGPT, DeepSeek, OpenAI…
  - gigachat  — Сбер GigaChat (из РФ, рубли, есть бесплатный лимит)
  - ollama    — локальная модель (бесплатно, без сети)

LLM_FALLBACK_PROVIDER — запасной: если основной закрыт целиком (ключ,
оплата, верификация организации), запросы уходят туда. У каждого провайдера
две модели: лёгкая для разбора постов (то, что агенты называют
parser_model) и сильная для оценки вакансий и писем.

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


class LLMUnavailable(LLMError):
    """Провайдер недоступен целиком: кончились деньги, ключ отозван, доступ закрыт.

    Отличается от LLMError тем, что ретраить бессмысленно — сломан не запрос, а
    доступ. Агенты обязаны на ней прерывать прогон, а не списывать попытку:
    14.09 баланс кончился на 30 часов, каждый часовой прогон исправно считал
    отказ сбоем разбора, и 255 постов выбили лимит попыток и выпали из очереди
    навсегда. Пост был цел — сломан был провайдер.
    """


# Признаки отказа на уровне доступа, а не запроса. Сверяем по типу исключения и
# по тексту: SDK для «кончились деньги» отдаёт обычный BadRequestError (400), и
# отличить его от настоящей ошибки в запросе можно только по сообщению.
_FATAL_TYPES = ("AuthenticationError", "PermissionDeniedError")
_FATAL_TEXTS = ("credit balance", "billing", "insufficient_quota",
                "authentication_error", "permission_error", "invalid x-api-key")


def _is_unavailable(exc: Exception) -> bool:
    if type(exc).__name__ in _FATAL_TYPES:
        return True
    low = str(exc).lower()
    return any(t in low for t in _FATAL_TEXTS)


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
        if _is_unavailable(exc):
            raise LLMUnavailable(str(exc)) from exc
        raise LLMError(str(exc)) from exc
    _log_usage(model, resp.usage, tag)
    u = resp.usage
    _record_usage("anthropic", sum((getattr(u, n, 0) or 0) for n in (
        "input_tokens", "cache_creation_input_tokens", "cache_read_input_tokens",
        "output_tokens")))
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


def _tier(model: str) -> str:
    """small — разбор постов (агенты передают parser_model), big — остальное.

    Агенты пишут имена моделей Anthropic; у других провайдеров свои модели,
    поэтому имя переводится в «уровень», а уровень — в модель провайдера.
    """
    return "small" if model and model == get_config().settings.parser_model else "big"


# --- OpenAI-совместимый (YandexGPT, DeepSeek, OpenAI…) ---

def _compat_conf(provider: str) -> dict:
    """Настройки OpenAI-совместимого подключения: openai (общее) или deepseek."""
    cfg = get_config().settings
    if provider == "deepseek":
        return {"name": "deepseek", "env": "DEEPSEEK", "base_url": cfg.deepseek_base_url,
                "api_key": cfg.deepseek_api_key, "model": cfg.deepseek_model,
                "model_small": cfg.deepseek_model_small, "auth": "Bearer",
                "headers": "", "json_mode": False}
    return {"name": "openai", "env": "OPENAI", "base_url": cfg.openai_base_url,
            "api_key": cfg.openai_api_key, "model": cfg.openai_model,
            "model_small": cfg.openai_model_small, "auth": cfg.openai_auth_scheme,
            "headers": cfg.openai_extra_headers, "json_mode": cfg.openai_json_mode}


def _openai_chat(system, user, tier, max_tokens, as_json, tag="", provider="openai"):
    c = _compat_conf(provider)
    if not (c["base_url"] and c["api_key"] and c["model"]):
        raise LLMUnavailable(f"{c['env']}_BASE_URL / {c['env']}_API_KEY / {c['env']}_MODEL "
                             f"не заданы (кабинет → «Настройки»)")
    model = (c["model_small"] if tier == "small" else "") or c["model"]
    headers = {"Authorization": f"{c['auth']} {c['api_key']}",
               "Content-Type": "application/json"}
    if (c["headers"] or "").strip():
        try:
            headers.update(json.loads(c["headers"]))
        except ValueError as exc:
            raise LLMUnavailable(f"{c['env']}_EXTRA_HEADERS не JSON: {exc}") from exc
    payload = {
        "model": model,
        "messages": [{"role": "system", "content": system},
                     {"role": "user", "content": user}],
        "temperature": 0.2,
        "max_tokens": max_tokens,
    }
    if as_json and c["json_mode"]:
        payload["response_format"] = {"type": "json_object"}
    url = c["base_url"].rstrip("/") + "/chat/completions"
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=180)
    except requests.RequestException as exc:
        raise LLMError(f"{url} недоступен: {exc}") from exc
    if r.status_code != 200:
        text = r.text[:300]
        # 401/403 — ключ, 402 — оплата; insufficient_quota OpenAI отдаёт с 429.
        fatal = r.status_code in (401, 402, 403) or _is_unavailable(Exception(text))
        raise (LLMUnavailable if fatal else LLMError)(f"HTTP {r.status_code}: {text}")
    data = r.json()
    usage = data.get("usage") or {}
    # DeepSeek сообщает, сколько входа пришло из кэша: по этой строке видно,
    # работает ли скидка на повторяющиеся резюме.
    hit = usage.get("prompt_cache_hit_tokens")
    log.info("[usage]%s model=%s in=%s out=%s%s", f" {tag}" if tag else "", model,
             usage.get("prompt_tokens", "?"), usage.get("completion_tokens", "?"),
             f" cache_hit={hit}" if hit is not None else "")
    _record_usage(c["name"], int(usage.get("total_tokens") or 0) or
                  int(usage.get("prompt_tokens") or 0) + int(usage.get("completion_tokens") or 0))
    try:
        return data["choices"][0]["message"]["content"] or ""
    except (KeyError, IndexError, TypeError) as exc:
        raise LLMError(f"неожиданный ответ: {str(data)[:200]}") from exc


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
        # 401/403 на OAuth — ключ не тот или отозван: ретрай не поможет.
        cls = LLMUnavailable if r.status_code in (401, 403) else LLMError
        raise cls(f"GigaChat OAuth HTTP {r.status_code}: {r.text[:200]}")
    data = r.json()
    exp = data.get("expires_at", 0) or 0
    if exp > 1e12:        # иногда приходит в миллисекундах
        exp /= 1000.0
    _giga_token["value"] = data["access_token"]
    _giga_token["exp"] = exp or (now + 1500)
    return _giga_token["value"]


def _gigachat_chat(system, user, max_tokens, tier="big"):
    cfg = get_config().settings
    token = _gigachat_token()
    model = (cfg.gigachat_model_small if tier == "small" else "") or cfg.gigachat_model
    try:
        r = requests.post(
            GIGA_CHAT,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "model": model,
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
        # 402 — исчерпан оплаченный лимит, 401/403 — доступ закрыт.
        cls = LLMUnavailable if r.status_code in (401, 402, 403) else LLMError
        raise cls(f"GigaChat HTTP {r.status_code}: {r.text[:200]}")
    data = r.json()
    _record_usage("gigachat", int((data.get("usage") or {}).get("total_tokens") or 0))
    return data["choices"][0]["message"]["content"]


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


def _iter_top_level_objects(text: str):
    """Куски «{...}» верхнего уровня, с учётом строк и экранирования.

    Нужен, когда массив обрезан по лимиту токенов и целиком не парсится:
    берём объекты, которые успели закрыться, а незавершённый хвост
    отбрасываем — достраивать его значило бы выдумывать поля.
    """
    depth = 0
    start = None
    in_str = False
    esc = False
    for i, ch in enumerate(text):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}" and depth:
            depth -= 1
            if depth == 0 and start is not None:
                yield text[start:i + 1]
                start = None


def _loads_lenient_many(text: str) -> list[dict]:
    """Ответ модели как СПИСОК объектов: и один объект, и массив.

    Парсер просит один объект, но на посте с несколькими вакансиями модель
    возвращает массив. _loads_lenient на нём падает: он режет от первой «{»
    до последней «}» и получает «{...},{...}» — невалидный JSON. Такие посты
    уходили в skipped и помечались разобранными, то есть терялись насовсем.
    """
    t = _strip_fences(text).strip()
    if not t.startswith("["):
        return [_loads_lenient(t)]

    try:
        data = json.loads(t)
    except json.JSONDecodeError:
        pass
    else:
        if isinstance(data, list):
            return [d for d in data if isinstance(d, dict)]

    objs: list[dict] = []
    for block in _iter_top_level_objects(t):
        for candidate in (block, _repair_json(block)):
            try:
                obj = json.loads(candidate)
            except json.JSONDecodeError:
                continue
            if isinstance(obj, dict):
                objs.append(obj)
            break
    if objs:
        return objs
    raise json.JSONDecodeError("массив без пригодных объектов", t, 0)


PROVIDERS = ("anthropic", "openai", "deepseek", "gigachat", "ollama")
TASKS = ("parse", "match", "letter", "facts")

# ── Дневной лимит токенов ────────────────────────────────────────────────────
# 03.11 YandexGPT за два часа обошёлся в ~$20: переоценка прогнала 200
# вакансий, у провайдера нет кэша промпта, и каждое резюме оплачивалось
# заново. LLM_DAILY_TOKENS — потолок на сутки (по UTC) по всем провайдерам;
# дошли — модель «закрыта» до завтра (LLMUnavailable: агенты прерывают прогон,
# попытки постам не списываются), в бот уходит одно сообщение. 0 — без лимита.
USAGE_FILE = "data/llm_usage.json"


def _usage_path():
    import os
    from pathlib import Path
    return Path(os.environ.get("LLM_USAGE_FILE", USAGE_FILE))


# Лимит по умолчанию: без него перерасход замечали только по счёту
# провайдера (Яндекс 03.11 и 04-06.11). 5 млн токенов — с запасом на обычные
# сутки (разбор, оценка, письма) и на переоценку. 0 — без лимита, явно.
DEFAULT_DAILY_TOKENS = 5_000_000


def _daily_limit() -> int:
    import os
    raw = os.environ.get("LLM_DAILY_TOKENS")
    if raw is None or not raw.strip():
        return DEFAULT_DAILY_TOKENS
    try:
        return int(float(raw))
    except ValueError:
        return DEFAULT_DAILY_TOKENS


def usage_today() -> dict:
    """{'total': N, 'by_provider': {...}, 'limit': L} за сегодня (UTC)."""
    from datetime import datetime, timezone
    try:
        data = json.loads(_usage_path().read_text())
    except (OSError, ValueError):
        data = {}
    today = datetime.now(timezone.utc).date().isoformat()
    day = data.get("days", {}).get(today, {})
    return {"total": sum(day.values()), "by_provider": day, "limit": _daily_limit()}


def _record_usage(provider: str, tokens: int) -> None:
    import fcntl
    from datetime import datetime, timedelta, timezone
    if not tokens or tokens <= 0:
        return
    path = _usage_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path.with_suffix(".lock"), "w") as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            try:
                data = json.loads(path.read_text())
            except (OSError, ValueError):
                data = {}
            days = data.setdefault("days", {})
            now = datetime.now(timezone.utc).date()
            day = days.setdefault(now.isoformat(), {})
            day[provider] = day.get(provider, 0) + int(tokens)
            cutoff = (now - timedelta(days=31)).isoformat()
            data["days"] = {d: v for d, v in days.items() if d >= cutoff}
            path.write_text(json.dumps(data))
    except OSError as exc:
        log.warning("[llm] учёт токенов не записан: %s", exc)


def _check_budget() -> None:
    limit = _daily_limit()
    if limit <= 0:
        return
    used = usage_today()["total"]
    if used < limit:
        return
    _notify_budget(used, limit)
    raise LLMUnavailable(f"дневной лимит токенов исчерпан: {used} из {limit} "
                         f"(LLM_DAILY_TOKENS) — продолжим завтра")


def _notify_budget(used: int, limit: int) -> None:
    """Одно сообщение в бот за сутки о том, что лимит исчерпан."""
    from datetime import datetime, timezone
    flag = _usage_path().with_suffix(".notified")
    today = datetime.now(timezone.utc).date().isoformat()
    try:
        if flag.exists() and flag.read_text().strip() == today:
            return
        flag.parent.mkdir(parents=True, exist_ok=True)
        flag.write_text(today)
    except OSError:
        return
    try:
        from jobsignal.agents.notify_bot import _send
        _send(f"⏸ <b>Дневной лимит модели исчерпан</b>: {used} из {limit} токенов.\n"
              "Разбор, оценка и письма продолжатся завтра. Лимит меняется в кабинете: "
              "«Настройки» → «Экономия».")
    except Exception as exc:  # noqa: BLE001
        log.warning("[llm] сообщение о лимите не отправлено: %s", exc)

# Основной провайдер закрыт целиком — до этого времени сразу идём в запасной,
# не тратя на каждый запрос заведомо отказной вызов.
_primary_down_until = 0.0
PRIMARY_RECHECK_S = 15 * 60


def _call(provider, system, user, model, max_tokens, as_json, tag="", cache_system=False):
    tier = _tier(model)
    if provider == "ollama":
        return _ollama_chat(system, user, max_tokens, as_json=as_json)
    if provider == "gigachat":
        return _gigachat_chat(system, user, max_tokens, tier=tier)
    if provider in ("openai", "deepseek"):
        return _openai_chat(system, user, tier, max_tokens, as_json, tag=tag,
                            provider=provider)
    if provider == "anthropic":
        return _anthropic_text(system, user, model, max_tokens, tag=tag,
                               cache_system=cache_system)
    raise LLMUnavailable(f"неизвестный LLM_PROVIDER={provider!r}: нужен один из "
                         + ", ".join(PROVIDERS))


# Провайдер закрыт целиком — до этого времени сразу идём в запасной.
_down_until: dict[str, float] = {}


def route(task: str | None) -> str:
    """Провайдер задачи: LLM_ROUTE_<TASK>, иначе LLM_PROVIDER."""
    cfg = get_config().settings
    chosen = getattr(cfg, f"llm_route_{task}", "") if task in TASKS else ""
    return (chosen or cfg.llm_provider or "anthropic").strip().lower()


def _raw(system, user, model, max_tokens, as_json, tag="", cache_system=False, task=None):
    global _primary_down_until
    cfg = get_config().settings
    primary = route(task)
    fallback = (cfg.llm_fallback_provider or "").strip().lower()
    if fallback == primary:
        fallback = ""
    if tag != "llm_check":
        _check_budget()
    if fallback and time.time() < _down_until.get(primary, 0):
        return _call(fallback, system, user, model, max_tokens, as_json, tag, cache_system)
    try:
        return _call(primary, system, user, model, max_tokens, as_json, tag, cache_system)
    except LLMUnavailable as exc:
        if not fallback:
            raise
        _down_until[primary] = time.time() + PRIMARY_RECHECK_S
        _primary_down_until = _down_until[primary]
        log.warning("[llm] провайдер %s закрыт (%s) — на %d мин переключаюсь "
                    "на запасной %s", primary, str(exc)[:160], PRIMARY_RECHECK_S // 60,
                    fallback)
        return _call(fallback, system, user, model, max_tokens, as_json, tag, cache_system)


def check(provider: str) -> dict:
    """Проверка провайдера: отвечают ли обе модели. Для run.py llm-check."""
    cfg = get_config().settings
    out = {}
    for tier, model in (("small", cfg.parser_model), ("big", cfg.anthropic_model)):
        try:
            text = _call(provider, "Отвечай одним словом.", "Скажи: ok", model, 20,
                         as_json=False, tag="llm_check")
            out[tier] = ("ok", (text or "").strip()[:40])
        except Exception as exc:  # noqa: BLE001
            kind = "закрыт" if isinstance(exc, LLMUnavailable) else "сбой"
            out[tier] = (kind, f"{type(exc).__name__}: {str(exc)[:200]}")
    return out


def complete_text(system: str, user: str, model: str, max_tokens: int = 1024,
                  tag: str = "", cache_system: bool = False, task: str | None = None) -> str:
    """task — parse | match | letter | facts: какой провайдер её делает (LLM_ROUTE_*)."""
    return _raw(system, user, model, max_tokens, as_json=False, tag=tag,
                cache_system=cache_system, task=task)


def complete_json(system: str, user: str, model: str, max_tokens: int = 1024,
                  tag: str = "", cache_system: bool = False, task: str | None = None) -> dict:
    """JSON-ответ. LLMError — транзиентный сбой (ретрай); ValueError — кривой JSON."""
    raw = _raw(system, user, model, max_tokens, as_json=True, tag=tag,
               cache_system=cache_system, task=task)
    try:
        return _loads_lenient(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"модель вернула невалидный JSON: {raw[:200]}") from exc
