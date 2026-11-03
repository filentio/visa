"""Чтение и правка config/.env из кабинета.

Зачем. Ключи и провайдер модели менялись только на сервере через nano, а
03.11, когда Anthropic закрыл доступ, переключить конвейер на другую модель
можно было лишь по SSH. Теперь настройки модели правятся на странице
«Настройки».

Как пишем. Правим строки на месте: комментарии, порядок и остальные ключи
файла не трогаем; отсутствующие ключи дописываем в конец. Перед записью —
копия .env.bak-<время>, права 600. После записи обновляем os.environ этого
процесса и сбрасываем кэш настроек: процессы по таймеру и так читают файл
заново при каждом запуске, а долгоживущие (кабинет, бот) иначе работали бы
на старых значениях.

Секреты наружу не отдаются никогда: кабинет получает только «задан» и
последние четыре символа.
"""
from __future__ import annotations

import os
import re
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = ROOT / "config" / ".env"

_LINE_RE = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=(.*)$")


def _unquote(raw: str) -> str:
    v = raw.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
        inner = v[1:-1]
        return inner.replace('\\"', '"') if v[0] == '"' else inner
    # Как python-dotenv: « #» и дальше у значения без кавычек — комментарий.
    m = re.search(r"\s#", v)
    return (v[:m.start()] if m else v).strip()


def _quote(value: str) -> str:
    if value == "" or re.fullmatch(r"[A-Za-z0-9_./:@+=,\-]+", value):
        return value
    if "'" not in value:
        return f"'{value}'"
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def read(path: Path | None = None) -> dict[str, str]:
    path = path or ENV_PATH
    out: dict[str, str] = {}
    if not path.exists():
        return out
    for line in path.read_text(encoding="utf-8").splitlines():
        m = _LINE_RE.match(line)
        if m:
            out[m.group(1)] = _unquote(m.group(2))
    return out


def mask(value: str) -> dict:
    """Секрет для кабинета: задан ли и последние четыре символа."""
    value = value or ""
    return {"set": bool(value), "hint": ("…" + value[-4:]) if len(value) >= 8 else ""}


def update(changes: dict[str, str | None], path: Path | None = None) -> None:
    """Записать изменения. None — удалить ключ из файла."""
    path = path or ENV_PATH
    for k, v in changes.items():
        if not re.fullmatch(r"[A-Z_][A-Z0-9_]*", k):
            raise ValueError(f"недопустимое имя настройки: {k}")
        if v is not None and ("\n" in v or "\r" in v):
            raise ValueError(f"{k}: значение в одну строку")
    lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
    done: set[str] = set()
    new_lines: list[str] = []
    for line in lines:
        m = _LINE_RE.match(line)
        key = m.group(1) if m else None
        if key in changes:
            if key in done:
                continue                      # дубль ключа — оставляем одну строку
            done.add(key)
            if changes[key] is None:
                continue
            new_lines.append(f"{key}={_quote(changes[key])}")
        else:
            new_lines.append(line)
    tail = [f"{k}={_quote(v)}" for k, v in changes.items()
            if k not in done and v is not None]
    if tail:
        new_lines += ["", f"# изменено в кабинете {time.strftime('%d.%m.%Y %H:%M')}"] + tail
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        bak = path.with_name(f".env.bak-{time.strftime('%Y%m%d-%H%M%S')}")
        bak.write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
        os.chmod(bak, 0o600)
    tmp = path.with_name(".env.tmp")
    tmp.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
    os.chmod(tmp, 0o600)
    tmp.replace(path)
    _apply_here(changes)


def _apply_here(changes: dict[str, str | None]) -> None:
    """Новые значения — в этот процесс: os.environ, кэш настроек, клиенты LLM."""
    for k, v in changes.items():
        if v is None:
            os.environ.pop(k, None)
        else:
            os.environ[k] = v
    from jobsignal import config, llm
    config.get_config.cache_clear()
    llm._client = None
    llm._giga_token["value"] = None
    llm._primary_down_until = 0.0
    llm._down_until.clear()
