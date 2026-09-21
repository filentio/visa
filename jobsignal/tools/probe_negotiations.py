"""Разведка: как устроена страница «Мои отклики» на hh.ru.

Учёт ответов — нулевая задача в docs/ROADMAP.md: сейчас replied_at не
проставляет никто, поэтому «ноль ответов на сто откликов» означает не
«не отвечают», а «не смотрим». Прежде чем писать разбор, надо узнать
фактическую структуру страницы, а не предположить её: в июне сбор с hh
встал ровно потому, что структуру угадали один раз и больше не проверяли.

Скрипт ничего не меняет — только печатает форму данных. Запуск:

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_negotiations.py

Страница откликов доступна только авторизованному, поэтому нужны cookies из
state.json. Playwright при этом не нужен: сборщик с hh.ru ходит обычным
requests, и здесь достаточно того же плюс cookies (Chromium в пике съедает
660 МБ при 900 свободных — дёргать его ради чтения списка незачем).
"""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

import requests

NEGOTIATIONS_URL = "https://hh.ru/applicant/negotiations"
STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}
DEFAULT_STATE_PATH = "/root/autoapply/state.json"


def load_cookies(path: Path) -> dict:
    """Cookies из storage_state Playwright — в вид, понятный requests."""
    data = json.loads(path.read_text())
    jar = {}
    for c in data.get("cookies", []):
        if "hh.ru" in (c.get("domain") or ""):
            jar[c["name"]] = c["value"]
    return jar


def shape(value, depth=0, path="") -> list[str]:
    """Печатает форму данных, а не сами данные: ключи, типы, размеры.

    Значения не выводим намеренно — в отзывах лежат имена рекрутёров и тексты
    писем, а вывод пойдёт в переписку.
    """
    out = []
    pad = "  " * depth
    if isinstance(value, dict):
        out.append(f"{pad}{path or '{}'}: объект, {len(value)} ключей")
        if depth < 3:
            for k, v in list(value.items())[:25]:
                out.extend(shape(v, depth + 1, k))
    elif isinstance(value, list):
        out.append(f"{pad}{path}: список, {len(value)} элементов")
        if value and depth < 3:
            out.extend(shape(value[0], depth + 1, f"{path}[0]"))
    else:
        t = type(value).__name__
        hint = ""
        if isinstance(value, str) and len(value) < 40:
            hint = f" = {value!r}"      # короткие строки это чаще всего коды статусов
        out.append(f"{pad}{path}: {t}{hint}")
    return out


def main() -> int:
    state = Path(os.environ.get("HH_STATE_PATH", DEFAULT_STATE_PATH))
    if not state.exists():
        print(f"нет файла сессии: {state}")
        return 1
    cookies = load_cookies(state)
    print(f"cookies для hh.ru: {len(cookies)}")

    r = requests.get(NEGOTIATIONS_URL, headers=HEADERS, cookies=cookies,
                     timeout=30, allow_redirects=True)
    print(f"HTTP {r.status_code}, итоговый адрес: {r.url}")
    if "account/login" in r.url or "auth" in r.url:
        print("НЕ АВТОРИЗОВАНЫ: сессия протухла, state.json надо обновить")
        return 2

    m = STATE_RE.search(r.text)
    if not m:
        print("блок HH-Lux-InitialState не найден — структура страницы другая")
        Path("/tmp/negotiations.html").write_text(r.text)
        print("HTML сохранён в /tmp/negotiations.html, размер:", len(r.text))
        return 3

    data = json.loads(m.group(1))
    print(f"\nключей в состоянии: {len(data)}")
    # Ищем, где лежит собственно список откликов: имя ключа заранее неизвестно.
    cands = [k for k in data
             if any(w in k.lower() for w in ("negotiation", "topic", "response"))]
    print("ключи, похожие на отклики:", cands or "не нашлось")
    print("\nвсе ключи верхнего уровня:")
    for k in sorted(data):
        v = data[k]
        size = len(v) if isinstance(v, (list, dict)) else ""
        print(f"  {k}: {type(v).__name__} {size}")

    for k in cands:
        print(f"\n===== форма {k} =====")
        print("\n".join(shape(data[k], path=k)))

    # Полное состояние — на диск, чтобы можно было доразобраться без повторного
    # обращения к сайту. В нём персональные данные, так что в git ему нельзя.
    out = Path("/tmp/negotiations_state.json")
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1))
    print(f"\nполное состояние: {out} ({out.stat().st_size // 1024} КБ)")
    print("в нём персональные данные — в репозиторий не кладём")
    return 0


if __name__ == "__main__":
    sys.exit(main())
