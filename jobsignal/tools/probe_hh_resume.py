"""Разведка: подбор вакансий, который hh делает сам под резюме.

Зачем. Замер 21.09: вакансий с баллом от 80 в очереди 391, а отправить можно 7.
Ограничение системы — не оценка и не квота, а канал: автоотклик работает только
через hh, а hh даёт 294 вакансии из 1261 за неделю. Из трёх способов расширить
пул этот единственный не требует ни рисковать телеграм-аккаунтом, ни учиться
заполнять чужие формы: тот же канал, который уже работает целиком.

Что проверяем. Где лежит этот подбор и в какой форме. Если он отдаётся тем же
`vacancySearchResult.vacancies`, что и обычный поиск, то весь разбор карточек
(`_card_to_item`) переиспользуется как есть, и работы остаётся на полчаса.
Если формой отличается — нужен отдельный разбор.

Гадать не будем: сегодня три гипотезы подряд не подтвердились, и каждая
проверялась одним запросом.

Скрипт только читает. Запуск:

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_hh_resume.py
"""
from __future__ import annotations

import html as htmlmod
import json
import os
import re
import sys
from pathlib import Path

import requests

STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}
DEFAULT_STATE_PATH = "/root/autoapply/state.json"
RESUMES_URL = "https://hh.ru/applicant/resumes"

# Ключи карточки, которые нужны разбору из hh_collector._card_to_item. Если они
# есть — подбор читается существующим кодом без изменений.
CARD_KEYS = ("vacancyId", "name", "company", "area", "compensation")


def cookies() -> dict:
    path = Path(os.environ.get("HH_STATE_PATH", DEFAULT_STATE_PATH))
    data = json.loads(path.read_text())
    return {c["name"]: c["value"] for c in data.get("cookies", [])
            if "hh.ru" in (c.get("domain") or "")}


def state_of(url: str, jar: dict, params: dict | None = None):
    """(состояние, пояснение). Состояние None — страница не отдала данных."""
    try:
        r = requests.get(url, params=params or {}, headers=HEADERS,
                         cookies=jar, timeout=30)
    except requests.RequestException as exc:
        return None, f"запрос не прошёл: {exc}"
    if "account/login" in r.url or "auth" in r.url:
        return None, "перекинуло на вход — сессия протухла"
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    m = STATE_RE.search(r.text)
    if not m:
        return None, f"нет блока состояния ({len(r.text)} симв.)"
    try:
        return json.loads(htmlmod.unescape(m.group(1).strip())), f"HTTP 200, {r.url}"
    except json.JSONDecodeError as exc:
        return None, f"состояние не разбирается: {exc}"


def find_resumes(jar: dict) -> list[str]:
    """Идентификаторы резюме. Нужны, чтобы спросить подбор именно под них."""
    state, note = state_of(RESUMES_URL, jar)
    print(f"страница резюме: {note}")
    if not state:
        return []
    found: list[str] = []

    def walk(node):
        if isinstance(node, dict):
            for k, v in node.items():
                if k in ("_attributes", "resume") and isinstance(v, dict):
                    h = v.get("hash") or v.get("id")
                    if isinstance(h, str) and h not in found:
                        found.append(h)
                if k in ("resumeId", "hash") and isinstance(v, (str, int)):
                    sv = str(v)
                    if len(sv) > 5 and sv not in found:
                        found.append(sv)
                walk(v)
        elif isinstance(node, list):
            for x in node[:50]:
                walk(x)

    walk(state)
    return found[:5]


def report(label: str, state, note: str) -> int:
    """Печатает, нашёлся ли в состоянии список вакансий. Возвращает сколько."""
    print(f"\n=== {label}")
    print(f"    {note}")
    if not state:
        return 0
    res = state.get("vacancySearchResult")
    if isinstance(res, dict) and isinstance(res.get("vacancies"), list):
        cards = res["vacancies"]
        print(f"    vacancySearchResult.vacancies: {len(cards)} карточек "
              f"(всего найдено: {res.get('resultsFound')})")
        if cards:
            have = [k for k in CARD_KEYS if k in cards[0]]
            print(f"    ключи для _card_to_item: {len(have)} из {len(CARD_KEYS)} "
                  f"— {', '.join(have)}")
            missing = [k for k in CARD_KEYS if k not in cards[0]]
            if missing:
                print(f"    НЕТ: {', '.join(missing)}")
        return len(cards)
    # Не тот путь — покажем, что вообще похоже на вакансии.
    likely = [k for k in state
              if any(w in k.lower() for w in ("vacanc", "suitable", "similar",
                                              "recommend"))]
    print(f"    vacancySearchResult.vacancies нет; похожие ключи: {likely or '—'}")
    for k in likely:
        v = state[k]
        if isinstance(v, dict):
            print(f"      {k}: объект, ключи {sorted(v)[:8]}")
        elif isinstance(v, list):
            print(f"      {k}: список, {len(v)}")
    return 0


def main() -> int:
    jar = cookies()
    print(f"cookies для hh.ru: {len(jar)}\n")

    ids = find_resumes(jar)
    print(f"идентификаторы резюме: {ids or 'не нашлись'}")
    rid = ids[0] if ids else None

    # Кандидаты. Первый — обычный поиск с привязкой к резюме: если сработает,
    # переиспользуется весь существующий разбор. Остальные — отдельные
    # страницы рекомендаций.
    checks = [
        ("лента подходящих (главная соискателя)",
         "https://hh.ru/applicant/vacancy_feed", None),
        ("поиск с привязкой к резюме",
         "https://hh.ru/search/vacancy",
         {"resume": rid, "from": "resumelist", "items_on_page": 50} if rid else None),
        ("похожие вакансии под резюме",
         f"https://hh.ru/applicant/similar-vacancies/{rid}" if rid else None, None),
    ]
    best = 0
    for label, url, params in checks:
        if not url or (params is None and "search/vacancy" in (url or "")):
            print(f"\n=== {label}\n    пропущено: нет идентификатора резюме")
            continue
        state, note = state_of(url, jar, params)
        best = max(best, report(label, state, note))

    print("\n" + "-" * 60)
    if best:
        print(f"Подбор читается, максимум карточек за запрос: {best}.")
        print("Если ключи карточки на месте — hh_collector._card_to_item "
              "подходит как есть.")
    else:
        print("Ни один из адресов не отдал список вакансий. Нужен другой путь: "
              "скажи, и посмотрю сохранённое состояние целиком.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
