"""Разведка, шаг 2: чем связать отклик на hh.ru с нашей базой.

Первый шаг (tools/probe_negotiations.py) показал, что в состоянии страницы
лежит applicantNegotiations.topicList, а у каждой записи есть lastState —
готовый признак ответа (RESPONSE → INVITATION / INTERVIEW / DISCARD / HIRED).
Чего не хватает: идентификатора вакансии. У записи 52 ключа, первый разведчик
печатал 25 и до него не дошёл.

Скрипт читает уже сохранённый /tmp/negotiations_state.json — нового обращения
к hh не делает. Значения печатает только для коротких полей: в записях лежат
имена рекрутёров и названия компаний, а вывод идёт в переписку.

Запуск:

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_topic.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

STATE_FILE = Path("/tmp/negotiations_state.json")
# Ключи, значения которых печатать нельзя: имена, компании, тексты писем.
SECRET_HINTS = ("name", "company", "employer", "manager", "text", "letter",
                "title", "phone", "email", "url", "href", "logo")


def safe(key: str, value) -> str:
    """Тип и размер — всегда; значение — только если поле явно техническое."""
    t = type(value).__name__
    if isinstance(value, (dict, list)):
        return f"{t}[{len(value)}]"
    low = key.lower()
    if any(h in low for h in SECRET_HINTS):
        return f"{t} (значение скрыто)"
    if isinstance(value, str) and len(value) > 60:
        return f"{t}, {len(value)} симв."
    return f"{t} = {value!r}"


def main() -> int:
    if not STATE_FILE.exists():
        print(f"нет файла {STATE_FILE} — сначала прогони probe_negotiations.py")
        return 1
    data = json.loads(STATE_FILE.read_text())
    neg = data.get("applicantNegotiations") or {}
    topics = neg.get("topicList") or []
    print(f"откликов на странице: {len(topics)}, всего: {neg.get('total')}, "
          f"страниц: {neg.get('pageCount')}, фильтр: {neg.get('filterInUse')!r}")

    if not topics:
        print("список пуст — разбирать нечего")
        return 2

    t = topics[0]
    print(f"\n===== все {len(t)} ключей первой записи =====")
    for k in sorted(t):
        print(f"  {k}: {safe(k, t[k])}")

    # Главный вопрос: где идентификатор вакансии. Ищем по имени ключа на любой
    # глубине — он может лежать и во вложенном объекте.
    print("\n===== где может быть id вакансии =====")
    found = []

    def walk(node, path=""):
        if isinstance(node, dict):
            for k, v in node.items():
                p = f"{path}.{k}" if path else k
                if "vacanc" in k.lower():
                    found.append((p, safe(k, v)))
                walk(v, p)
        elif isinstance(node, list) and node:
            walk(node[0], f"{path}[0]")

    walk(t)
    for p, s in found:
        print(f"  {p}: {s}")
    if not found:
        print("  ключей со словом vacancy нет — связывать придётся иначе")

    # Раз lastState это наш признак ответа — посмотрим, какие значения реально
    # встречаются на странице и сколько их. Считаем по всем 20 записям.
    print("\n===== состояния на этой странице =====")
    for field in ("initialState", "lastState", "lastSubState", "archived"):
        counts = {}
        for item in topics:
            counts[item.get(field)] = counts.get(item.get(field), 0) + 1
        print(f"  {field}: {counts}")

    print("\n===== счётчики по всем откликам =====")
    c = data.get("applicantNegotiationsCounters") or {}
    for group in ("new", "total"):
        print(f"  {group}: {c.get(group)}")

    # Ссылки: как запросить остальные страницы и другие фильтры.
    print("\n===== постраничность =====")
    paging = neg.get("paging") or {}
    pages = paging.get("pages") or []
    if pages:
        print(f"  первая запись пагинации: {pages[0]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
