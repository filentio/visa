"""Какие сайты и системы анкет стоят за вакансиями типа «анкета/ссылка».

Зачем. Около 95 вакансий в день приходят со ссылкой вместо контакта:
Huntflow, Google Forms, страницы карьеры компаний, ссылки на посты. Чтобы
автоматизировать отклик, сначала надо знать, что встречается чаще — особенно
среди вакансий, которые прошли оценку.

    .venv/bin/python tools/form_domains.py [--days 30]

Только чтение базы.
"""
from __future__ import annotations

import argparse
import re
import sqlite3
from collections import Counter, defaultdict
from urllib.parse import urlparse

# Системы анкет по домену/пути — для подписи в отчёте.
ATS = [
    ("huntflow", r"huntflow\.(ru|io)"),
    ("google forms", r"docs\.google\.com/forms|forms\.gle"),
    ("yandex forms", r"forms\.yandex\.ru"),
    ("talantix", r"talantix\.ru"),
    ("potok", r"potok\.io|app\.potok"),
    ("friend.work", r"friend\.work"),
    ("e-staff", r"e-staff\.ru"),
    ("skillaz", r"skillaz\.ru"),
    ("greenhouse", r"greenhouse\.io"),
    ("lever", r"lever\.co"),
    ("workable", r"workable\.com"),
    ("teamtailor", r"teamtailor\.com"),
    ("notion", r"notion\.(site|so)"),
    ("tally", r"tally\.so"),
    ("typeform", r"typeform\.com"),
    ("telegram post", r"t\.me/"),
    ("hh (not vacancy)", r"hh\.ru"),
    ("linkedin", r"linkedin\.com"),
]


def kind(url: str) -> str:
    for name, rx in ATS:
        if re.search(rx, url, re.I):
            return name
    return "сайт компании / прочее"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=30)
    opts = ap.parse_args()
    c = sqlite3.connect("data/jobsignal.db")
    rows = c.execute(
        "SELECT v.id, v.link, v.status, "
        " (SELECT MAX(score) FROM match_scores m WHERE m.vacancy_id = v.id) "
        "FROM vacancies v LEFT JOIN raw_posts r ON r.id = v.raw_post_id "
        "WHERE v.is_primary = 1 AND v.contact_type = 'form' "
        "AND COALESCE(r.posted_at, v.created_at) >= datetime('now', ?)",
        (f"-{opts.days} days",)).fetchall()
    total, good = Counter(), Counter()
    hosts, good_hosts = Counter(), Counter()
    examples: dict[str, list] = defaultdict(list)
    for vid, link, status, score in rows:
        link = link or ""
        k = kind(link)
        host = (urlparse(link if "://" in link else "https://" + link).hostname or "—").lower()
        host = re.sub(r"^www\.", "", host)
        total[k] += 1
        hosts[host] += 1
        if (score or 0) >= 70:
            good[k] += 1
            good_hosts[host] += 1
            if len(examples[k]) < 3:
                examples[k].append(link[:110])
    print(f"Вакансий «анкета/ссылка» за {opts.days} дн.: {len(rows)}, "
          f"из них с оценкой ≥70: {sum(good.values())}\n")
    print(f"{'система':<26}{'всего':>7}{'≥70':>6}")
    for k, n in total.most_common():
        print(f"{k:<26}{n:>7}{good[k]:>6}")
    print("\nСайты среди подходящих (≥70):")
    for h, n in good_hosts.most_common(25):
        print(f"  {n:>4}  {h}")
    print("\nСайты среди всех (топ-25):")
    for h, n in hosts.most_common(25):
        print(f"  {n:>4}  {h}")
    print("\nПримеры ссылок подходящих:")
    for k, links in examples.items():
        for ln in links:
            print(f"  [{k}] {ln}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
