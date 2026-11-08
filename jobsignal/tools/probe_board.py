"""Разведка одной площадки: первая вакансия из выдачи целиком и её страница.

    .venv/bin/python tools/probe_board.py habr|getmatch|geekjob|avito ["запрос"]

Печатает первую карточку выдачи (JSON до 3000 знаков) и про страницу вакансии:
есть ли ld+json JobPosting и сколько знаков в его описании. Только GET.
"""
from __future__ import annotations

import json
import re
import sys
from urllib.parse import quote

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
H = {"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9"}


def ld_jobposting(html: str) -> dict | None:
    for m in re.finditer(r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', html, re.S):
        try:
            data = json.loads(m.group(1))
        except ValueError:
            continue
        for d in (data if isinstance(data, list) else [data]):
            if isinstance(d, dict) and d.get("@type") == "JobPosting":
                return d
    return None


def page(url: str) -> None:
    r = requests.get(url, headers=H, timeout=25)
    print(f"\nстраница {url}: {r.status_code}, {len(r.content) // 1024} КБ")
    jp = ld_jobposting(r.text)
    if jp:
        desc = re.sub(r"<[^>]+>", " ", jp.get("description") or "")
        print(f"ld+json JobPosting: ключи {list(jp)[:14]}")
        print(f"описание {len(desc)} знаков: {' '.join(desc.split())[:400]}")
    else:
        print("ld+json JobPosting нет; next-data:", bool(re.search(r'__NEXT_DATA__', r.text)))
        for rx in (r'class="[^"]*vacancy-description[^"]*"', r'class="[^"]*description[^"]*"'):
            m = re.search(rx, r.text)
            if m:
                print("блок описания:", m.group(0)[:80])
                break


def main() -> int:
    name = sys.argv[1] if len(sys.argv) > 1 else "habr"
    q = quote(" ".join(sys.argv[2:]) or "product owner")
    if name == "habr":
        d = requests.get(f"https://career.habr.com/api/frontend/vacancies?q={q}&type=all"
                         "&per_page=25&sort=date", headers=H, timeout=25).json()
        item = d["list"][0]
        print("meta:", json.dumps(d.get("meta"), ensure_ascii=False)[:300])
        print(json.dumps(item, ensure_ascii=False, indent=1)[:3000])
        page("https://career.habr.com" + item["href"])
    elif name == "getmatch":
        d = requests.get(f"https://getmatch.ru/api/offers?q={q}&limit=20", headers=H,
                         timeout=25).json()
        item = d["offers"][0]
        print("meta:", json.dumps(d.get("meta"), ensure_ascii=False)[:300])
        print(json.dumps(item, ensure_ascii=False, indent=1)[:3000])
        url = item.get("url") or f"https://getmatch.ru/vacancies/{item.get('id')}"
        page(url if url.startswith("http") else "https://getmatch.ru" + url)
    elif name in ("geekjob", "avito"):
        url = (f"https://geekjob.ru/vacancies?qs={q}" if name == "geekjob"
               else f"https://career.avito.com/vacancies/?q={q}")
        html = requests.get(url, headers=H, timeout=25).text
        links = sorted(set(re.findall(
            r'href="((?:https?://[^"]+)?/vacanc[^"]*?/[\w-]*\d[\w-]*/?)"', html)))
        print(f"ссылок на вакансии: {len(links)}; первые: {links[:5]}")
        if links:
            first = links[0]
            base = "https://geekjob.ru" if name == "geekjob" else "https://career.avito.com"
            page(first if first.startswith("http") else base + first)
    else:
        print("площадка: habr | getmatch | geekjob | avito")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
