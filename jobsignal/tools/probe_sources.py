"""Разведка новых источников вакансий: что отвечает серверу и в каком виде.

Только GET страниц поиска, без входа и без прокси. Для каждого кандидата —
код ответа, размер, тип и признаки: JSON со списком вакансий или HTML, в
котором есть карточки/встроенное состояние. По результату решаем, какой
источник подключать коллектором.

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_sources.py ["product owner"]
"""
from __future__ import annotations

import json
import re
import sys
from urllib.parse import quote

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")


def candidates(q: str) -> list[tuple[str, str]]:
    e = quote(q)
    return [
        ("habr api", f"https://career.habr.com/api/frontend/vacancies?q={e}&type=all&per_page=25"),
        ("habr html", f"https://career.habr.com/vacancies?q={e}&type=all"),
        ("getmatch api", f"https://getmatch.ru/api/offers?q={e}&limit=20"),
        ("getmatch html", f"https://getmatch.ru/vacancies?q={e}"),
        ("superjob html", f"https://www.superjob.ru/vacancy/search/?keywords={e}"),
        ("trudvsem api", f"http://opendata.trudvsem.ru/api/v1/vacancies?text={e}&limit=20"),
        ("geekjob html", f"https://geekjob.ru/vacancies?qs={e}"),
        ("yandex jobs api", f"https://yandex.ru/jobs/api/publications?text={e}&page_size=20"),
        ("ozon career", f"https://career.ozon.ru/vacancy/?query={e}"),
        ("tbank career", f"https://www.tbank.ru/career/it/vacancies/?search={e}"),
        ("sber rabota", f"https://rabota.sber.ru/search/?query={e}"),
        ("avito career", f"https://career.avito.com/vacancies/?q={e}"),
        ("vk team", f"https://team.vk.company/vacancy/?search={e}"),
        ("alfa career", f"https://job.alfabank.ru/vacancies?search={e}"),
    ]


def look(name: str, url: str) -> None:
    try:
        r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9",
                                       "Accept": "application/json, text/html;q=0.9"},
                         timeout=25, allow_redirects=True)
    except Exception as exc:  # noqa: BLE001
        print(f"{name:<16} ✗ {type(exc).__name__}: {str(exc)[:120]}")
        return
    ctype = r.headers.get("content-type", "").split(";")[0]
    body = r.text or ""
    sign = []
    if "json" in ctype:
        try:
            data = r.json()
            if isinstance(data, dict):
                sign.append("ключи: " + ", ".join(list(data)[:8]))
                for k in ("list", "vacancies", "items", "offers", "results", "data"):
                    v = data.get(k)
                    if isinstance(v, list):
                        sign.append(f"{k}[{len(v)}]")
                        if v and isinstance(v[0], dict):
                            sign.append("поля: " + ", ".join(list(v[0])[:10]))
                        break
                    if isinstance(v, dict):
                        sub = [kk for kk, vv in v.items() if isinstance(vv, list)]
                        if sub:
                            sign.append(f"{k}.{sub[0]}[{len(v[sub[0]])}]")
            elif isinstance(data, list):
                sign.append(f"список[{len(data)}]")
        except ValueError:
            sign.append("json не разобран")
    else:
        for label, rx in (("__NEXT_DATA__", r'id="__NEXT_DATA__"'),
                          ("initial-state", r"__INITIAL_STATE__|window\.__STATE__|data-ssr"),
                          ("ld+json", r'application/ld\+json'),
                          ("captcha", r"captcha|ddos-guard|Just a moment|cf-chl"),
                          ("vacancy-links", r'href="[^"]*/vacanc[^"]*\d')):
            if re.search(rx, body, re.I):
                n = len(re.findall(rx, body, re.I))
                sign.append(f"{label}×{n}" if n > 1 else label)
        title = re.search(r"<title[^>]*>(.*?)</title>", body, re.S | re.I)
        if title:
            sign.append("title: " + " ".join(title.group(1).split())[:60])
    print(f"{name:<16} {r.status_code} {len(r.content) // 1024:>5} КБ {ctype:<18} "
          f"{r.url[:70]}\n{'':<16} {' | '.join(sign)[:400]}")


def main() -> int:
    q = " ".join(sys.argv[1:]) or "product owner"
    print(f"запрос: {q}\n")
    for name, url in candidates(q):
        look(name, url)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
