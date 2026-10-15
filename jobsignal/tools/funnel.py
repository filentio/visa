"""Воронка hh за последние сутки: где теряются вакансии до отклика.

Зачем. Вопрос «почему сегодня нет откликов» возникает регулярно, и ответ
каждый раз собирался из пяти команд: журнал отправки, журнал сбора, запросы к
базе. А причина почти всегда в одном звене: не собрали, не догрузили
описание, не разобрали, не оценили, оценили ниже порога или отсекли как дубль.
Этот отчёт показывает все звенья разом.

    cd /opt/jobsignal_local && .venv/bin/python tools/funnel.py
    cd /opt/jobsignal_local && .venv/bin/python tools/funnel.py 72   # за 72 часа
"""
from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

DB = "data/jobsignal.db"


def main() -> int:
    hours = int(sys.argv[1]) if len(sys.argv) > 1 else 24
    since = f"-{hours} hours"
    auto = int(os.environ.get("HH_AUTO_THRESHOLD", "80"))
    match = int(os.environ.get("MATCH_THRESHOLD", "70"))
    c = sqlite3.connect(DB)
    q = lambda sql, *a: c.execute(sql, a).fetchone()[0]  # noqa: E731

    hh_ch = c.execute("select id from channels where handle in ('hh_ru','hh.ru','hh')"
                      ).fetchone()
    hh_id = hh_ch[0] if hh_ch else -1

    print(f"=== воронка hh за {hours} ч (порог оценки {match}, автоотклик от {auto}) ===\n")
    print("1. собрано постов с hh:        ",
          q("select count(*) from raw_posts where channel_id=? "
            "and posted_at > datetime('now', ?)", hh_id, since))
    print("   ждут описания:              ",
          q("select count(*) from raw_posts where channel_id=? "
            "and awaiting_details=1", hh_id))
    print("   не разобраны:               ",
          q("select count(*) from raw_posts where channel_id=? and parsed=0 "
            "and awaiting_details=0", hh_id))
    # Разборщик общий для всех источников: если он тонет в телеграме, hh
    # стоит в той же очереди. Поэтому общий хвост рядом.
    print("   очередь разбора всего:      ",
          q("select count(*) from raw_posts where parsed=0 and awaiting_details=0 "
            "and text is not null and coalesce(parse_attempts,0) < 3"),
          f"(за прогон берётся {os.environ.get('PARSER_BATCH_LIMIT', '200')})")

    base = ("from vacancies v where v.contact_type='hh' and v.is_primary=1 "
            "and v.created_at > datetime('now', ?)")
    print("\n2. вакансий hh создано:        ", q("select count(*) " + base, since))
    print("   по статусам:                ",
          dict(c.execute("select v.status, count(*) " + base +
                         " group by 1 order by 2 desc", (since,)).fetchall()))
    print("   скрыто как дубль:           ",
          q("select count(*) from vacancies where contact_type='hh' "
            "and is_primary=0 and created_at > datetime('now', ?)", since))

    rows = c.execute(
        "select (select max(score) from match_scores where vacancy_id=v.id) s "
        + base, (since,)).fetchall()
    scores = [r[0] for r in rows if r[0] is not None]
    if scores:
        buckets = {f"от {auto}": 0, f"{match}-{auto-1}": 0, f"ниже {match}": 0}
        for s in scores:
            if s >= auto:
                buckets[f"от {auto}"] += 1
            elif s >= match:
                buckets[f"{match}-{auto-1}"] += 1
            else:
                buckets[f"ниже {match}"] += 1
        print(f"\n3. оценено:                     {len(scores)} · {buckets}")
        print("   лучшие баллы:               ", sorted(scores, reverse=True)[:10])
    else:
        print("\n3. оценено:                     0")

    print(f"\n4. в очереди автоотклика (от {auto}, всё время):",
          q("select count(*) from vacancies v where contact_type='hh' "
            "and is_primary=1 and status in ('matched','drafted') "
            "and (select max(score) from match_scores where vacancy_id=v.id) >= ? "
            "and id not in (select vacancy_id from applications "
            "where sent_at is not null)", auto))

    print("\n5. откликов отправлено:        ",
          dict(c.execute("select channel, count(*) from applications "
                         "where sent_at > datetime('now', ?) group by 1",
                         (since,)).fetchall()) or 0)
    try:
        print("   попыток и их исходы:        ",
              dict(c.execute("select result, count(*) from hh_apply_attempts "
                             "where created_at > datetime('now', ?) group by 1",
                             (since,)).fetchall()) or "нет")
    except sqlite3.OperationalError:
        print("   попыток и их исходы:         таблицы ещё нет")
    return 0


if __name__ == "__main__":
    sys.exit(main())
