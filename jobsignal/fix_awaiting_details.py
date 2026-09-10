"""Признак «пост ждёт описания» + пометка уже собранных hh-постов без него.

У hh описание вакансии лежит на отдельной странице, и за прогон их
догружается не больше HH_DESC_LIMIT. Посты, которым описания не хватило,
раньше всё равно уходили в парсер — и матчер оценивал вакансию по одному
заголовку, ровно от чего уходили. Теперь такие посты помечаются и ждут
следующего прогона.

Скрипт добавляет колонку и проставляет признак постам канала hh_ru, которые
уже собраны, ещё не разобраны и не содержат строки «Описание:». Телеграм-постов
это не касается: там текст объявления приходит целиком сразу.
"""
from __future__ import annotations

import shutil
import sqlite3
import sys
from datetime import datetime

DB = "data/jobsignal.db"


def main() -> int:
    conn = sqlite3.connect(DB)
    cols = {r[1] for r in conn.execute("PRAGMA table_info(raw_posts)")}
    need_column = "awaiting_details" not in cols
    if not need_column:
        print("= awaiting_details: колонка уже есть")

    ch = conn.execute(
        "SELECT id FROM channels WHERE handle = 'hh_ru'"
    ).fetchone()
    if not ch:
        conn.close()
        print("! канала hh_ru нет — помечать нечего")
        return 0

    to_mark = conn.execute(
        "SELECT count(*) FROM raw_posts "
        "WHERE channel_id = ? AND parsed = 0 AND text NOT LIKE '%Описание:%'",
        (ch[0],),
    ).fetchone()[0]
    print(f"= постов hh без описания в очереди: {to_mark}")

    if not need_column and not to_mark:
        conn.close()
        print("нечего делать")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = f"{DB}.bak-{stamp}"
    shutil.copy(DB, backup)
    print(f"копия: {backup}")

    if need_column:
        conn.execute(
            "ALTER TABLE raw_posts ADD COLUMN awaiting_details "
            "INTEGER NOT NULL DEFAULT 0"
        )
        print("+ awaiting_details")

    if to_mark:
        n = conn.execute(
            "UPDATE raw_posts SET awaiting_details = 1 "
            "WHERE channel_id = ? AND parsed = 0 "
            "AND text NOT LIKE '%Описание:%'",
            (ch[0],),
        ).rowcount
        print(f"+ помечено ждущими описания: {n}")

    conn.commit()
    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
