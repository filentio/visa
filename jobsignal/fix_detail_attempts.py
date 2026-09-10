"""Счётчик попыток догрузить описание вакансии с hh.ru.

hh иногда отдаёт вместо страницы вакансии заглушку со входом: код 200,
блок состояния на месте, но vacancyView в нём нет. Соседние вакансии в тот
же момент открываются — значит сбор цел, просто эту страницу нам не дают.

Без счётчика такой пост остаётся в долгах навсегда: каждый прогон он
занимает слот бюджета HH_DESC_LIMIT и каждый прогон получает ту же
заглушку. Колонка повторяет parse_attempts у парсера — после
HH_DETAIL_GIVE_UP попыток пост уходит в разбор с тем, что есть.

Задним числом ничего не проставляется: сколько раз система уже билась в
закрытую страницу до появления счётчика, в базе не записано, и выдумывать
это число нельзя. Все существующие посты начинают с нуля попыток.
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
    if "detail_attempts" in cols:
        conn.close()
        print("= detail_attempts: колонка уже есть — нечего делать")
        return 0

    waiting = conn.execute(
        "SELECT count(*) FROM raw_posts WHERE awaiting_details = 1"
    ).fetchone()[0]
    print(f"= постов в долгах по описаниям: {waiting} (все начнут с 0 попыток)")

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = f"{DB}.bak-{stamp}"
    shutil.copy(DB, backup)
    print(f"копия: {backup}")

    conn.execute(
        "ALTER TABLE raw_posts ADD COLUMN detail_attempts "
        "INTEGER NOT NULL DEFAULT 0"
    )
    print("+ detail_attempts")
    conn.commit()
    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
