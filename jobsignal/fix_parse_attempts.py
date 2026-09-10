"""Счётчик попыток разбора у raw_posts + возврат потерянных постов в очередь.

Парсер при любом сбое разбора помечал пост parsed=1, и пост выбывал из
очереди навсегда: повторный прогон его не поднимал. Так на прогоне 10.09
потерялись четыре поста, где модель вернула массив вакансий вместо одного
объекта. Теперь при сбое parsed не ставится, а считается попытка — этой
колонки в таблице ещё нет, скрипт её добавляет.

Значение 0 для существующих постов означает «попыток не считали», а не
«сбоев не было»: до сегодняшнего дня их никто не учитывал. Единственное
исключение — четыре поста ниже: у них ровно одна неудачная попытка,
зафиксированная в логе прогона, и им проставляется 1.
"""
from __future__ import annotations

import shutil
import sqlite3
import sys
from datetime import datetime

DB = "data/jobsignal.db"

# Посты, потерянные на прогоне 10.09.2026: модель вернула массив вакансий,
# разбор упал, пост был закрыт как разобранный. Текст постов в базе цел.
LOST_POSTS = (7875, 7958, 7973, 7975)


def main() -> int:
    conn = sqlite3.connect(DB)
    cols = {r[1] for r in conn.execute("PRAGMA table_info(raw_posts)")}

    need_column = "parse_attempts" not in cols
    if not need_column:
        print("= parse_attempts: колонка уже есть")

    placeholders = ",".join("?" * len(LOST_POSTS))
    lost = conn.execute(
        f"SELECT id, parsed FROM raw_posts WHERE id IN ({placeholders})",
        LOST_POSTS,
    ).fetchall()
    to_reset = [r[0] for r in lost if r[1]]
    missing = set(LOST_POSTS) - {r[0] for r in lost}
    if missing:
        print(f"! в базе нет постов: {sorted(missing)}")
    if not to_reset:
        print("= потерянные посты: уже в очереди, сбрасывать нечего")

    if not need_column and not to_reset:
        conn.close()
        print("нечего делать")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = f"{DB}.bak-{stamp}"
    shutil.copy(DB, backup)
    print(f"копия: {backup}")

    if need_column:
        conn.execute(
            "ALTER TABLE raw_posts ADD COLUMN parse_attempts "
            "INTEGER NOT NULL DEFAULT 0"
        )
        print("+ parse_attempts")

    if to_reset:
        conn.executemany(
            "UPDATE raw_posts SET parsed = 0, parse_attempts = 1 WHERE id = ?",
            [(pid,) for pid in to_reset],
        )
        print(f"+ возвращены в очередь: {to_reset} (parse_attempts=1)")

    conn.commit()
    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
