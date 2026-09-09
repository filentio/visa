"""Уникальные индексы на applications вместо INSERT OR IGNORE.

OR IGNORE гасил любое нарушение ограничений, а не только повтор: строка с
пропущенным NOT NULL-полем молча не создавалась. Защиту от дублей переносим
в схему, а код пишет обычный INSERT — теперь ошибка схемы видна.

Индексы частичные: у вакансии законно бывает черновик и отправленный отклик
одновременно. Запрещён только второй отправленный и второй черновик.

Данные скрипт не меняет — только создаёт индексы. Если в базе уже есть
нарушения, он ничего не создаёт и печатает их: решать, что с ними делать,
должен человек.
"""
from __future__ import annotations

import shutil
import sqlite3
import sys
from datetime import datetime

DB = "data/jobsignal.db"

INDEXES = [
    ("ux_applications_sent_vacancy",
     "CREATE UNIQUE INDEX ux_applications_sent_vacancy "
     "ON applications (vacancy_id) WHERE sent_at IS NOT NULL",
     "SELECT vacancy_id, COUNT(*) FROM applications "
     "WHERE sent_at IS NOT NULL GROUP BY vacancy_id HAVING COUNT(*) > 1"),
    ("ux_applications_draft_vacancy",
     "CREATE UNIQUE INDEX ux_applications_draft_vacancy "
     "ON applications (vacancy_id) WHERE is_draft = 1",
     "SELECT vacancy_id, COUNT(*) FROM applications "
     "WHERE is_draft = 1 GROUP BY vacancy_id HAVING COUNT(*) > 1"),
]


def main() -> int:
    conn = sqlite3.connect(DB)
    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index'")}

    todo = []
    for name, ddl, check in INDEXES:
        if name in have:
            print(f"= {name}: уже есть")
            continue
        bad = conn.execute(check).fetchall()
        if bad:
            print(f"! {name}: нарушения, индекс не создаётся: {bad}")
            continue
        todo.append((name, ddl))

    if not todo:
        conn.close()
        print("нечего делать")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = f"{DB}.bak-{stamp}"
    shutil.copy(DB, backup)
    print(f"копия: {backup}")

    for name, ddl in todo:
        conn.execute(ddl)
        print(f"+ {name}")
    conn.commit()
    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
