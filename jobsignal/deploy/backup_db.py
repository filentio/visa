"""Ежедневная копия базы.

Зачем именно сейчас. 04.10 ошибочное удаление снесло 72 настоящие отметки
откликов. Выручила копия — но самая свежая оказалась от 10 сентября, то есть
спасла случайность, а не устройство системы. Отметки за три недели между
копией и аварией восстановить было уже нечем.

Почему не `cp`. База работает в режиме WAL: копирование файла во время записи
даёт копию без хвоста журнала, местами битую, и узнаёшь об этом в тот
единственный день, когда копия понадобилась.

Почему питон, а не утилита sqlite3. Её в системе нет, а ставить пакет ради
одной команды незачем: `Connection.backup()` делает ровно тот же
согласованный снимок на горячей базе, и интерпретатор уже стоит.

    .venv/bin/python deploy/backup_db.py
"""
from __future__ import annotations

import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

DB = Path(os.environ.get("JOBSIGNAL_DB", "data/jobsignal.db"))
DIR = Path(os.environ.get("JOBSIGNAL_BACKUP_DIR", "data/backups"))
# Неделя: ошибку замечают в тот же день или на следующий, а семь файлов по
# полсотни мегабайт диск при двух гигабайтах терпит. Месяц — уже нет.
KEEP = int(os.environ.get("JOBSIGNAL_BACKUP_KEEP", "7"))


def main() -> int:
    if not DB.exists():
        print(f"нет базы {DB}")
        return 1
    DIR.mkdir(parents=True, exist_ok=True)
    out = DIR / f"jobsignal-{datetime.now():%Y%m%d-%H%M%S}.db"

    src = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
    try:
        dst = sqlite3.connect(out)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()

    # Снимок проверяем, а не верим отсутствию исключения: копия, которую не
    # открыть, это отсутствие копии, и знать об этом надо сегодня, а не в
    # день аварии.
    chk = sqlite3.connect(out)
    try:
        ok = chk.execute("pragma quick_check").fetchone()[0]
        vac = chk.execute("select count(*) from vacancies").fetchone()[0]
        app = chk.execute("select count(*) from applications").fetchone()[0]
    finally:
        chk.close()
    if ok != "ok":
        print(f"копия {out} не прошла проверку: {ok}")
        out.unlink(missing_ok=True)
        return 1
    print(f"копия готова: {out} ({out.stat().st_size / 2**20:.0f} МБ, "
          f"вакансий {vac}, откликов {app})")

    old = sorted(DIR.glob("jobsignal-*.db"),
                 key=lambda p: p.stat().st_mtime, reverse=True)[KEEP:]
    for p in old:
        print(f"удаляю старую копию: {p}")
        p.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
