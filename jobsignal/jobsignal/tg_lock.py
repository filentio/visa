"""Замок на файл сессии Telegram (data/tg_inbox.session).

Зачем. Сессия Telethon — это SQLite-файл, и два процесса с ним одновременно
ловят «database is locked». До поиска каналов это почти не случалось: сторож
входящих работает секунды раз в 5 минут, сверка исходящих — редко. Поиск
каналов держит сессию минутами (глобальный поиск, обход подписок), поэтому
всем, кто открывает сессию, нужен общий замок.

Сторож при занятом замке пропускает прогон (через 5 минут будет следующий),
остальные ждут.
"""
from __future__ import annotations

import fcntl
import os
import time
from contextlib import contextmanager
from pathlib import Path


def _lock_path() -> Path:
    custom = os.environ.get("TG_LOCK_PATH")
    if custom:
        return Path(custom)
    run = Path("/run")
    if run.is_dir() and os.access(run, os.W_OK):
        return run / "jobsignal-tg.lock"
    return Path("data/jobsignal-tg.lock")


@contextmanager
def tg_session_lock(wait_s: float = 0):
    """Взять замок; yield True — взят, False — не дождались за wait_s секунд."""
    path = _lock_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    fh = open(path, "w")
    got = False
    deadline = time.monotonic() + wait_s
    try:
        while True:
            try:
                fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
                got = True
                break
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    break
                time.sleep(1)
        yield got
    finally:
        if got:
            fcntl.flock(fh, fcntl.LOCK_UN)
        fh.close()
