"""Уведомление в телеграм о сорвавшемся юните — цель для OnFailure=.

Смысл ровно один: сломанный сбор должен доходить до человека, а не только
до журнала. Именно так мёртвый hh-сбор прожил незамеченным с конца июня.

Запускается как `alert_failed.py <имя-юнита>`; последние строки журнала
юнита прикладываются к сообщению.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / "config" / ".env")
sys.path.insert(0, str(ROOT))

from jobsignal.agents.notify_bot import _send  # noqa: E402


def _journal_tail(unit: str, lines: int = 12) -> str:
    try:
        out = subprocess.run(
            ["journalctl", "-u", unit, "-n", str(lines), "--no-pager",
             "-o", "cat"],
            capture_output=True, text=True, timeout=20,
        )
        return (out.stdout or "").strip()
    except Exception as exc:  # noqa: BLE001
        return f"(журнал не прочитать: {exc})"


# Автоотклик идёт каждые 5 минут: лежащий прокси без этого давал бы 12
# одинаковых тревог в час, и настоящая новая поломка тонула бы в них.
REPEAT_MIN = int(os.environ.get("ALERT_REPEAT_MIN", "60"))
STATE_DIR = ROOT / "data" / "alerts"


def _should_send(unit: str) -> tuple[bool, int]:
    """Слать ли тревогу и сколько повторов было подавлено с прошлой."""
    path = STATE_DIR / f"{unit.replace('/', '_')}.json"
    now = time.time()
    try:
        state = json.loads(path.read_text())
    except (OSError, ValueError):
        state = {}
    last = float(state.get("sent_at", 0))
    muted = int(state.get("muted", 0))
    send = now - last >= REPEAT_MIN * 60
    try:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(
            {"sent_at": now, "muted": 0} if send
            else {"sent_at": last, "muted": muted + 1}))
    except OSError as exc:
        print(f"состояние тревог не записано: {exc}")
        return True, muted
    return send, muted


def main() -> int:
    unit = sys.argv[1] if len(sys.argv) > 1 else "неизвестный юнит"
    send, muted = _should_send(unit)
    if not send:
        print(f"повтор тревоги {unit} в пределах {REPEAT_MIN} мин — не шлём")
        return 0
    tail = _journal_tail(unit)
    repeats = (f"\nС прошлой тревоги этот юнит падал ещё {muted} раз.\n"
               if muted else "")
    text = (f"🔴 <b>jobsignal: сбой</b>\n"
            f"Юнит: <code>{unit}</code>\n{repeats}\n"
            f"<pre>{tail[-2500:]}</pre>")
    ok = _send(text)
    print("отправлено" if ok else "отправить не удалось")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
