"""Уведомление в телеграм о сорвавшемся юните — цель для OnFailure=.

Смысл ровно один: сломанный сбор должен доходить до человека, а не только
до журнала. Именно так мёртвый hh-сбор прожил незамеченным с конца июня.

Запускается как `alert_failed.py <имя-юнита>`; последние строки журнала
юнита прикладываются к сообщению.
"""
from __future__ import annotations

import subprocess
import sys
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


def main() -> int:
    unit = sys.argv[1] if len(sys.argv) > 1 else "неизвестный юнит"
    tail = _journal_tail(unit)
    text = (f"🔴 <b>jobsignal: сбой</b>\n"
            f"Юнит: <code>{unit}</code>\n\n"
            f"<pre>{tail[-2500:]}</pre>")
    ok = _send(text)
    print("отправлено" if ok else "отправить не удалось")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
