"""Посмотреть сопроводительные письма: что ушло и что пишется сейчас.

Зачем. 08.10 в генератор писем перенесли правила скилла human-voice: без
штампов, с опорой на текст вакансии, без вывода за читателя. Проверить, как
это звучит на деле, можно только глазами — линтер видит слова, а не
построение.

Два режима:
  * sent — последние письма, которые реально ушли (applications.message_text);
  * new  — свежая генерация под вакансии из очереди. Ничего не сохраняет и
           не отправляет, только печатает.

Под каждым письмом — найденные шаблонные маркеры (letter_style.style_tells).

    cd /opt/jobsignal_local && .venv/bin/python tools/letter_preview.py sent 5
    cd /opt/jobsignal_local && .venv/bin/python tools/letter_preview.py new 3
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from jobsignal.letter_style import style_tells  # noqa: E402

DB = "data/jobsignal.db"


def _show(title: str, text: str) -> None:
    print("=" * 72)
    print(title)
    print("-" * 72)
    print(text.strip())
    tells = style_tells(text)
    print("-" * 72)
    print(f"{len(text.strip())} знаков · маркеры: "
          + (", ".join(tells) if tells else "нет"))
    print()


def sent(n: int) -> int:
    c = sqlite3.connect(DB)
    rows = c.execute(
        """select a.sent_at, a.channel, v.role, v.company, a.message_text
             from applications a join vacancies v on v.id = a.vacancy_id
            where a.sent_at is not null and coalesce(a.message_text,'') != ''
            order by a.sent_at desc limit ?""", (n,)).fetchall()
    if not rows:
        print("отправленных писем пока нет")
        return 0
    for when, ch, role, co, text in rows:
        _show(f"{str(when)[:16]} · {ch} · {role or '—'} · {co or '—'}", text)
    return 0


def new(n: int) -> int:
    from jobsignal.agents.notify_bot import _cover_text
    c = sqlite3.connect(DB)
    rows = c.execute(
        """select v.id, v.role, v.company, coalesce(p.text, v.description, ''),
                  (select profile_key from match_scores
                    where vacancy_id = v.id order by score desc limit 1),
                  (select max(score) from match_scores where vacancy_id = v.id)
             from vacancies v left join raw_posts p on p.id = v.raw_post_id
            where v.is_primary = 1 and v.status in ('matched', 'drafted')
            order by 6 desc, v.id desc limit ?""", (n,)).fetchall()
    if not rows:
        print("в очереди нет вакансий со статусом matched/drafted — "
              "генерировать не подо что")
        return 0
    for vid, role, co, text, prof, score in rows:
        letter = _cover_text(role or "", co or "", prof or "Senior PM/PO",
                             vacancy_text=text or "")
        _show(f"#{vid} · {score}% · {role or '—'} · {co or '—'} · профиль {prof}",
              letter)
    return 0


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "new"
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    if mode == "sent":
        return sent(n)
    if mode == "new":
        return new(n)
    print("режим: sent или new")
    return 1


if __name__ == "__main__":
    sys.exit(main())
