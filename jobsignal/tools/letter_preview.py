"""Посмотреть сопроводительные письма: что ушло и что пишется сейчас.

Зачем. 08.10 в генератор писем перенесли правила скилла human-voice: без
штампов, с опорой на текст вакансии, без вывода за читателя. Проверить, как
это звучит на деле, можно только глазами — линтер видит слова, а не
построение.

Два режима:
  * sent — последние письма, которые реально ушли (applications.message_text);
  * new  — свежая генерация под вакансии из очереди. Ничего не сохраняет и
           не отправляет, только печатает.

Под каждым письмом — замечания проверки (letters.review: обрыв, штампы,
длина канала, числа не из резюме).

    cd /opt/jobsignal_local && .venv/bin/python tools/letter_preview.py sent 5
    cd /opt/jobsignal_local && .venv/bin/python tools/letter_preview.py new 3
    cd /opt/jobsignal_local && .venv/bin/python tools/letter_preview.py reset

reset — стереть неотправленные черновики. Автоотклик берёт готовый черновик,
если он есть, поэтому после правки правил голоса старые письма ушли бы как
есть. Отправленные письма не трогает.
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


def _show(title: str, text: str, issues: list | None = None) -> None:
    print("=" * 72)
    print(title)
    print("-" * 72)
    print(text.strip() or "(письма нет)")
    if issues is None:
        issues = style_tells(text)
    print("-" * 72)
    print(f"{len(text.strip())} знаков · замечания: "
          + (", ".join(issues) if issues else "нет"))
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
    """Свежая генерация тем же сервисом, что автоотклик (jobsignal/letters.py),
    по лучшему баллу среди ДЕЙСТВУЮЩИХ профилей и по каналу вакансии."""
    from jobsignal import letters, profiles
    active = profiles.names()
    c = sqlite3.connect(DB)
    marks = ",".join("?" * len(active)) or "''"
    rows = c.execute(
        f"""select v.id, v.role, v.company, v.contact_type,
                  (select profile_key from match_scores where vacancy_id = v.id
                     and profile_key in ({marks}) order by score desc limit 1),
                  (select max(score) from match_scores where vacancy_id = v.id
                     and profile_key in ({marks})) best
             from vacancies v
            where v.is_primary = 1 and v.status in ('matched', 'drafted')
              and best is not null
            order by best desc, v.id desc limit ?""", (*active, *active, n)).fetchall()
    stale = c.execute(
        f"""select count(*) from vacancies v
             where v.is_primary = 1 and v.status in ('matched', 'drafted')
               and exists (select 1 from match_scores m where m.vacancy_id = v.id)
               and not exists (select 1 from match_scores m where m.vacancy_id = v.id
                                 and m.profile_key in ({marks}))""", active).fetchone()[0]
    if stale:
        print(f"в очереди {stale} вакансий с оценками только под снятые профили — "
              f"они не в работе; переоценка: tools/rescore.py --profiles-changed\n")
    if not rows:
        print("в очереди нет вакансий с оценкой под действующие профили")
        return 0
    channel_of = {"hh": "hh", "tg": "telegram", "form": "form", "email": "email"}
    for vid, role, co, ct, prof, score in rows:
        ch = channel_of.get(ct or "", "hh")
        letter = letters.for_vacancy(vid, channel=ch, profile=prof)
        _show(f"#{vid} · {score}% · {role or '—'} · {co or '—'}\n"
              f"профиль «{letter.profile}» · канал {letter.channel}",
              letter.text, letter.issues)
    return 0


def reset() -> int:
    c = sqlite3.connect(DB)
    n = c.execute(
        """update vacancies set draft_text = null
            where draft_text is not null
              and status in ('matched', 'drafted')
              and id not in (select vacancy_id from applications
                              where sent_at is not null)""").rowcount
    c.commit()
    print(f"неотправленных черновиков стёрто: {n} — перепишутся при отправке")
    return 0


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "new"
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    if mode == "sent":
        return sent(n)
    if mode == "new":
        return new(n)
    if mode == "reset":
        return reset()
    print("режим: sent, new или reset")
    return 1


if __name__ == "__main__":
    sys.exit(main())
