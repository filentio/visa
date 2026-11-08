"""Воронка откликов: что получает ответы и сколько вакансий даёт каждый порог.

Зачем. Порог автоотклика (HH_AUTO_THRESHOLD, 78%) выставлен на глаз. Отчёт
раскладывает отправленные отклики по лучшему баллу вакансии, профилю,
источнику и наличию письма и показывает долю ответов и приглашений в каждой
группе, а рядом — сколько свежих вакансий в сутки проходит каждый порог.

    cd /opt/jobsignal_local && .venv/bin/python tools/funnel_report.py [--days 60]

Только чтение базы. Балл — лучший среди всех оценок вакансии (и снятых
профилей тоже): решение об отклике принималось по нему.
"""
from __future__ import annotations

import argparse
import sqlite3
from collections import defaultdict
from datetime import datetime, timedelta, timezone

BUCKETS = [(0, 59), (60, 69), (70, 74), (75, 77), (78, 81), (82, 85), (86, 89), (90, 100)]
POSITIVE = ("INVITATION", "INTERVIEW", "HIRED")


def bucket(score: int) -> str:
    for lo, hi in BUCKETS:
        if lo <= score <= hi:
            return f"{lo}-{hi}"
    return "?"


def pct(a: int, b: int) -> str:
    return f"{100 * a / b:4.0f}%" if b else "   —"


def table(title: str, rows: dict[str, list], order=None) -> None:
    print(f"\n{title}")
    print(f"  {'группа':<28} {'откл':>5} {'ответ':>6} {'доля':>5} {'приглаш':>8} {'доля':>5}")
    keys = order or sorted(rows, key=lambda k: -len(rows[k]))
    for k in keys:
        apps = rows.get(k) or []
        if not apps:
            continue
        n = len(apps)
        rep = sum(1 for a in apps if a["replied"])
        pos = sum(1 for a in apps if a["state"] in POSITIVE)
        print(f"  {k[:28]:<28} {n:>5} {rep:>6} {pct(rep, n)} {pos:>8} {pct(pos, n)}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=90, help="отклики за столько дней")
    opts = ap.parse_args()
    c = sqlite3.connect("data/jobsignal.db")
    since = (datetime.now(timezone.utc) - timedelta(days=opts.days)).strftime("%Y-%m-%d")

    best = {vid: (sc or 0, prof or "")
            for vid, sc, prof in c.execute(
                "SELECT vacancy_id, score, profile_key FROM match_scores m1 WHERE score = "
                "(SELECT MAX(score) FROM match_scores m2 WHERE m2.vacancy_id = m1.vacancy_id)")}

    apps = []
    for vid, channel, text, sent, replied, state, ctype in c.execute(
            "SELECT a.vacancy_id, a.channel, a.message_text, a.sent_at, a.replied_at, "
            "a.reply_state, v.contact_type FROM applications a "
            "JOIN vacancies v ON v.id = a.vacancy_id "
            "WHERE a.sent_at IS NOT NULL AND a.sent_at >= ?", (since,)):
        sc, prof = best.get(vid, (0, ""))
        st = (state or "").upper()
        apps.append({"score": sc, "profile": prof or "—", "channel": channel or "—",
                     "source": ctype or "—", "letter": bool((text or "").strip()),
                     "replied": bool(replied) and st != "SILENT", "state": st,
                     "sent": sent, "replied_at": replied})

    n = len(apps)
    rep = sum(a["replied"] for a in apps)
    pos = sum(a["state"] in POSITIVE for a in apps)
    print(f"Отклики за {opts.days} дн.: {n}, ответили {rep} ({pct(rep, n).strip()}), "
          f"приглашений {pos} ({pct(pos, n).strip()})")
    states = defaultdict(int)
    for a in apps:
        states[a["state"] or "нет данных"] += 1
    print("Состояния на hh: " + ", ".join(f"{k} {v}" for k, v in sorted(states.items())))

    by = defaultdict(list)
    for a in apps:
        by[bucket(a["score"])].append(a)
    table("По лучшему баллу вакансии", by, [f"{lo}-{hi}" for lo, hi in BUCKETS])

    for key, title in (("profile", "По профилю (лучшая оценка)"),
                       ("channel", "По каналу отклика"),
                       ("source", "По источнику вакансии")):
        g = defaultdict(list)
        for a in apps:
            g[a[key]].append(a)
        table(title, g)
    g = defaultdict(list)
    for a in apps:
        g["с письмом" if a["letter"] else "без письма"].append(a)
    table("Письмо", g)

    # Сколько вакансий в сутки проходит каждый порог (последние 14 дней, hh).
    since14 = (datetime.now(timezone.utc) - timedelta(days=14)).strftime("%Y-%m-%d")
    fresh = defaultdict(int)
    for vid, ctype in c.execute(
            "SELECT id, contact_type FROM vacancies WHERE is_primary = 1 AND created_at >= ?",
            (since14,)):
        sc = best.get(vid, (None, ""))[0]
        if sc is None:
            continue
        for thr in (65, 70, 72, 75, 78, 80, 85, 90):
            if sc >= thr:
                fresh[(ctype or "—", thr)] += 1
    print("\nСвежих вакансий в сутки выше порога (за 14 дней):")
    srcs = sorted({k[0] for k in fresh})
    print("  порог  " + "".join(f"{s[:10]:>11}" for s in srcs))
    for thr in (65, 70, 72, 75, 78, 80, 85, 90):
        print(f"  {thr:>4}%  " + "".join(f"{fresh[(s, thr)] / 14:>11.1f}" for s in srcs))

    # Очередь автоотклика при разных порогах (hh, ещё не отправленные).
    print("\nОчередь hh сейчас (matched/drafted) при пороге:")
    q = list(c.execute(
        "SELECT v.id FROM vacancies v WHERE v.is_primary = 1 AND v.contact_type = 'hh' "
        "AND v.status IN ('matched', 'drafted')"))
    for thr in (70, 72, 75, 78, 80, 85):
        print(f"  {thr}%: {sum(1 for (vid,) in q if best.get(vid, (0, ''))[0] >= thr)}")

    # Запас без оценки: вакансии new по типу контакта и возрасту.
    print("\nНе оценены (status new) по типу контакта и возрасту:")
    now = datetime.now(timezone.utc)
    ages = defaultdict(lambda: [0, 0, 0, 0])
    for ctype, created in c.execute(
            "SELECT contact_type, created_at FROM vacancies WHERE is_primary = 1 "
            "AND status = 'new'"):
        try:
            d = datetime.fromisoformat(str(created)[:19]).replace(tzinfo=timezone.utc)
            days = (now - d).days
        except ValueError:
            days = 999
        row = ages[ctype or "—"]
        row[0 if days <= 7 else 1 if days <= 14 else 2 if days <= 30 else 3] += 1
    print(f"  {'тип':<12} {'≤7д':>6} {'8-14д':>6} {'15-30д':>7} {'старше':>7}")
    for k, r in sorted(ages.items(), key=lambda kv: -sum(kv[1])):
        print(f"  {k:<12} {r[0]:>6} {r[1]:>6} {r[2]:>7} {r[3]:>7}")

    # Время до ответа.
    lags = []
    for a in apps:
        if a["replied"] and a["replied_at"] and a["sent"]:
            try:
                d = (datetime.fromisoformat(str(a["replied_at"])[:19])
                     - datetime.fromisoformat(str(a["sent"])[:19]))
                lags.append(d.total_seconds() / 86400)
            except ValueError:
                pass
    if lags:
        lags.sort()
        print(f"\nДней до ответа: медиана {lags[len(lags) // 2]:.1f}, "
              f"90% — до {lags[int(len(lags) * 0.9) - 1 if len(lags) > 1 else 0]:.1f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
