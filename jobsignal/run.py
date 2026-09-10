"""JobSignal CLI entry point."""
import logging
import os
import sys

from dotenv import load_dotenv
load_dotenv("config/.env")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s INFO jobsignal: %(message)s",
    datefmt="%H:%M:%S",
)

COMMANDS = (
    "initdb", "seed-channels", "collect", "posts",
    "raw-dedup", "parse", "dedup", "match", "compose",
    "hh-collect", "find-channels",          # NEW: search tgstat/telemetr
    "hh-apply", "hh-preview",                # NEW: автоотклик на hh.ru
    "channels",               # NEW: list channels
    "pipeline", "status", "dashboard", "serve",
)


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    arg = sys.argv[2] if len(sys.argv) > 2 else None

    if cmd == "initdb":
        from jobsignal.db import get_session_factory
        get_session_factory()
        logging.info("Таблицы созданы (или уже существовали).")

    elif cmd == "seed-channels":
        import yaml
        from jobsignal.db import get_session_factory, Channel
        sf = get_session_factory()
        s = sf()
        with open("config/channels.yaml") as f:
            data = yaml.safe_load(f)
        added = 0
        for ch in data.get("channels", []):
            handle = ch["handle"].lstrip("@")
            if not s.query(Channel).filter_by(handle=handle).first():
                s.add(Channel(handle=handle, title=ch.get("title"), niche=ch.get("niche")))
                added += 1
        s.commit()
        s.close()
        logging.info("Каналов добавлено: %d", added)

    elif cmd == "collect":
        from jobsignal.agents.collector import CollectorAgent as Collector
        from jobsignal.config import get_config; result = Collector(get_config()).run()
        logging.info("сбор: %s", result)

    elif cmd == "posts":
        from jobsignal.db import get_session_factory, RawPost
        from sqlalchemy import desc
        s = get_session_factory()()
        posts = s.query(RawPost).order_by(desc(RawPost.id)).limit(20).all()
        for p in posts:
            print(f"[{p.id}] {(p.text or '')[:120]}")
        s.close()

    elif cmd == "raw-dedup":
        from jobsignal.agents.raw_dedup import RawDeduplicatorAgent
        from jobsignal.config import get_config
        logging.info("дедуп сырых постов: %s",
                     RawDeduplicatorAgent(get_config()).run())

    elif cmd == "parse":
        limit = int(arg) if arg else None
        from jobsignal.agents.parser import Parser
        result = Parser().run(limit=limit)
        logging.info("парсинг: %s", result)

    elif cmd == "dedup":
        from jobsignal.agents.dedup import DeduplicatorAgent as Deduplicator
        from jobsignal.config import get_config; result = Deduplicator(get_config()).run()
        logging.info("дедуп: %s", result)

    elif cmd == "match":
        limit = int(arg) if arg else None
        from jobsignal.agents.matcher import MatcherAgent as Matcher
        from jobsignal.config import get_config; result = Matcher(get_config()).run(limit=limit)
        logging.info("матчинг: %s", result)

    elif cmd == "compose":
        limit = int(arg) if arg else 10
        from jobsignal.agents.composer import Composer
        from jobsignal.db import get_session_factory, Vacancy, VacancyStatus
        from sqlalchemy import desc
        s = get_session_factory()()
        composer = Composer()
        vacancies = (
            s.query(Vacancy)
            .filter(Vacancy.status == VacancyStatus.matched,
                    Vacancy.recruiter_handle.isnot(None))
            .order_by(desc(Vacancy.created_at))
            .limit(limit)
            .all()
        )
        for v in vacancies:
            text = composer.generate(v)
            print(f"\n--- @{v.recruiter_handle} | {v.role} @ {v.company} ---")
            print(text)
        s.close()

    elif cmd == "build-recruiters":
        from jobsignal.agents.recruiter_builder import RecruiterBuilder
        result = RecruiterBuilder().build()
        logging.info("рекрутёры: %s", result)

    elif cmd == "hh-collect":
        from jobsignal.agents.hh_collector import HHCollector
        result = HHCollector().run()
        logging.info("hh сбор: %s", result)

    elif cmd in ("hh-apply", "hh-preview"):
        """Автоотклик на hh.ru.

        hh-preview            — очередь и письма, hh.ru не открывается
        hh-apply              — DRY-RUN: открывает вакансии, ничего не отправляет
        hh-apply --send       — боевой режим (в пределах OUTREACH_PER_HOUR/DAY)
        --limit N             — бюджет прогона: сколько вакансий вообще открыть
        --show                — печатать письма целиком
        """
        import argparse
        ap = argparse.ArgumentParser(prog=f"run.py {cmd}")
        ap.add_argument("--limit", type=int, default=None)
        ap.add_argument("--send", action="store_true",
                        help="реально отправлять отклики (без флага — dry-run)")
        ap.add_argument("--show", action="store_true", help="печатать письма целиком")
        ap.add_argument("--headed", action="store_true", help="показать браузер")
        opts = ap.parse_args(sys.argv[2:])

        from jobsignal.config import get_config
        cfg = get_config()

        if cmd == "hh-preview":
            from jobsignal.agents.hh_apply import preview
            result = preview(cfg, limit=opts.limit)
        else:
            from jobsignal.agents.hh_apply import HHApplyAgent
            result = HHApplyAgent(
                cfg, dry_run=not opts.send, limit=opts.limit,
                headless=not opts.headed,
            ).run()

        rate = result.get("rate", {})
        print()
        print(f"Очередь hh.ru: {result['queue_size']} вакансий")
        print(f"Лимиты: за час {rate.get('sent_hour')}/{rate.get('per_hour')}, "
              f"за сутки {rate.get('sent_day')}/{rate.get('per_day')}, "
              f"можно сейчас: {rate.get('allowed_now')}")
        if "attempt_budget" in result:
            print(f"Бюджет прогона: {result['attempt_budget']} вакансий | "
                  f"режим: {'DRY-RUN' if result['dry_run'] else 'БОЕВОЙ'}")
            print(f"Обработано: {result['attempted']} | отправлено: {result['applied']} | "
                  f"итоги: {result['by_result']}")
        if result.get("stopped_reason"):
            print(f"Остановка: {result['stopped_reason']}")
        if result.get("error"):
            print(f"ОШИБКА: {result['error']}")
        print("-" * 78)
        for it in result["items"]:
            vid = it.get("id", it.get("vacancy_id"))
            print(f"#{str(vid):<5} {it['score']:>3}%  {it['role']} — {it['company']}")
            print(f"       {it['url']}   [{it.get('profile') or '—'}]"
                  f"{'  → ' + it['result'] if it.get('result') else ''}")
            if it.get("detail"):
                print(f"       {it['detail']}")
            letter = it.get("letter") or ""
            if opts.show:
                print("       " + "\n       ".join(letter.splitlines()))
            else:
                first = (letter.splitlines() or [""])[0]
                print(f"       письмо ({len(letter)} симв.): {first[:90]}…")
            print()

    elif cmd == "find-channels":
        """Search tgstat/telemetr for new channels."""
        from jobsignal.agents.channel_finder import ChannelFinder
        queries = sys.argv[2:] if len(sys.argv) > 2 else None
        finder = ChannelFinder(queries=queries)
        result = finder.run()
        logging.info("поиск каналов: %s", result)

    elif cmd == "channels":
        """List active channels."""
        from jobsignal.db import get_session_factory, Channel
        s = get_session_factory()()
        chs = s.query(Channel).filter_by(active=True).all()
        print(f"{'handle':<30} {'title':<30} {'niche':<12} {'source'}")
        print("-" * 80)
        for c in chs:
            print(f"@{c.handle:<29} {(c.title or ''):<30} {(c.niche or ''):<12} {c.source or 'manual'}")
        print(f"\nИтого: {len(chs)} активных каналов")
        s.close()

    elif cmd == "pipeline":
        from jobsignal.orchestrator import Orchestrator
        # Сбор с hh.ru. Телеграм-часть конвейера от него не зависит, поэтому
        # прогон продолжаем — но поломку не проглатываем: она идёт в лог
        # уровнем error и в код возврата. Раньше здесь был warning, и мёртвый
        # сбор («добавлено 0» с конца июня) никто не замечал два с половиной
        # месяца.
        hh_error = None
        try:
            from jobsignal.agents.hh_collector import HHCollector
            HHCollector().run()
        except Exception as exc:
            hh_error = exc
            logging.error("[hh] СБОР С HH.RU СЛОМАН: %s", exc, exc_info=True)
        Orchestrator().run_once()
        if hh_error is not None:
            logging.error("[hh] конвейер отработал, но вакансии с hh.ru не "
                          "собирались: %s", hh_error)
            sys.exit(1)

    elif cmd == "status":
        from jobsignal.db import get_session_factory, Channel, RawPost, Vacancy, Application, Application as Reply
        from sqlalchemy import func
        s = get_session_factory()()
        logging.info("каналов          %d", s.query(func.count(Channel.id)).scalar())
        logging.info("сырых постов     %d", s.query(func.count(RawPost.id)).scalar())
        logging.info("вакансий         %d", s.query(func.count(Vacancy.id)).scalar())
        logging.info("откликов         %d", s.query(func.count(Application.id)).filter(Application.sent_at.isnot(None)).scalar())
        logging.info("ответов          %d", s.query(func.count(Application.id)).filter(Application.replied_at.isnot(None)).scalar())
        s.close()

    elif cmd in ("dashboard", "serve"):
        from jobsignal.dashboard.app import serve
        port = int(os.environ.get("DASHBOARD_PORT", "5000"))
        serve(port=port)

    else:
        print(f"Команды: {', '.join(COMMANDS)}")


if __name__ == "__main__":
    main()
