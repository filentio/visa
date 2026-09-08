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
    "parse", "dedup", "match", "compose",
    "hh-collect", "find-channels",          # NEW: search tgstat/telemetr
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
        from jobsignal.agents.collector import Collector
        result = Collector().run()
        logging.info("сбор: %s", result)

    elif cmd == "posts":
        from jobsignal.db import get_session_factory, RawPost
        from sqlalchemy import desc
        s = get_session_factory()()
        posts = s.query(RawPost).order_by(desc(RawPost.id)).limit(20).all()
        for p in posts:
            print(f"[{p.id}] {(p.text or '')[:120]}")
        s.close()

    elif cmd == "parse":
        limit = int(arg) if arg else None
        from jobsignal.agents.parser import Parser
        result = Parser().run(limit=limit)
        logging.info("парсинг: %s", result)

    elif cmd == "dedup":
        from jobsignal.agents.dedup import Deduplicator
        result = Deduplicator().run()
        logging.info("дедуп: %s", result)

    elif cmd == "match":
        limit = int(arg) if arg else None
        from jobsignal.agents.matcher import Matcher
        result = Matcher().run(limit=limit)
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

    elif cmd == "notify-check":
        from jobsignal.agents.notify_bot import NotifyBot
        result = NotifyBot().check_and_notify()
        logging.info("уведомления: %s", result)

    elif cmd == "build-recruiters":
        from jobsignal.agents.recruiter_builder import RecruiterBuilder
        result = RecruiterBuilder().build()
        logging.info("рекрутёры: %s", result)

    elif cmd == "hh-collect":
        from jobsignal.agents.hh_collector import HHCollector
        result = HHCollector().run()
        logging.info("hh сбор: %s", result)

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
        from jobsignal.orchestrator import run_pipeline
        # also collect from hh.ru
        try:
            from jobsignal.agents.hh_collector import HHCollector
            HHCollector().run()
        except Exception as e:
            logging.warning("hh collect error: %s", e)
        run_pipeline()

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
