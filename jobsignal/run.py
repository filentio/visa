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
    "hh-session",             # NEW: жива ли сессия hh.ru и надолго ли
    "channels",               # NEW: list channels
    "pipeline", "status", "dashboard", "serve",
)


HH_LOCK_PATH = "/run/jobsignal-hh.lock"


def _take_hh_lock():
    """Взять замок на обращения к hh.ru.

    Сбор по таймеру и отклик ходят на один адрес, который ddos-guard уже
    держит на прицеле; одновременно им там делать нечего. Тот же файл берёт
    flock(1) в юните конвейера — механизм один, так что замок общий.
    Возвращает открытый файл (держать до конца работы) или None, если занято.
    """
    import fcntl

    fh = open(HH_LOCK_PATH, "w")
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        fh.close()
        return None
    return fh


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
        hh-apply --send --auto — то же по таймеру: свой порог, без присмотра
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
        ap.add_argument("--auto", action="store_true",
                        help="прогон по таймеру: порог HH_AUTO_THRESHOLD и "
                             "только при OUTREACH_MODE=full_auto")
        opts = ap.parse_args(sys.argv[2:])

        from jobsignal.config import get_config, OutreachMode
        cfg = get_config()

        # Автоматический прогон отличается от ручного тремя вещами.
        threshold = None
        max_sends = None
        if opts.auto:
            # Первое: он спрашивает разрешение у OUTREACH_MODE. Иначе вернуть
            # систему под присмотр значило бы не поменять строчку в .env, а
            # вспомнить про таймер и погасить его.
            mode = cfg.settings.outreach_mode
            if mode is not OutreachMode.FULL_AUTO:
                logging.warning(
                    "[hh_apply] OUTREACH_MODE=%s — автоматическая отправка "
                    "выключена, прогон пропущен", mode.value)
                sys.exit(0)
            # Второе: планка выше. За ручной отправкой стоит человек, который
            # посмотрел вакансию, за этой — никто.
            threshold = int(cfg.settings.hh_auto_threshold)
            # Одна отправка за прогон: интервал между откликами задаёт таймер
            # (десять слотов в час), а прогон живёт минуту-две. Иначе паузы
            # 4,5-6,5 минуты держали бы замок /run/jobsignal-hh.lock почти
            # весь час, и конвейер пропускал бы прогоны — вместе с разбором,
            # оценкой и уведомлениями, которым hh вообще не нужен.
            max_sends = int(os.environ.get("HH_MAX_SENDS_PER_RUN", "1"))
            logging.info(
                "[hh_apply] автоматический прогон: порог %d%%; вакансии от "
                "%d%% до %d%% остаются в дашборде и ждут решения руками",
                threshold, cfg.settings.match_threshold, threshold - 1)

        if cmd == "hh-preview":
            from jobsignal.agents.hh_apply import preview
            result = preview(cfg, limit=opts.limit)   # в сеть не ходит, замок не нужен
        else:
            lock = _take_hh_lock()
            if lock is None:
                if opts.auto:
                    # Для таймера это не сбой, а штатное расхождение: конвейер
                    # ещё работает с hh. Своё время отправка возьмёт на
                    # следующем срабатывании, тревогу поднимать незачем.
                    logging.info("[hh_apply] замок %s занят — конвейер ещё "
                                 "работает с hh.ru; пропускаю прогон",
                                 HH_LOCK_PATH)
                    sys.exit(0)
                logging.error("[hh] сейчас идёт другая работа с hh.ru "
                              "(сбор по таймеру или другой отклик) — "
                              "замок %s занят, выходим", HH_LOCK_PATH)
                sys.exit(1)
            from jobsignal.agents.hh_apply import HHApplyAgent
            result = HHApplyAgent(
                cfg, dry_run=not opts.send, limit=opts.limit,
                headless=not opts.headed, threshold=threshold,
                max_sends=max_sends,
            ).run()

        rate = result.get("rate", {})
        print()
        print(f"Очередь hh.ru: {result['queue_size']} вакансий")
        print(f"Лимиты: за час {rate.get('sent_hour')}/{rate.get('per_hour')}, "
              f"за сутки {rate.get('sent_day')}/{rate.get('per_day')}, "
              f"можно сейчас: {rate.get('allowed_now')}")
        if result.get("threshold"):
            print(f"Порог очереди: от {result['threshold']}%")
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

    elif cmd == "hh-session":
        """Состояние сессии hh.ru.

        Без флагов — только чтение файла, в сеть не ходит.
        --probe   — сходить на hh и проверить, пускают ли (поднимает Chromium)
        --notify  — отправить тревогу в телеграм, если сессия плоха
        """
        import argparse
        ap = argparse.ArgumentParser(prog="run.py hh-session")
        ap.add_argument("--probe", action="store_true",
                        help="живая проверка на hh.ru, а не только файл")
        ap.add_argument("--notify", action="store_true",
                        help="отправить тревогу в телеграм, если сессия плоха")
        ap.add_argument("--force-notify", action="store_true",
                        help="отправить, даже если такая тревога уже уходила")
        opts = ap.parse_args(sys.argv[2:])

        from jobsignal import hh_session
        lock = None
        if opts.probe:
            # Живая проба открывает hh.ru — тот же адрес, что сбор и отклик,
            # и тот же замок. Читать файл замок не мешает, поэтому берём его
            # только под пробу.
            lock = _take_hh_lock()
            if lock is None:
                logging.error("[hh] сейчас идёт другая работа с hh.ru — "
                              "замок %s занят; проверяю только файл",
                              HH_LOCK_PATH)
                opts.probe = False
        st = hh_session.status(probe_live=bool(opts.probe))
        print()
        print(hh_session.report(st))
        print()
        if opts.notify or opts.force_notify:
            sent = hh_session.notify(st, force=opts.force_notify)
            print("тревога отправлена" if sent else
                  "тревогу не отправлял (сессия в порядке или уже сообщал)")
        sys.exit(0 if st.level is not hh_session.Level.DEAD else 1)

    elif cmd == "pipeline":
        from jobsignal.orchestrator import Orchestrator

        # Сессия hh.ru живёт до первого разлогина на стороне hh и
        # обновляется только руками. Проверяем её в начале каждого прогона:
        # дешёвая проверка файла — всегда, живая проба — раз в несколько
        # часов. Плохая сессия прогон не останавливает (сбор и телеграм от
        # неё не зависят), но уходит сообщением в телеграм — раньше система
        # просто переставала отправлять отклики, и молча.
        from jobsignal import hh_session
        try:
            hh_session.guard()
        except Exception as exc:  # noqa: BLE001 — присмотр не должен ронять работу
            logging.error("[hh_session] проверка сессии сорвалась: %s",
                          exc, exc_info=True)

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
