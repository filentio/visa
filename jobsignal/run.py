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
    "hh-resumes",             # NEW: какое резюме hh к какому профилю; пауза автоотклика
    "hh-replies",             # NEW: учёт ответов на отклики hh.ru
    "tg-watch",               # NEW: сторож входящих в телеграме
    "tg-outreach",            # NEW: отклик в телеграм полуавтоматом
    "proxy-traffic",          # NEW: расход трафика прокси; --topup после пополнения
    "tg-scout",               # NEW: поиск новых каналов с вакансиями и надзор за старыми
    "llm-check",              # NEW: отвечает ли модель (основной и запасной провайдер)
    "llm-use",                # NEW: все задачи модели на одного провайдера
    "admin-token",            # NEW: токен служебного доступа к кабинету (Bearer)
    "set-option",             # NEW: правка разрешённой настройки config/.env
    "boards-collect",         # NEW: Хабр Карьера и getmatch (--force — не ждать интервала)
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
        # Упавший или аварийно остановленный прогон — код 1, иначе таймер
        # ломался бы десять раз в час без единой тревоги: OnFailure смотрит
        # только на код возврата, а ошибка оставалась напечатанной строкой.
        if cmd == "hh-apply" and (result.get("error") or result.get("aborted")):
            sys.exit(1)

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

    elif cmd == "hh-resumes":
        """Резюме аккаунта hh и их профили; пауза автоотклика из-за резюме.

        Без флагов — из кэша (если свежий), иначе со страницы hh.
        --refresh — перечитать список резюме на hh (через прокси)
        --unblock — снять паузу после отклика не с тем резюме
        """
        import argparse
        import asyncio
        ap = argparse.ArgumentParser(prog="run.py hh-resumes")
        ap.add_argument("--refresh", action="store_true")
        ap.add_argument("--unblock", action="store_true")
        opts = ap.parse_args(sys.argv[2:])

        from jobsignal import hh_resumes, hh_session, profiles
        if opts.unblock:
            print("пауза снята" if hh_resumes.unblock() else "паузы не было")
            from jobsignal.agents.hh_apply import _resume_problem_clear
            _resume_problem_clear()
            return

        found = None if opts.refresh else hh_resumes.read_cache()
        if not found:
            lock = _take_hh_lock()
            if lock is None:
                print(f"сейчас идёт другая работа с hh.ru (замок {HH_LOCK_PATH}) — "
                      "повтори через пару минут")
                sys.exit(1)

            async def _fetch():
                from playwright.async_api import async_playwright
                path = hh_session.state_path()
                async with async_playwright() as pw:
                    br = await pw.chromium.launch(headless=True,
                                                  proxy=hh_session.playwright_proxy())
                    ctx = await br.new_context(
                        storage_state=str(path) if path.exists() else None,
                        locale="ru-RU", viewport={"width": 1440, "height": 900})
                    ctx.set_default_timeout(60000)
                    await hh_session.block_heavy(ctx, "hh-resumes")
                    page = await ctx.new_page()
                    try:
                        return await hh_resumes.account_resumes(page, use_cache=False)
                    finally:
                        await br.close()

            found = asyncio.run(_fetch())
            from jobsignal import proxy_meter
            proxy_meter.flush()
        print("резюме на hh:")
        for h, t in (found or {}).items():
            print(f"  «{t}»  ({h[:10]}…)")
        if not found:
            print("  не найдены — проверь сессию: run.py hh-session --probe")
        mapping, problems = hh_resumes.match_profiles(found or {})
        print("\nпрофили:")
        for p in profiles.all_profiles():
            m = mapping.get(p.key)
            print(f"  {p.key} «{p.name}» (ищем «{p.hh_resume_title or '—'}») → "
                  + (f"«{m['title']}»" if m else "НЕ СОПОСТАВЛЕНО"))
        for pr in problems:
            print(f"  ⚠ {pr}")
        b = hh_resumes.blocked()
        print("\nавтоотклик: " + (f"НА ПАУЗЕ с {b.get('at')}: {b.get('reason')}\n"
                                  "  снять: run.py hh-resumes --unblock" if b
                                  else "пауз из-за резюме нет"))

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
        # Площадки помимо hh (Хабр Карьера, getmatch). Сбой — в журнал, не в
        # код возврата: это дополнительный источник, конвейер от него не зависит.
        try:
            from jobsignal.agents.boards_collector import BoardsCollector
            BoardsCollector().run()
        except Exception as exc:  # noqa: BLE001
            logging.error("[boards] сбор с площадок сломан: %s", exc, exc_info=True)
        # Отказ провайдера LLM — не «сбой прогона», а остановка всей системы:
        # без разбора и оценки очередь на отклик не пополняется. Ловим отдельно,
        # чтобы в журнале была причина человеческим языком, а в телеграм ушла
        # тревога через OnFailure. 14.09 такой отказ длился 30 часов и выглядел
        # в журнале как обычные сбои разбора.
        failed_agents: list[str] = []
        try:
            from jobsignal.llm import LLMUnavailable
            results = Orchestrator().run_once()
            failed_agents = Orchestrator.failed(results)
        except LLMUnavailable as exc:
            logging.error("[llm] ДОСТУП К МОДЕЛИ ЗАКРЫТ: %s", exc)
            logging.error("[llm] Проверь баланс и ключ в консоли провайдера. "
                          "Разбор и оценка стоят, отклики не пополняются. "
                          "Попытки постам не списаны — прогон догонит сам.")
            sys.exit(1)
        if failed_agents:
            logging.error("[pipeline] упали агенты: %s — прогон доведён до "
                          "конца, но результат неполный", ", ".join(failed_agents))
        if hh_error is not None:
            logging.error("[hh] конвейер отработал, но вакансии с hh.ru не "
                          "собирались: %s", hh_error)
        if hh_error is not None or failed_agents:
            sys.exit(1)

    elif cmd == "hh-replies":
        # Учёт ответов. Отдельной командой, а не внутри pipeline: конвейер
        # ходит на hh каждый час под замком, а список откликов меняется
        # медленно — раз в сутки достаточно, и лишние обращения к сайту с
        # одного адреса ни к чему.
        from jobsignal.agents.hh_replies import HHRepliesAgent, HHRepliesBroken
        try:
            rep = HHRepliesAgent().run()
        except HHRepliesBroken as exc:
            logging.error("[hh_replies] УЧЁТ ОТВЕТОВ СЛОМАН: %s", exc)
            sys.exit(1)
        c = rep["counters"]
        print(f"Откликов на hh: {c.get('all', '?')} | без ответа "
              f"{c.get('awaiting', '?')} | отказ {c.get('discard', '?')} | "
              f"приглашение {c.get('invitation', '?')} | собеседование "
              f"{c.get('interview', '?')} | оффер {c.get('hired', '?')}")
        print(f"Сопоставлено с базой: {rep['matched']} из {rep['seen']} | "
              f"новых ответов записано: {rep['replied_new']}")

    elif cmd == "tg-watch":
        # Сторож входящих: ботов-рекрутёров и контактов из нашей базы.
        # Отдельной командой и своим таймером: к hh не обращается, замок
        # ему не нужен, а частота нужна куда выше — вопрос без ответа
        # закрывает отклик за тринадцать минут.
        from jobsignal.agents.tg_watch import TGWatchAgent, TGWatchBroken
        try:
            rep = TGWatchAgent().run()
        except TGWatchBroken as exc:
            logging.error("[tg_watch] СТОРОЖ НЕ РАБОТАЕТ: %s", exc)
            sys.exit(1)
        print(f"Диалогов под наблюдением: {rep['watched']} | "
              f"новых сообщений: {rep['notified']}")

    elif cmd == "tg-outreach":
        # Готовит пачку откликов в телеграм и присылает в бот. Отправляет
        # человек: автоматическая рассылка в личку — повод для блокировки
        # аккаунта, а терять можно не сервис, а переписку.
        limit = int(arg) if arg else None
        from jobsignal.agents.tg_outreach import TGOutreachAgent
        rep = TGOutreachAgent().run(limit=limit)
        left = rep["queue"] - rep["offered"]
        print(f"Очередь: {rep['queue']} | предложено: {rep['offered']}"
              + (f" | не вышло: {rep['failed']}" if rep.get("failed") else "")
              + (f" | осталось в очереди: {left}" if left > 0 else ""))

    elif cmd == "boards-collect":
        from jobsignal.agents.boards_collector import BoardsCollector
        print(BoardsCollector(force="--force" in sys.argv[2:]).run())

    elif cmd == "set-option":
        """run.py set-option KEY VALUE — правка одной настройки из белого списка.

        Порог автоотклика, источники для оценки, дневной лимит токенов. Значение
        проверяется; прежний .env — рядом (.env.bak-<время>). Таймерные службы
        читают .env при каждом запуске, перезапуск не нужен.
        """
        import re as _re
        from jobsignal import env_store
        allowed = {
            "HH_AUTO_THRESHOLD": r"(6\d|7\d|8\d|9[0-5])",
            "MATCH_THRESHOLD": r"(5\d|6\d|7\d|8\d|9[0-5])",
            "MATCH_SOURCES": r"[a-z]+(,[a-z]+){0,9}",
            "LLM_DAILY_TOKENS": r"\d{1,9}",
            "MATCH_MAX_AGE_DAYS": r"\d{1,3}",
        }
        if len(sys.argv) != 4 or sys.argv[2] not in allowed:
            print("run.py set-option KEY VALUE; KEY из: " + ", ".join(allowed))
            sys.exit(2)
        key, val = sys.argv[2], sys.argv[3].strip()
        if not _re.fullmatch(allowed[key], val):
            print(f"{key}: недопустимое значение {val!r}")
            sys.exit(2)
        old = env_store.read().get(key, "(не задано)")
        env_store.update({key: val})
        print(f"{key}: {old} → {val}")

    elif cmd == "admin-token":
        """Новый токен служебного доступа к кабинету (DASHBOARD_API_TOKEN).

        Печатается один раз. Его вставляют в секрет среды Claude (тип Bearer,
        сайт — адрес кабинета); прежний токен перестаёт действовать.
        --revoke — убрать токен совсем.
        """
        import secrets
        from jobsignal import env_store
        if "--revoke" in sys.argv[2:]:
            env_store.update({"DASHBOARD_API_TOKEN": None})
            print("токен удалён")
        else:
            tok = secrets.token_urlsafe(36)
            env_store.update({"DASHBOARD_API_TOKEN": tok})
            print("Новый токен (виден один раз, в чаты не вставлять):\n")
            print(tok)
            print("\nВставь его в секрет среды Claude: тип Bearer, сайт — адрес кабинета.")
        os.system("systemctl restart jobsignal-dashboard")

    elif cmd == "llm-use":
        """Все задачи модели — одному провайдеру, без запасного.

        run.py llm-use deepseek               — основной DeepSeek, маршруты сброшены
        run.py llm-use deepseek --drop-openai — и стереть ключ OpenAI-совместимого
                                                (Яндекса), чтобы его не вызвало ничто
        """
        import argparse
        ap = argparse.ArgumentParser(prog="run.py llm-use")
        ap.add_argument("provider")
        ap.add_argument("--drop-openai", action="store_true")
        opts = ap.parse_args(sys.argv[2:])
        from jobsignal import env_store, llm
        if opts.provider not in llm.PROVIDERS:
            print(f"провайдер — один из: {', '.join(llm.PROVIDERS)}")
            sys.exit(2)
        if opts.drop_openai and opts.provider == "openai":
            print("--drop-openai стирает ключ того же провайдера — так нельзя")
            sys.exit(2)
        changes = {"LLM_PROVIDER": opts.provider, "LLM_FALLBACK_PROVIDER": None,
                   "LLM_ROUTE_PARSE": None, "LLM_ROUTE_MATCH": None,
                   "LLM_ROUTE_LETTER": None, "LLM_ROUTE_FACTS": None}
        if opts.drop_openai:
            changes.update({"OPENAI_API_KEY": None, "OPENAI_BASE_URL": None,
                            "OPENAI_MODEL": None, "OPENAI_MODEL_SMALL": None,
                            "OPENAI_EXTRA_HEADERS": None})
        env_store.update(changes)
        print(f"все задачи модели → {opts.provider}; запасного нет"
              + ("; ключ OpenAI-совместимого стёрт" if opts.drop_openai else ""))
        print("копия прежнего config/.env — рядом, .env.bak-<время>")
        print("службы перечитают настройки при следующем запуске; бот и кабинет:")
        print("  systemctl restart jobsignal-bot jobsignal-dashboard")

    elif cmd == "llm-check":
        # Отвечают ли обе модели (лёгкая — разбор, сильная — оценка и письма)
        # у основного и запасного провайдера. Ключи на экран не выводятся.
        from jobsignal import llm
        from jobsignal.config import get_config
        st = get_config().settings
        roles = [("основной", st.llm_provider)]
        if st.llm_fallback_provider:
            roles.append(("запасной", st.llm_fallback_provider))
        for task, label in (("parse", "разбор"), ("match", "оценка"),
                            ("letter", "письма"), ("facts", "проверка фактов")):
            roles.append((f"задача «{label}»", llm.route(task)))
        seen, names = set(), []
        for role, name in roles:
            name = (name or "").strip().lower()
            if name and name not in seen:
                seen.add(name)
                names.append((role, name))
        bad = False
        for i, (role, name) in enumerate(names):
            print(f"{role}: {name}")
            for tier, (kind, detail) in llm.check(name.strip().lower()).items():
                label = "лёгкая (разбор)" if tier == "small" else "сильная (оценка, письма)"
                print(f"  {label:26} {'✅' if kind == 'ok' else '❌'} {kind}: {detail}")
                bad = bad or (kind != "ok" and role != "основной")
        sys.exit(1 if bad else 0)

    elif cmd == "tg-scout":
        # Раз в неделю по таймеру. --dry-run: показать, что нашлось и какое
        # решение было бы принято, ничего не записывая и не отключая.
        from jobsignal.agents.channel_scout import ChannelScoutAgent
        rep = ChannelScoutAgent(dry_run="--dry-run" in sys.argv).run()
        aud = rep["audit"]
        print(f"Кандидатов: {rep['candidates']} | проверено: {rep['evaluated']} | "
              f"добавлено: {len(rep['added'])} | ждут решения: {len(rep['pending'])} | "
              f"отказ: {rep['rejected']} | не проверены (сбой): {rep['errors']}")
        if rep["added"]:
            print("  добавлены: " + ", ".join("@" + h for h in rep["added"]))
        if rep["pending"]:
            print("  ждут решения: " + ", ".join("@" + h for h in rep["pending"]))
        if aud["disabled"]:
            print("  отключены: " + ", ".join("@" + h for h, _, _ in aud["disabled"]))
        if rep["telegram_problem"]:
            print("  ВНИМАНИЕ: " + rep["telegram_problem"])
        if rep["dry_run"]:
            print("Это предпросмотр (--dry-run): ничего не записано.")

    elif cmd == "proxy-traffic":
        from jobsignal import proxy_meter
        if arg == "--topup":
            gb = sys.argv[3] if len(sys.argv) > 3 else None
            proxy_meter.topup(float(gb.replace(",", ".")) if gb else None)
            print("Счётчик пакета обнулён, предупреждения снова включены.")
            if gb and not os.environ.get("PROXY_PACKAGE_GB"):
                print(f"Чтобы получать предупреждения, впиши в config/.env: "
                      f"PROXY_PACKAGE_GB={gb}")
        print(proxy_meter.report())

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
