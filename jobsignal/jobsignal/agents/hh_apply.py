"""HH Apply Agent — автоотклик на hh.ru через Playwright.

Ниша агента: contact_type='hh'. Телеграм-путь остаётся за OutreachAgent —
здесь только вакансии, на которые отклик отправляется кнопкой на hh.ru.

Схема работы повторяет остальные агенты конвейера:
  очередь из БД (status matched/drafted) → письмо из Vacancy.draft_text
  → отклик на hh.ru → Application(sent_at) + status='applied'.

Свою генерацию писем агент не делает: если draft_text пуст, письмо создаёт
тот же composer, который дергает кнопка «Сгенерировать» в дашборде.

Лимиты отправки не свои, а общесистемные — jobsignal.ratelimit
(OUTREACH_PER_HOUR / OUTREACH_PER_DAY, те самые «за час 0/3, за сутки 0/15»).

Сессия hh.ru — storage_state Playwright, по умолчанию /root/autoapply/state.json.
Логин и капчу проходит человек один раз (autoapply login); агент только
проверяет, что сессия жива.

Селекторы hh.ru завязаны на data-qa и периодически меняются — все они
собраны в SEL ниже. Если отклик перестал отправляться, чинить нужно там.
Проверка залогиненности намеренно НЕ смотрит на вёрстку меню (она уже
ломалась дважды), а определяется редиректом на /account/login.

Но прежде чем идти в селекторы, стоит посмотреть на итог прогона: часть
вакансий hh просто не отдаёт — отвечает 403 и рисует форму логина вместо
вакансии при живой сессии. Это результат forbidden, а не сломанный
селектор; кнопки на такой странице нет, потому что нет и самой вакансии.
"""
from __future__ import annotations

import asyncio
import enum
import json
import logging
import os
import random
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from sqlalchemy import text as sql_text

from ..db import Vacancy, VacancyStatus, get_session_factory
from ..ratelimit import rate_status
from .base import BaseAgent
from .outreach import detect_contact_type, normalize_hh_link

log = logging.getLogger("jobsignal")

BASE_URL = "https://hh.ru"
# Страница, которую hh отдаёт только залогиненному: гостя редиректит на логин.
LOGIN_PROBE_URL = f"{BASE_URL}/applicant/resumes"
LOGIN_URL_MARKERS = ("/account/login", "/account/signup", "/auth/applicant")

DEFAULT_STATE_PATH = "/root/autoapply/state.json"
# Дефолтный бюджет предпросмотра: сколько вакансий показать, если --limit не задан.
DEFAULT_DRY_LIMIT = 10
# Сколько вакансий открыть на один слот отправки в боевом режиме и сколько
# открыть максимум за прогон — см. _run_async, «бюджет попыток».
DEFAULT_ATTEMPTS_PER_SEND = 5
MAX_LIVE_ATTEMPTS = 40
# Сколько раз просить composer переписать письмо, если оно вышло непригодным.
LETTER_TRIES = 2
# Сколько прогонов терпеть 403 по одной вакансии, прежде чем убрать её из
# очереди: страница может быть закрыта временно, но чаще — навсегда.
FORBIDDEN_GIVE_UP = 3
# Сколько 403 подряд без единой прочитанной страницы считать общим блоком hh,
# а не свойством вакансий, и останавливать прогон.
FORBIDDEN_ABORT = 5
# Сколько прогонов терпеть отсутствие кнопки «Откликнуться». В отличие от
# needs_manual это не приговор: кнопки может не быть из-за недогруженной
# вёрстки или подвисшего рендера, и на следующем прогоне она появляется.
# Но если её нет столько раз подряд — вакансия закрыта для откликов, и дальше
# она только жжёт бюджет попыток.
NO_RESPOND_GIVE_UP = 3

SEL = {
    "already_applied": (
        '[data-qa="vacancy-response-link-view-topic"], '
        '[data-qa="vacancy-response-success"], '
        # Появляется сразу после отправки: «приложите сопроводительное к
        # отклику». Признак того, что отклик уже ушёл — без письма.
        '[data-qa="responded-success-attach-cover-letter"]'
    ),
    # Только селекторы страницы вакансии. vacancy-serp__vacancy_response,
    # который тащился из autoapply, живёт в блоке «похожие вакансии» сбоку —
    # с ним отклик мог уйти на ЧУЖУЮ вакансию.
    "respond_button": (
        '[data-qa="vacancy-response-link-top"], '
        '[data-qa="vacancy-response-link-bottom"]'
    ),
    # Модалка «Вы откликаетесь из другой страны» (сервер европейский, поэтому
    # она встаёт почти на каждой вакансии). data-qa был вымышленным: на живой
    # странице кнопка — relocation-warning-confirm. Из-за этого подтверждение
    # не нажималось, модалка перекрывала форму, поле письма не находилось и
    # прогон уходил в no_letter_field, то есть отклик не отправлялся вообще.
    # Старый селектор оставлен запасным на случай возврата прежней вёрстки.
    "relocation_confirm": (
        '[data-qa="relocation-warning-confirm"], '
        '[data-qa="vacancy-response-popup-relocation-warning-confirmation"]'
    ),
    # Ссылка на странице успешного отклика: hh сам предлагает приложить письмо
    # к уже созданному отклику. Тот же data-qa участвует в already_applied как
    # признак «отклик ушёл» — здесь он нужен как кнопка.
    "attach_letter": '[data-qa="responded-success-attach-cover-letter"]',
    "letter_toggle": (
        '[data-qa="vacancy-response-letter-toggle"], '
        '[data-qa="add-cover-letter"]'
    ),
    "letter_input": (
        '[data-qa="vacancy-response-popup-form-letter-input"], '
        '[data-qa="vacancy-response-letter-informer-textarea"], '
        'textarea[name="letter"]'
    ),
    # Страховка от переименования data-qa у поля письма. На странице отклика
    # до клика по letter_toggle нет НИ ОДНОЙ textarea (проверено на живой
    # странице), поэтому первая появившаяся после раскрытия — это письмо.
    # Использовать только ПОСЛЕ клика по переключателю.
    "letter_input_any": "textarea",
    "submit": (
        '[data-qa="vacancy-response-submit-popup"], '
        '[data-qa="vacancy-response-letter-submit"], '
        '[data-qa="vacancy-response-submit"]'
    ),
    "questions_form": '[data-qa="task-body"], [data-qa="vacancy-response-questions"]',
    # hh умеет отдать страницу вакансии БЕЗ редиректа, подменив содержимое
    # формой логина: url остаётся /vacancy/NNN, а вакансии на странице нет.
    "login_wall": (
        '[data-qa="account-login-form"], '
        '[data-qa="applicant-login-card"], '
        '[data-qa="applicant-login-input-email"]'
    ),
    "vac_title": '[data-qa="vacancy-title"]',
    "vac_archived": (
        '[data-qa="vacancy-title-archived-text"], '
        '[data-qa="vacancy-archive-description"], '
        '[data-qa="vacancy-archive-message"], '
        '[data-qa="vacancy-archived"]'
    ),
}

# Текстовый признак архива — страховка, если data-qa снова переименуют.
ARCHIVED_TEXT_RE = re.compile(r"в\s+архиве|vacancy\s+is\s+archived", re.IGNORECASE)

# hh прямо пишет, почему не отдал вакансию: она видна только приглашённым
# соискателям и работодателю. Аккаунт при этом залогинен — это не сессия и не
# селекторы, отклик по такой вакансии невозможен в принципе.
# Страницу логина hh отдаёт то по-русски, то по-английски — держим оба текста.
NOT_ALLOWED_TEXT_RE = re.compile(
    r"недоступна\s+при\s+текущей\s+авторизации"
    r"|not\s+available\s+under\s+current\s+authorization",
    re.IGNORECASE,
)

PROFILE_LABELS = ("Senior AI PM", "CPO / Head of Product", "Senior PM/PO")

# Минимальная длина осмысленного письма. Всё короче — заглушка или обрыв.
MIN_LETTER_CHARS = 200
SENTENCE_END = ".!?…»)\""

# Хвост письма — подпись, а не предложение: точки в конце там не бывает.
_FAREWELL_RE = re.compile(
    r"^(?:с уважением|всего доброго|с наилучшими пожеланиями"
    r"|best regards|kind regards|regards|sincerely)[,.!]?$",
    re.IGNORECASE,
)
# Один контакт: телефон, почта, ссылка или хендл. Необязательную метку
# («тел.:», «Telegram:») отбрасываем перед сверкой.
_CONTACT_RE = re.compile(
    r"^(?:тел(?:\.|ефон)?|моб(?:\.|ильный)?|phone|e-?mail|почта|tg|telegram)?"
    r"\s*:?\s*"
    r"(?:\+?[\d\s()\-]{7,}"
    r"|[^\s@]+@[^\s@]+\.[a-z]{2,}"
    r"|(?:https?://|t\.me/|@)\S+)$",
    re.IGNORECASE,
)
# ЗАЩИТА ОТ ДЕФЕКТА №3: контакты часто идут одной строкой через разделитель
# («+7 926 211-10-88 / @handle»). Целиком такая строка ни под телефон, ни под
# хендл не подходила, не отрезалась как подпись — и готовое письмо браковалось
# как «обрывается на полуслове».
# Слэш режем только с пробелом рядом: внутри «t.me/handle» он часть контакта.
_CONTACT_SPLIT_RE = re.compile(r"[,;|·•]+|\s+/+\s*|\s*/+\s+")
# ФИО в подписи — только с заглавных, иначе под шаблон попадёт обычная проза.
_SIGN_NAME_RE = re.compile(r"^(?:[А-ЯЁA-Z][\w'’\-]*\s*){1,4}$")


def _is_contact_line(line: str) -> bool:
    """Строка целиком из контактов — возможно нескольких через разделитель."""
    parts = [p.strip() for p in _CONTACT_SPLIT_RE.split(line)]
    parts = [p for p in parts if p]
    return bool(parts) and all(_CONTACT_RE.match(p) for p in parts)


def _strip_signature(text: str) -> str:
    """Убирает хвостовые строки подписи, чтобы проверять конец самого письма."""
    lines = text.splitlines()
    while lines:
        last = lines[-1].strip()
        if (
            not last
            or _FAREWELL_RE.match(last)
            or _is_contact_line(last)
            or _SIGN_NAME_RE.match(last)
        ):
            lines.pop()
            continue
        break
    return "\n".join(lines).strip()


def letter_problem(text: str) -> str | None:
    """Причина, по которой письмо нельзя отправлять, или None.

    Генератор упирается в max_tokens и обрывает текст на полуслове. Отправить
    обрубок хуже, чем не отправить, поэтому проверяем перед откликом.
    Подпись («С уважением / Павел / +7…») терминальной точки не имеет —
    её отрезаем, иначе нормальные письма отбраковываются как обрезанные.
    """
    text = (text or "").strip()
    if not text:
        return "письмо пустое"
    if len(text) < MIN_LETTER_CHARS:
        return f"письмо слишком короткое ({len(text)} симв.) — похоже на заглушку"
    body = _strip_signature(text)
    if not body:
        return "в письме только подпись"
    if body[-1] not in SENTENCE_END:
        return "письмо обрывается на полуслове (упёрлось в лимит токенов)"
    return None


class Result(str, enum.Enum):
    """Итог по одной вакансии."""

    APPLIED = "applied"                    # отклик отправлен с письмом
    APPLIED_NO_LETTER = "applied_no_letter"  # отклик ушёл на клике, без письма
    DRY_RUN = "dry_run"                    # проверено, отправка не выполнялась
    ALREADY_APPLIED = "already_applied"    # на hh отклик уже есть
    NO_RESPOND_BUTTON = "no_respond_button"
    FORBIDDEN = "forbidden"                # hh не отдал страницу (403 + логин)
    NOT_ALLOWED = "not_allowed"            # вакансия закрыта для этого аккаунта
    GONE = "gone"                          # hh ответил 404
    NEEDS_MANUAL = "needs_manual"          # анкета работодателя — руками
    ARCHIVED = "archived"                  # вакансия снята
    NO_LETTER = "no_letter"                # письмо не удалось получить
    NO_LETTER_FIELD = "no_letter_field"
    NO_SUBMIT = "no_submit"
    UNCONFIRMED = "unconfirmed"            # submit нажат, подтверждения нет
    FAILED = "failed"

    @property
    def sent(self) -> bool:
        """Отклик на hh существует по нашей вине — списывает квоту.

        ЗАЩИТА ОТ ДЕФЕКТА №8: у вакансий без анкеты hh отправляет отклик уже
        на клике «Откликнуться», ДО того как есть куда вписать письмо (письмо
        предлагается приложить потом). Такой отклик код раньше не замечал:
        поля письма нет — значит no_letter_field, «ничего не отправлено».
        Отправка уходила молча, в applications не попадала и квоту не
        списывала, то есть OUTREACH_PER_HOUR можно было превысить, а
        работодатель получал пустой отклик.
        """
        return self in (Result.APPLIED, Result.APPLIED_NO_LETTER)

    @property
    def is_success(self) -> bool:
        return self.sent

    @property
    def closes_vacancy(self) -> bool:
        """Вакансию можно перевести в applied: отклик есть (наш или ранее).

        applied_no_letter тоже сюда: отклик на hh существует, повторно
        открывать вакансию незачем — письмо дописывается в самом отклике.
        """
        return self in (Result.APPLIED, Result.APPLIED_NO_LETTER,
                        Result.ALREADY_APPLIED)

    @property
    def drops_vacancy(self) -> bool:
        """Отклик невозможен и это не изменится — из очереди можно убрать.

        NEEDS_MANUAL сюда же: анкета работодателя — свойство самой вакансии,
        повтором оно не лечится. Без этого вакансия оставалась drafted и
        занимала слот в КАЖДОМ прогоне: #1096 успела съесть 7 попыток, #5426 — 5,
        и обе так и не могли быть отправлены. Вакансия не теряется — статус
        skipped прячет её только из фильтра «в работе», причина лежит в
        hh_apply_attempts, откликнуться руками по-прежнему можно.
        """
        return (self in (Result.ARCHIVED, Result.GONE, Result.NOT_ALLOWED,
                         Result.NEEDS_MANUAL))


@dataclass
class Attempt:
    vacancy_id: int
    role: str
    company: str
    url: str
    score: int
    profile: str
    letter: str
    result: Result
    detail: str = ""
    screenshot: str | None = None

    def as_dict(self) -> dict:
        return {
            "vacancy_id": self.vacancy_id,
            "role": self.role,
            "company": self.company,
            "url": self.url,
            "score": self.score,
            "profile": self.profile,
            "letter": self.letter,
            "result": self.result.value,
            "detail": self.detail,
            "screenshot": self.screenshot,
        }


@dataclass
class RunReport:
    dry_run: bool
    queue_size: int = 0
    attempt_budget: int = 0
    send_budget: int = 0
    attempts: list[Attempt] = field(default_factory=list)
    stopped_reason: str = ""
    error: str = ""

    @property
    def applied(self) -> int:
        return sum(1 for a in self.attempts if a.result.is_success)

    def as_dict(self) -> dict:
        by_result: dict[str, int] = {}
        for a in self.attempts:
            by_result[a.result.value] = by_result.get(a.result.value, 0) + 1
        return {
            "agent": HHApplyAgent.name,
            "dry_run": self.dry_run,
            "queue_size": self.queue_size,
            "attempt_budget": self.attempt_budget,
            "send_budget": self.send_budget,
            "attempted": len(self.attempts),
            "applied": self.applied,
            "by_result": by_result,
            "stopped_reason": self.stopped_reason,
            "error": self.error,
            "rate": rate_status(),
            "items": [a.as_dict() for a in self.attempts],
            "finished_at": datetime.now(timezone.utc).isoformat(),
        }


# ── очередь и письма ─────────────────────────────────────────────────────────

def _best_score(v: Vacancy) -> tuple[int, str]:
    best, label = 0, ""
    for m in (v.match_scores or []):
        if (m.score or 0) > best:
            best, label = m.score, m.profile_key or ""
    return best, label


def _hh_url(v: Vacancy) -> str | None:
    link = v.link or ""
    if not re.search(r"hh\.ru/(vacancy/\d+|applicant/vacancy)", link):
        return None
    return normalize_hh_link(link)


def _hh_key(url: str) -> str:
    """Номер вакансии на hh — по нему очередь дедуплицируется.

    Одну и ту же вакансию приносят разные каналы, и в БД она лежит несколькими
    строками с разными id. Ключом дедупа id быть не может.
    """
    m = re.search(r"/vacancy/(\d+)", url or "")
    return m.group(1) if m else ""


class HHApplyAgent(BaseAgent):
    name = "hh_apply"

    def __init__(
        self,
        config,
        dry_run: bool = True,
        limit: int | None = None,
        headless: bool | None = None,
        state_path: str | None = None,
    ) -> None:
        super().__init__(config)
        self.dry_run = dry_run
        self.limit = limit
        self.headless = (
            headless
            if headless is not None
            else os.environ.get("HH_HEADLESS", "1") not in ("0", "false", "no")
        )
        self.state_path = Path(
            state_path or os.environ.get("HH_STATE_PATH", DEFAULT_STATE_PATH)
        )
        self.delay = (
            float(os.environ.get("HH_DELAY_MIN", "12")),
            float(os.environ.get("HH_DELAY_MAX", "35")),
        )
        # Пауза между просто открытыми страницами (архив, дубликат, нет
        # кнопки): отправки не было, ждать 12-35 секунд незачем, но и грузить
        # десятки вакансий подряд без паузы hh не любит.
        self.scan_delay = (
            float(os.environ.get("HH_SCAN_DELAY_MIN", "1.5")),
            float(os.environ.get("HH_SCAN_DELAY_MAX", "4")),
        )
        self.attempts_per_send = int(
            os.environ.get("HH_ATTEMPTS_PER_SEND", str(DEFAULT_ATTEMPTS_PER_SEND))
        )
        self.nav_timeout_ms = int(os.environ.get("HH_NAV_TIMEOUT_MS", "30000"))
        self.screenshots_dir = Path(
            os.environ.get("HH_SCREENSHOTS_DIR", "data/hh_screenshots")
        )
        self._Session = get_session_factory()
        # Ставится после успешной проверки логина; без него state.json не пишем.
        self._session_ok = False
        # Прочитана ли за прогон хотя бы одна страница вакансии. Отличает
        # «эту вакансию hh не отдаёт» от «hh не отдаёт ничего»: во втором
        # случае вычищать очередь по 403 нельзя.
        self._page_read = False

    # --- публичный вход ---------------------------------------------------

    def run(self) -> dict:
        report = asyncio.run(self._run_async())
        payload = report.as_dict()
        self._dump(payload)
        log.info(
            "[hh_apply] %s | очередь %d | бюджет попыток %d | отправок %d | итоги %s",
            "DRY-RUN (без отправки)" if self.dry_run else "БОЕВОЙ РЕЖИМ",
            payload["queue_size"], payload["attempt_budget"],
            payload["applied"], payload["by_result"],
        )
        return payload

    # --- очередь ----------------------------------------------------------

    def _queue(self, session) -> list[dict]:
        """Вакансии hh.ru, готовые к отклику: matched/drafted, живая ссылка,
        нет уже отправленного отклика по этому же каналу."""
        threshold = int(getattr(self.config.settings, "match_threshold", 0) or 0)
        rows = (
            session.query(Vacancy)
            .filter(
                Vacancy.is_primary.is_(True),
                Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted]),
            )
            .all()
        )
        sent_ids = {
            r[0]
            for r in session.execute(
                sql_text(
                    "SELECT DISTINCT vacancy_id FROM applications "
                    "WHERE sent_at IS NOT NULL"
                )
            )
        }
        # ЗАЩИТА ОТ ДЕФЕКТА №5: sent_ids отсекает только ту строку, из которой
        # отклик ушёл. Дубликат той же вакансии hh (другая строка, другой id)
        # оставался в очереди и на прогоне неизбежно получал already_applied —
        # впустую открытая страница и съеденный слот. Отсекаем по номеру
        # вакансии на hh: и уже отправленные, и закрытые статусом applied.
        done_keys = {
            _hh_key(r[0])
            for r in session.execute(
                sql_text(
                    "SELECT DISTINCT v.link FROM vacancies v "
                    "LEFT JOIN applications a "
                    "  ON a.vacancy_id = v.id AND a.sent_at IS NOT NULL "
                    "WHERE v.status = 'applied' OR a.id IS NOT NULL"
                )
            )
        }
        done_keys.discard("")

        queue: list[dict] = []
        for v in rows:
            ct = v.contact_type or detect_contact_type(v)
            if ct != "hh":
                continue
            url = _hh_url(v)
            if not url:
                log.debug("[hh_apply] #%s: ссылка не похожа на вакансию hh — пропуск", v.id)
                continue
            if v.id in sent_ids:
                continue
            key = _hh_key(url)
            if key and key in done_keys:
                log.debug("[hh_apply] #%s: вакансия %s уже закрыта — пропуск", v.id, key)
                continue
            score, profile = _best_score(v)
            if score < threshold:
                continue
            queue.append(
                {
                    "id": v.id,
                    "key": key,
                    "url": url,
                    "role": v.role or "—",
                    "company": v.company or "—",
                    "score": score,
                    "profile": profile,
                    "draft": (v.draft_text or "").strip(),
                }
            )
        # Внутри одного балла — свежие вперёд (id по возрастанию = от самых
        # старых): в прогоне 09.09 половина очереди была снятыми вакансиями
        # многомесячной давности, и они выели весь бюджет попыток.
        queue.sort(key=lambda x: (-x["score"], -x["id"]))

        # Дедуп: из нескольких строк одной вакансии hh остаётся первая по
        # этому порядку, то есть самая свежая с лучшим баллом.
        unique: list[dict] = []
        seen: set[str] = set()
        for item in queue:
            if item["key"]:
                if item["key"] in seen:
                    log.debug("[hh_apply] #%s: дубликат вакансии %s — пропуск",
                              item["id"], item["key"])
                    continue
                seen.add(item["key"])
            unique.append(item)
        return unique

    # --- письмо -----------------------------------------------------------

    def _letter(self, session, item: dict) -> str:
        """Письмо из Vacancy.draft_text; если пусто или непригодно — генерим тем
        же composer'ом, что и кнопка в дашборде, и сохраняем в draft_text.

        ЗАЩИТА ОТ ДЕФЕКТА №4. Раньше draft_text брался как есть. Composer при
        сбое модели молча возвращает заглушку на две строки («Прилагаю резюме —
        готов обсудить»); она попадала в draft_text и после этого вакансия
        браковалась НАВСЕГДА: письмо непригодно, а перегенерации не было.
        Поэтому непригодный черновик считаем отсутствующим, генерацию повторяем
        LETTER_TRIES раз и сохраняем результат только если он прошёл проверку.
        """
        draft = item["draft"]
        if draft and letter_problem(draft) is None:
            return draft

        # Единственная точка генерации в проекте — общая с дашбордом.
        from .notify_bot import _cover_text

        v = session.get(Vacancy, item["id"])
        profile = item["profile"] if item["profile"] in PROFILE_LABELS else "Senior PM/PO"
        text = ""
        for try_no in range(1, LETTER_TRIES + 1):
            text = (_cover_text(
                role=v.role or "",
                company=v.company or "",
                profile_key=profile,
                recruiter_name="",
            ) or "").strip()
            problem = letter_problem(text)
            if problem is None:
                break
            log.warning(
                "[hh_apply] #%s: письмо не годится (попытка %d/%d): %s",
                v.id, try_no, LETTER_TRIES, problem,
            )
        else:
            # Непригодный текст в draft_text не пишем — иначе следующий прогон
            # снова упрётся в него вместо новой генерации.
            return text or draft

        v.draft_text = text
        if v.status == VacancyStatus.matched:
            v.status = VacancyStatus.drafted
        session.commit()
        item["draft"] = text
        log.info("[hh_apply] #%s: письмо сгенерировано (%d симв.)", v.id, len(text))
        return text

    # --- запись результата ------------------------------------------------

    def _ensure_log_table(self, session) -> None:
        session.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS hh_apply_attempts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vacancy_id INTEGER NOT NULL,
                result VARCHAR(32) NOT NULL,
                detail TEXT,
                screenshot VARCHAR(512),
                dry_run INTEGER NOT NULL DEFAULT 0,
                created_at DATETIME NOT NULL
            )
        """))
        session.execute(sql_text(
            "CREATE INDEX IF NOT EXISTS ix_hh_apply_attempts_vacancy "
            "ON hh_apply_attempts(vacancy_id)"
        ))
        session.commit()

    def _record(self, session, attempt: Attempt) -> None:
        """Одна транзакция: журнал попытки + (при успехе) Application и статус.

        Неудача вакансию не теряет: статус остаётся matched/drafted, письмо уже
        лежит в draft_text, причина — в hh_apply_attempts. Прогон можно повторить.
        """
        now = datetime.now(timezone.utc)
        session.execute(
            sql_text(
                "INSERT INTO hh_apply_attempts "
                "(vacancy_id, result, detail, screenshot, dry_run, created_at) "
                "VALUES (:vid, :res, :detail, :shot, :dry, :now)"
            ),
            {
                "vid": attempt.vacancy_id, "res": attempt.result.value,
                "detail": attempt.detail, "shot": attempt.screenshot,
                "dry": 1 if self.dry_run else 0, "now": now,
            },
        )

        if attempt.result.sent:
            # Письмо пишем только если оно действительно ушло: у
            # APPLIED_NO_LETTER отклик создан кликом, текста не было — пустая
            # строка тут честная, канал 'hh' отличает её от ручной отметки.
            # Вторую отправленную строку по той же вакансии не создаём: очередь
            # такие вакансии уже отсекает, но повтор внутри прогона удвоил бы
            # списание квоты. То же условие закреплено уникальным индексом
            # ux_applications_sent_vacancy.
            already = session.execute(
                sql_text(
                    "SELECT 1 FROM applications "
                    "WHERE vacancy_id = :vid AND sent_at IS NOT NULL"
                ),
                {"vid": attempt.vacancy_id},
            ).first()
            if not already:
                session.execute(
                    sql_text(
                        "INSERT INTO applications "
                        "(vacancy_id, message_text, channel, is_draft, sent_at, created_at) "
                        "VALUES (:vid, :text, 'hh', 0, :now, :now)"
                    ),
                    {
                        "vid": attempt.vacancy_id,
                        "text": (attempt.letter
                                 if attempt.result is Result.APPLIED else ""),
                        "now": now,
                    },
                )
            elif attempt.result is Result.APPLIED and attempt.letter:
                # Письмо дописано в отклик, который раньше ушёл пустым
                # (applied_no_letter). Второй строки быть не может — её
                # запрещает ux_applications_sent_vacancy, — поэтому
                # заполняем текст в существующей, но только если он пуст:
                # затирать уже отправленное письмо нечем и незачем.
                updated = session.execute(
                    sql_text(
                        "UPDATE applications SET message_text = :text "
                        "WHERE vacancy_id = :vid AND sent_at IS NOT NULL "
                        "AND (message_text IS NULL OR message_text = '')"
                    ),
                    {"vid": attempt.vacancy_id, "text": attempt.letter},
                ).rowcount
                if updated:
                    log.info("[hh_apply] вакансия #%s: письмо дописано в "
                             "существующий отклик", attempt.vacancy_id)
            else:
                log.warning(
                    "[hh_apply] вакансия #%s: отправленный отклик уже есть, "
                    "вторую строку не пишем", attempt.vacancy_id,
                )

        if attempt.result.closes_vacancy:
            self._set_status(session, attempt, VacancyStatus.applied)
        elif attempt.result.drops_vacancy:
            self._set_status(session, attempt, VacancyStatus.skipped)
        elif attempt.result is Result.FORBIDDEN:
            # Закрытая страница не лечится повтором: если hh отдаёт 403 уже
            # FORBIDDEN_GIVE_UP прогонов подряд, вакансия только тратит бюджет.
            # _page_read обязателен: пока за прогон не прочитано ни одной
            # страницы, 403 говорит о блоке hh целиком, а не об этой вакансии.
            seen = self._seen_with_result(session, attempt.vacancy_id, Result.FORBIDDEN)
            if seen >= FORBIDDEN_GIVE_UP and self._page_read:
                self._set_status(session, attempt, VacancyStatus.skipped)
                log.warning(
                    "[hh_apply] #%s: hh закрыл страницу %d раз — убираю из очереди",
                    attempt.vacancy_id, seen,
                )
        elif attempt.result is Result.NO_RESPOND_BUTTON:
            # Тот же ограниченный ретрай, что и для 403, но без оглядки на
            # _page_read: этот исход сам по себе означает, что страница
            # прочитана — кнопку искали именно на ней.
            seen = self._seen_with_result(
                session, attempt.vacancy_id, Result.NO_RESPOND_BUTTON)
            if seen >= NO_RESPOND_GIVE_UP:
                self._set_status(session, attempt, VacancyStatus.skipped)
                log.warning(
                    "[hh_apply] #%s: кнопки «Откликнуться» нет %d раз — "
                    "убираю из очереди", attempt.vacancy_id, seen,
                )

        session.commit()

    def _seen_with_result(self, session, vacancy_id: int, result: Result) -> int:
        """Сколько раз эта вакансия уже получала такой исход, включая текущий:
        _record пишет попытку в журнал до того, как считает её здесь."""
        return session.execute(
            sql_text(
                "SELECT count(*) FROM hh_apply_attempts "
                "WHERE vacancy_id = :vid AND result = :res"
            ),
            {"vid": vacancy_id, "res": result.value},
        ).scalar() or 0

    def _set_status(self, session, attempt: Attempt, status: VacancyStatus) -> None:
        """Статус вакансии — сразу всем строкам БД про эту же вакансию hh.

        Иначе закрытие одной строки просто пропускает в очередь её дубликат:
        дедуп прогона оставляет по вакансии одну строку, и на следующем
        прогоне вместо закрытой поднимается вторая — с тем же исходом.
        """
        v = session.get(Vacancy, attempt.vacancy_id)
        if v is not None:
            v.status = status

        key = _hh_key(attempt.url)
        if not key:
            return
        twins = (
            session.query(Vacancy)
            .filter(
                Vacancy.id != attempt.vacancy_id,
                Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted]),
                # LIKE только чтобы сузить выборку: 136098514 совпадёт и с
                # 1360985141, поэтому номер сверяем ещё раз точным разбором.
                Vacancy.link.like(f"%/vacancy/{key}%"),
            )
            .all()
        )
        for twin in twins:
            if _hh_key(twin.link or "") == key:
                twin.status = status
                log.info("[hh_apply] #%s: та же вакансия %s — статус %s",
                         twin.id, key, status.value)

    def _dump(self, payload: dict) -> None:
        try:
            out = Path("data/hh_apply_last.json")
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as exc:  # noqa: BLE001 — отчёт диагностический
            log.warning("[hh_apply] не удалось сохранить отчёт: %s", exc)

    # --- основной цикл ----------------------------------------------------

    async def _run_async(self) -> RunReport:
        report = RunReport(dry_run=self.dry_run)
        session = self._Session()
        try:
            self._ensure_log_table(session)
            queue = self._queue(session)
            report.queue_size = len(queue)

            rate = rate_status(session)
            report.send_budget = 0 if self.dry_run else rate["allowed_now"]

            # ── ЗАЩИТА ОТ ДЕФЕКТА №1 ──────────────────────────────────────
            # Два независимых счётчика:
            #   attempts_left — сколько вакансий вообще откроем. Списывается
            #                   на КАЖДОЙ открытой вакансии, включая ошибки.
            #   sends_left    — сколько откликов реально отправим. Списывается
            #                   только на подтверждённой отправке и сверяется
            #                   с БД после каждой.
            # В autoapply бюджет списывался только за applied/dry_run
            # (ApplyOutcome.consumed_quota), поэтому вакансии с ошибкой
            # прокручивали весь список: бюджет 3 → обработано 40.
            # ЗАЩИТА ОТ ДЕФЕКТА №6: в боевом режиме бюджет попыток равнялся
            # бюджету отправок (3), а списывается он на КАЖДОЙ открытой
            # вакансии. В прогоне 09.09 годными были 4 вакансии из 71 — три
            # архивные страницы съедали весь прогон, и не уходило ни одного
            # отклика. Попыток даём с запасом (attempts_per_send на слот),
            # реальные отправки по-прежнему ограничены sends_left.
            if self.limit is not None:
                attempts_left = max(0, self.limit)
            elif self.dry_run:
                attempts_left = DEFAULT_DRY_LIMIT
            else:
                attempts_left = min(
                    len(queue),
                    MAX_LIVE_ATTEMPTS,
                    max(report.send_budget * self.attempts_per_send,
                        report.send_budget),
                )
            report.attempt_budget = attempts_left
            sends_left = report.send_budget
            forbidden_streak = 0

            if not queue:
                report.stopped_reason = "очередь пуста"
                return report
            if attempts_left <= 0:
                report.stopped_reason = "бюджет прогона 0"
                return report
            if not self.dry_run and sends_left <= 0:
                report.stopped_reason = (
                    f"лимит откликов исчерпан: за час {rate['sent_hour']}/{rate['per_hour']}, "
                    f"за сутки {rate['sent_day']}/{rate['per_day']}"
                    + (f", следующий слот ~{rate['next_slot']}" if rate["next_slot"] else "")
                )
                log.warning("[hh_apply] %s", report.stopped_reason)
                return report

            async with self._browser() as page:
                if not await self._logged_in(page):
                    report.error = (
                        f"сессия hh.ru недействительна ({self.state_path}). "
                        "На сервере нет дисплея, поэтому логин проходится на "
                        "машине с экраном — /root/autoapply/tools/hh_login_local.py "
                        "--upload — и state.json копируется сюда. Проверка: "
                        "cd /root/autoapply && .venv/bin/python -m autoapply.cli "
                        "login --check -c config.yaml"
                    )
                    log.error("[hh_apply] %s", report.error)
                    return report

                for item in queue:
                    if attempts_left <= 0:
                        report.stopped_reason = "бюджет прогона исчерпан"
                        break
                    if not self.dry_run and sends_left <= 0:
                        report.stopped_reason = "лимит откликов исчерпан"
                        break

                    # Вакансию открываем — попытка списывается независимо
                    # от того, чем она закончится.
                    attempts_left -= 1

                    # Письмо генерим ЛЕНИВО, уже после проверок страницы:
                    # больше половины очереди — снятые вакансии, и письмо для
                    # них было бы выброшенным вызовом модели.
                    attempt = await self._process(
                        page, item, lambda: self._letter(session, item)
                    )

                    report.attempts.append(attempt)
                    self._record(session, attempt)

                    if attempt.result is Result.FORBIDDEN and not self._page_read:
                        forbidden_streak += 1
                        if forbidden_streak >= FORBIDDEN_ABORT:
                            report.stopped_reason = (
                                f"hh ответил 403 на {forbidden_streak} вакансий подряд и "
                                "ни одной страницы не отдал — похоже на общий блок, "
                                "а не на закрытые вакансии; прогон остановлен"
                            )
                            log.error("[hh_apply] %s", report.stopped_reason)
                            break
                    else:
                        forbidden_streak = 0

                    log.info(
                        "[hh_apply] %s #%s %s — %s | %s%s",
                        "✓" if attempt.result.is_success else "×",
                        item["id"], item["role"], item["company"],
                        attempt.result.value,
                        f" — {attempt.detail}" if attempt.detail else "",
                    )

                    if attempt.result.is_success:
                        # Квоту пересверяем по БД: она общая с tg-откликами.
                        fresh = rate_status(session)
                        sends_left = min(sends_left - 1, fresh["allowed_now"])
                        if sends_left > 0 and attempts_left > 0:
                            await self._sleep()
                    elif attempts_left > 0:
                        await self._sleep(self.scan_delay)

            if not report.stopped_reason:
                report.stopped_reason = "очередь пройдена"
            return report

        except Exception as exc:  # noqa: BLE001 — отчёт должен вернуться всегда
            log.exception("[hh_apply] прогон упал")
            report.error = f"{type(exc).__name__}: {exc}"
            return report
        finally:
            session.close()

    def _attempt(self, item: dict, letter: str, result: Result,
                 detail: str = "", screenshot: str | None = None) -> Attempt:
        return Attempt(
            vacancy_id=item["id"], role=item["role"], company=item["company"],
            url=item["url"], score=item["score"], profile=item["profile"],
            letter=letter, result=result, detail=detail, screenshot=screenshot,
        )

    async def _sleep(self, delay: tuple[float, float] | None = None) -> None:
        low, high = delay or self.delay
        await asyncio.sleep(random.uniform(low, high))

    # --- браузер ----------------------------------------------------------

    def _browser(self):
        agent = self

        class _Ctx:
            async def __aenter__(self):
                from playwright.async_api import async_playwright

                self._pw = await async_playwright().start()
                self._browser = await self._pw.chromium.launch(headless=agent.headless)
                state = str(agent.state_path) if agent.state_path.exists() else None
                self._ctx = await self._browser.new_context(
                    storage_state=state,
                    locale="ru-RU",
                    viewport={"width": 1440, "height": 900},
                )
                self._ctx.set_default_timeout(agent.nav_timeout_ms)
                return await self._ctx.new_page()

            async def __aexit__(self, *exc):
                # Куки продлеваются — сохраняем, НО только если сессия была
                # рабочей: иначе рабочий state.json перезапишется гостевым.
                if agent._session_ok and agent.state_path.exists():
                    try:
                        await self._ctx.storage_state(path=str(agent.state_path))
                    except Exception as e:  # noqa: BLE001
                        log.warning("[hh_apply] не сохранил state.json: %s", e)
                # Каждый шаг закрытия — отдельно: сбой одного не должен
                # оставить висеть процесс chromium.
                for step in (self._ctx.close, self._browser.close, self._pw.stop):
                    try:
                        await step()
                    except Exception as e:  # noqa: BLE001
                        log.warning("[hh_apply] закрытие браузера (%s): %s",
                                    getattr(step, "__name__", step), e)
                return False

        return _Ctx()

    async def _logged_in(self, page) -> bool:
        """ЗАЩИТА ОТ ДЕФЕКТА №2.

        Раньше факт логина проверялся наличием элементов меню
        (mainmenu_applicantProfile / mainmenu_myResumes), и оба уже исчезали
        при редизайне — приходилось править селекторы. Здесь опираемся на
        поведение, а не на вёрстку: гостя hh редиректит с /applicant/resumes
        на /account/login. Нет редиректа на логин — сессия жива.
        """
        try:
            await page.goto(LOGIN_PROBE_URL, wait_until="domcontentloaded")
        except Exception as exc:  # noqa: BLE001
            log.error("[hh_apply] не открылась %s: %s", LOGIN_PROBE_URL, exc)
            return False

        url = page.url
        if any(m in url for m in LOGIN_URL_MARKERS):
            log.error("[hh_apply] редирект на логин: %s", url)
            return False
        if "hh.ru" not in url:
            log.error("[hh_apply] неожиданный редирект: %s", url)
            return False
        log.info("[hh_apply] сессия жива, hh отдал %s", url)
        self._session_ok = True
        return True

    async def _visible(self, page, selector: str, timeout: float = 2.5):
        """Первый ВИДИМЫЙ элемент из совпадений селектора.

        Ждать видимости именно .first нельзя: hh рисует кнопку отклика в
        нескольких местах (шапка, тело, липкий блок), и первое совпадение в
        DOM часто скрыто — ожидание отваливалось по таймауту, хотя рабочая
        кнопка на странице была.
        """
        loc = page.locator(selector)
        deadline = asyncio.get_running_loop().time() + timeout
        while True:
            try:
                count = min(await loc.count(), 12)
            except Exception:  # noqa: BLE001 — страница могла перерисоваться
                count = 0
            for i in range(count):
                candidate = loc.nth(i)
                try:
                    if await candidate.is_visible():
                        return candidate
                except Exception:  # noqa: BLE001 — элемент исчез между count и проверкой
                    continue
            if asyncio.get_running_loop().time() >= deadline:
                return None
            await page.wait_for_timeout(250)

    async def _screenshot(self, vid: int, tag: str) -> str | None:
        try:
            self.screenshots_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            path = self.screenshots_dir / f"{stamp}-{vid}-{tag}.png"
            await self._page.screenshot(path=str(path), full_page=False)
            return str(path)
        except Exception as exc:  # noqa: BLE001 — скриншот диагностический
            log.debug("[hh_apply] скриншот не сохранён: %s", exc)
            return None

    # --- одна вакансия ----------------------------------------------------

    async def _process(self, page, item: dict, get_letter) -> Attempt:
        """Одна вакансия. get_letter вызывается только если отклик возможен."""
        self._page = page
        vid, url = item["id"], item["url"]

        # ЗАЩИТА ОТ ДЕФЕКТА №7: 11 вакансий прогона 09.09 получили
        # «кнопка отклика не найдена» — и это был ложный диагноз, уводящий в
        # селекторы. На самом деле hh отвечал на эти адреса 403 и вместо
        # вакансии рисовал форму логина (url при этом не менялся, поэтому
        # проверка редиректа молчала, а вёрстки вакансии на странице просто
        # не было). Код ответа надёжнее любого селектора — читаем его.
        try:
            resp = await page.goto(url, wait_until="domcontentloaded")
        except Exception as exc:  # noqa: BLE001
            return self._attempt(item, "", Result.FAILED, f"страница не открылась: {exc}")
        status = resp.status if resp is not None else 0

        if any(m in page.url for m in LOGIN_URL_MARKERS):
            return self._attempt(item, "", Result.FAILED, "hh выкинул на логин")

        if status == 404:
            return self._attempt(item, "", Result.GONE,
                                 "hh ответил 404 — вакансии больше нет")

        if status >= 400:
            detail, permanent = await self._forbidden_detail(page, status)
            return self._attempt(
                item, "",
                Result.NOT_ALLOWED if permanent else Result.FORBIDDEN,
                detail, await self._screenshot(vid, "forbidden"),
            )

        self._page_read = True

        if await self._archived(page):
            return self._attempt(item, "", Result.ARCHIVED, "вакансия в архиве")

        if await self._visible(page, SEL["already_applied"], timeout=1.5):
            return self._attempt(item, "", Result.ALREADY_APPLIED,
                                 "отклик уже был отправлен ранее")

        button = await self._visible(page, SEL["respond_button"], timeout=5)
        if button is None:
            # Форму логина hh отдаёт и с кодом 200 — проверяем перед тем, как
            # писать в отчёт «кнопка не найдена».
            if await self._visible(page, SEL["login_wall"], timeout=1.5):
                detail, permanent = await self._forbidden_detail(page, status)
                return self._attempt(
                    item, "",
                    Result.NOT_ALLOWED if permanent else Result.FORBIDDEN,
                    detail, await self._screenshot(vid, "forbidden"),
                )
            # Скриншот нужен и в dry-run: без него причина не восстанавливается.
            return self._attempt(item, "", Result.NO_RESPOND_BUTTON,
                                 "кнопка отклика не найдена",
                                 await self._screenshot(vid, "no-respond-button"))

        try:
            letter = get_letter()
        except Exception as exc:  # noqa: BLE001 — сбой генерации не роняет прогон
            log.exception("[hh_apply] #%s: ошибка генерации письма", vid)
            return self._attempt(item, "", Result.NO_LETTER, f"{type(exc).__name__}: {exc}")

        problem = letter_problem(letter)
        if problem:
            return self._attempt(item, letter, Result.NO_LETTER, problem)

        if self.dry_run:
            # Ни одного клика: часть вакансий отвечает на «Откликнуться»
            # мгновенной отправкой без попапа — в предпросмотре это был бы
            # реальный отклик. Поэтому только читаем состояние страницы.
            return self._attempt(item, letter, Result.DRY_RUN,
                                 "кнопка отклика на месте, письмо готово — отправка не выполнялась")

        try:
            return await self._submit(page, item, letter, button)
        except Exception as exc:  # noqa: BLE001 — одна вакансия не роняет прогон
            log.exception("[hh_apply] #%s: ошибка отправки", vid)
            return self._attempt(item, letter, Result.FAILED, f"{type(exc).__name__}: {exc}",
                                 await self._screenshot(vid, "error"))

    async def _forbidden_detail(self, page, status: int) -> tuple[str, bool]:
        """Почему страница не прочитана и навсегда ли это.

        Причину пишем так, чтобы её не искали в селекторах: код ответа,
        подменённая форма логина и — если hh сказал прямо — то, что вакансия
        просто закрыта для нашего аккаунта.
        """
        try:
            body = await page.locator("body").first.inner_text()
        except Exception:  # noqa: BLE001 — страница могла не дорисоваться
            body = ""
        if NOT_ALLOWED_TEXT_RE.search(body):
            return (
                f"hh не отдал вакансию этому аккаунту (HTTP {status}): она видна "
                "только приглашённым соискателям и работодателю — отклик невозможен",
                True,
            )
        detail = f"hh не отдал страницу (HTTP {status})"
        if await self._visible(page, SEL["login_wall"], timeout=1.5):
            detail += ", вместо вакансии форма логина"
        return detail, False

    async def _archived(self, page) -> bool:
        """Вакансия снята. hh пишет об этом в vacancy-title-archived-text
        внутри заголовка; текстовая проверка — страховка от переименования."""
        if await self._visible(page, SEL["vac_archived"], timeout=1.5):
            return True
        try:
            title = await page.locator(SEL["vac_title"]).first.inner_text()
        except Exception:  # noqa: BLE001
            return False
        return bool(ARCHIVED_TEXT_RE.search(title))

    async def _submit(self, page, item: dict, letter: str, button) -> Attempt:
        vid = item["id"]
        await button.click()
        await page.wait_for_load_state("domcontentloaded")

        # ЗАЩИТА ОТ ДЕФЕКТА №8. Первым делом — не ушёл ли отклик уже от самого
        # клика: у вакансий без анкеты «Откликнуться» отправляет сразу, а
        # письмо hh предлагает приложить после. Проверять это надо ДО поиска
        # поля письма, иначе состоявшаяся отправка уходит в отчёт как
        # «поле письма не найдено», то есть как будто ничего не произошло.
        if await self._visible(page, SEL["already_applied"], timeout=2):
            return await self._applied_before_letter(
                page, item, letter, "до поля письма")

        # hh может предупредить о другом регионе — подтверждаем.
        relocation = await self._visible(page, SEL["relocation_confirm"], timeout=2)
        if relocation is not None:
            await relocation.click()
            await page.wait_for_timeout(800)

        if await self._visible(page, SEL["questions_form"], timeout=1.5):
            return self._attempt(item, letter, Result.NEEDS_MANUAL,
                                 "вакансия требует анкету работодателя",
                                 await self._screenshot(vid, "questions"))

        # Поле письма скрыто за переключателем «Сопроводительное письмо»
        # (data-qa=vacancy-response-letter-toggle, на странице он есть и виден).
        # Порядок важен: до раскрытия поля нет вообще, поэтому ждать его первым
        # — потерянные секунды. Сначала короткая проверка на случай вёрстки,
        # где поле сразу на месте, затем раскрываем.
        letter_input = await self._visible(page, SEL["letter_input"], timeout=1.5)
        if letter_input is None:
            toggle = await self._visible(page, SEL["letter_toggle"], timeout=3)
            if toggle is not None:
                await toggle.click()
                await page.wait_for_timeout(500)
                letter_input = await self._visible(page, SEL["letter_input"], timeout=4)
                if letter_input is None:
                    letter_input = await self._visible(
                        page, SEL["letter_input_any"], timeout=2
                    )

        if letter_input is None:
            # Ещё одна сверка: hh мог отправить отклик пока мы искали поле.
            if await self._visible(page, SEL["already_applied"], timeout=1.5):
                return await self._applied_before_letter(
                    page, item, letter, "пока искалось поле письма")
            return self._attempt(item, letter, Result.NO_LETTER_FIELD,
                                 "поле сопроводительного письма не найдено",
                                 await self._screenshot(vid, "no-letter-field"))

        await letter_input.fill(letter)

        submit = await self._visible(page, SEL["submit"], timeout=5)
        if submit is None:
            return self._attempt(item, letter, Result.NO_SUBMIT,
                                 "кнопка отправки не найдена",
                                 await self._screenshot(vid, "no-submit"))

        await submit.click()

        if await self._confirmed(page):
            return self._attempt(item, letter, Result.APPLIED, "отклик отправлен")
        return self._attempt(item, letter, Result.UNCONFIRMED,
                             "не удалось подтвердить отправку",
                             await self._screenshot(vid, "unconfirmed"))

    async def _attach_letter(self, page, item: dict, letter: str) -> bool:
        """Дописать письмо в уже созданный отклик и подтвердить отправку.

        У вакансий без анкеты hh отправляет отклик сразу на «Откликнуться», а
        письмо предлагает приложить после — ссылкой «Приложить сопроводительное
        письмо» на странице успеха. Без этого шага результат нулевой при
        сделанной работе: отклик есть, работодатель видит одно резюме.
        """
        from playwright.async_api import TimeoutError as PWTimeout

        attach = await self._visible(page, SEL["attach_letter"], timeout=4)
        if attach is None:
            return False
        await attach.click()

        letter_input = await self._visible(page, SEL["letter_input"], timeout=6)
        if letter_input is None:
            letter_input = await self._visible(page, SEL["letter_input_any"], timeout=3)
        if letter_input is None:
            return False
        await letter_input.fill(letter)

        submit = await self._visible(page, SEL["submit"], timeout=4)
        if submit is None:
            return False
        await submit.click()

        # Подтверждение только по исчезновению поля: маркер already_applied
        # здесь бесполезен, он висит на странице с самого момента отклика и
        # подтвердил бы отправку письма, которого не было.
        try:
            await page.locator(SEL["letter_input"]).first.wait_for(
                state="hidden", timeout=8000
            )
            return True
        except PWTimeout:
            return False

    async def _applied_before_letter(self, page, item: dict, letter: str,
                                     where: str) -> Attempt:
        """Отклик уже ушёл от клика — пробуем дописать письмо в него."""
        if await self._attach_letter(page, item, letter):
            return self._attempt(
                item, letter, Result.APPLIED,
                f"hh отправил отклик сразу на клике ({where}), "
                f"письмо дописано в отклик",
                await self._screenshot(item["id"], "applied-letter-attached"),
            )
        return self._attempt(
            item, letter, Result.APPLIED_NO_LETTER,
            f"отклик ушёл без письма ({where}), дописать не удалось — "
            f"письмо нужно приложить в отклике вручную",
            await self._screenshot(item["id"], "applied-no-letter"),
        )

    async def _confirmed(self, page) -> bool:
        """Успех: появился маркер отклика либо исчезла форма письма."""
        from playwright.async_api import TimeoutError as PWTimeout

        if await self._visible(page, SEL["already_applied"], timeout=8):
            return True
        try:
            await page.locator(SEL["letter_input"]).first.wait_for(
                state="hidden", timeout=5000
            )
            return True
        except PWTimeout:
            return False


def preview(config, limit: int | None = None) -> dict:
    """Очередь и письма без обращения к hh.ru — быстрый взгляд «что отправится»."""
    agent = HHApplyAgent(config, dry_run=True, limit=limit)
    session = agent._Session()
    try:
        agent._ensure_log_table(session)
        queue = agent._queue(session)
        take = queue if limit is None else queue[:limit]
        items: list[dict[str, Any]] = []
        for item in take:
            had_draft = bool(item["draft"])  # _letter перезапишет поле, если сгенерит
            letter = agent._letter(session, item)
            items.append({
                **{k: item[k] for k in ("id", "url", "role", "company", "score", "profile")},
                "letter": letter,
                "letter_source": "draft_text" if had_draft else "composer",
                "detail": letter_problem(letter) or "",
            })
        return {"queue_size": len(queue), "shown": len(items),
                "rate": rate_status(session), "items": items}
    finally:
        session.close()
