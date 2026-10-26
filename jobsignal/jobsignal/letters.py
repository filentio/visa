"""Сопроводительные письма — один сервис на весь проект.

Зачем. До 26.10 писем было два генератора с разными правилами:

  * notify_bot._cover_text — кнопка в кабинете, бот, автоотклик hh. Поверх
    правил голоса стоял шаблон из базы (cover_templates, «Метрика-крючок»:
    5-7 предложений, «кто я + компании», «Буду рад обсудить»), который этим
    правилам противоречил;
  * composer.generate_draft — карточки отклика в Telegram: «ровно 3-4
    строки», резюме обрезано до 4000 знаков, свой список стилей.

Хуже всего было на стыке: автоотклик брал любой сохранённый черновик, если он
проходил проверку на обрыв. Туда попадало и короткое сообщение, написанное
для Telegram, и письмо под снятый профиль. А выбор стиля в кабинете ни на что
не влиял: сервер его не читал.

Теперь:
  * одно построение промпта: общие правила, правила канала (hh, telegram,
    email, form), акцент (необязательно), правила голоса letter_style;
  * одни входные данные: полный текст вакансии и полный текст резюме профиля
    с лучшим баллом среди действующих профилей;
  * одна проверка качества: обрыв, шаблонные маркеры, длина канала и числа,
    которых нет ни в резюме, ни в вакансии. Есть замечания — один повтор с
    перечнем замечаний, из двух вариантов берётся лучший;
  * черновик помнит канал, профиль и версию правил (vacancies.draft_channel,
    draft_profile, draft_version). Автоотклик переиспользует только черновик
    для hh под тот же профиль и по тем же правилам; остальное пишет заново.
"""
from __future__ import annotations

import logging
import re
import sqlite3
from dataclasses import dataclass, field

log = logging.getLogger("jobsignal")

# Меняется при любой правке правил ниже: старые черновики после этого не
# переиспользуются автоотправкой.
LETTER_VERSION = "2026-10-26c"

VAC_CAP = 3000      # текст вакансии: требования почти всегда в начале
RESUME_CAP = 8000   # резюме целиком: новые резюме 6700-7400 знаков

CHANNELS: dict[str, dict] = {
    "hh": {
        "label": "сопроводительное письмо к отклику на hh.ru",
        "min": 500, "max": 800, "ceil": 850,
        "form": "Два-три коротких абзаца, между ними пустая строка.",
        "notes": ("Резюме уже приложено к отклику: не предлагай его прислать "
                  "и не пересказывай. Начни с «Добрый день» (с именем, если "
                  "оно известно)."),
    },
    "telegram": {
        "label": "первое сообщение рекрутёру в личку Telegram",
        "min": 300, "max": 550, "ceil": 600,
        "form": "Три-пять коротких строк, без длинных абзацев.",
        "notes": ("PDF резюме приложен к сообщению: не предлагай его прислать. "
                  "Начни с короткого приветствия (с именем, если известно). "
                  "В конце предложи следующий шаг со своей стороны, без просьбы "
                  "рассказать о вакансии подробнее."),
    },
    "email": {
        "label": "письмо рекрутёру на почту",
        "min": 500, "max": 900, "ceil": 950,
        "form": "Два-три коротких абзаца, между ними пустая строка.",
        "notes": ("Резюме приложено к письму. Без темы письма и без подписи: "
                  "их добавят отдельно."),
    },
    "form": {
        "label": "текст в поле «сопроводительное письмо / о себе» анкеты на сайте компании",
        "min": 300, "max": 600, "ceil": 650,
        "form": "Один-два коротких абзаца.",
        "notes": "Без приветствия по имени: кто прочтёт анкету, неизвестно.",
    },
}
DEFAULT_CHANNEL = "hh"

# Акценты — выбор в кабинете. Без акцента модель сама решает по вакансии.
ACCENTS: dict[str, dict] = {
    "auto": {"label": "Под вакансию (авто)", "hint": ""},
    "metric": {"label": "Метрика-крючок",
               "hint": "АКЦЕНТ: после приветствия одна самая сильная измеримая "
                       "метрика из резюме, которая отвечает на их задачу."},
    "fit": {"label": "Под требования",
            "hint": "АКЦЕНТ: два-три точных совпадения опыта из резюме с "
                    "требованиями вакансии, каждое с конкретикой."},
    "builder": {"label": "Делаю руками",
                "hint": "АКЦЕНТ: умение самому собрать прототип или MVP (AI-стек "
                        "из резюме) и скорость проверки гипотез."},
    "domain": {"label": "Домен и масштаб",
               "hint": "АКЦЕНТ: доменный опыт и масштаб из резюме: аудитория, "
                       "команда, деньги, бизнес-метрики."},
}
# Старые ключи из кабинета и composer.STYLES.
_ACCENT_ALIASES = {"metric_hook": "metric", "fintech": "domain"}
DEFAULT_ACCENT = "auto"

BASE_RULES = (
    "Ты пишешь от первого лица, от имени кандидата, {channel_label}. "
    "Профиль кандидата: {profile_label}.\n\n"
    "Задача письма — чтобы рекрутёр за десять секунд увидел, чем кандидат "
    "закрывает задачу ЭТОЙ вакансии, и открыл резюме. Факты о кандидате "
    "только из его резюме ниже. Даты работы бери буквально: текущее место — "
    "то, у которого стоит «настоящее время», остальные описывай в прошедшем "
    "времени. Пиши только текст письма: без темы, без подписи, без markdown."
    "\n\n{channel_notes}"
)


@dataclass
class Letter:
    text: str
    channel: str
    profile: str
    accent: str
    version: str = LETTER_VERSION
    issues: list = field(default_factory=list)   # замечания, которые остались

    @property
    def ok(self) -> bool:
        """Можно отправлять: не обрыв и не заглушка (стиль — не повод не слать)."""
        from jobsignal.agents.hh_apply import letter_problem
        return bool(self.text) and letter_problem(self.text) is None


def norm_channel(channel: str | None) -> str:
    return channel if channel in CHANNELS else DEFAULT_CHANNEL


def norm_accent(accent: str | None) -> str:
    a = _ACCENT_ALIASES.get(accent or "", accent or "")
    return a if a in ACCENTS else DEFAULT_ACCENT


def system_prompt(channel: str, profile_label: str, accent: str = DEFAULT_ACCENT) -> str:
    from jobsignal.letter_style import voice_rules
    ch = CHANNELS[norm_channel(channel)]
    parts = [BASE_RULES.format(channel_label=ch["label"], profile_label=profile_label,
                               channel_notes=ch["notes"])]
    hint = ACCENTS[norm_accent(accent)]["hint"]
    if hint:
        parts.append(hint)
    parts.append(voice_rules(ch["min"], ch["max"], ch["form"]))
    return "\n\n".join(parts)


def _resume_text(profile_key: str) -> str:
    from jobsignal.agents.resume_parser import get_resume_text, parse_and_save
    text = get_resume_text(profile_key)
    if not text:
        if parse_and_save(profile_key).get("ok"):
            text = get_resume_text(profile_key)
    if not text:
        from jobsignal import profiles
        p = profiles.by_key(profile_key)
        text = p.cv_text() if p else ""
    return text or ""


FACT_SYSTEM = (
    "Ты проверяешь сопроводительное письмо кандидата на соответствие его "
    "резюме. Найди в письме утверждения О КАНДИДАТЕ, которых нет в резюме "
    "или которые искажают резюме: навыки, инструменты, методики, роли, "
    "компании, домены, масштаб, сроки и цифры (в том числе «15+» вместо 15, "
    "«10-15 минут» вместо 15, «рекомендательная система» вместо «система "
    "подбора»). Особенно ищи требования вакансии, приписанные кандидату. "
    "Пересказ своими словами без изменения смысла — не ошибка. Утверждения о "
    "вакансии и компании не проверяй.\n\n"
    "Отвечай строго одним JSON без markdown: "
    '{"unsupported": ["<цитата из письма> — <что в резюме на самом деле>"]}. '
    "Если всё подтверждается, верни пустой список."
)


def fact_check(text: str, resume: str, vacancy_text: str = "") -> list[str]:
    """Утверждения о кандидате, которых нет в резюме. Дешёвая модель разбора.

    Проверка цифр (invented_numbers) ловит только выдуманные числа. 26.10 в
    письмах нашлись «user story, UML» (UML в резюме нет, это требование
    вакансии), «15+ инициатив» (в резюме 15) и «до 10-15 минут» (в резюме
    15): числа такие в резюме встречаются, а смысл искажён.
    Отключается LETTER_FACT_CHECK=0. Сбой проверки письмо не блокирует.
    """
    import os
    if os.environ.get("LETTER_FACT_CHECK", "1") == "0":
        return []
    from jobsignal.config import get_config
    from jobsignal.llm import complete_json
    user = (f"РЕЗЮМЕ КАНДИДАТА:\n{resume[:RESUME_CAP]}\n\n"
            + (f"ТЕКСТ ВАКАНСИИ (для понимания, что из неё могли приписать):\n"
               f"{vacancy_text.strip()[:1500]}\n\n" if (vacancy_text or "").strip() else "")
            + f"ПИСЬМО:\n{text}")
    try:
        data = complete_json(FACT_SYSTEM, user, model=get_config().settings.parser_model,
                             max_tokens=700, tag="letter_facts", cache_system=True)
    except Exception as exc:  # noqa: BLE001
        log.warning("[letters] проверка фактов не выполнилась: %s", exc)
        return []
    items = data.get("unsupported") or []
    return [str(x)[:200] for x in items if str(x).strip()][:6]


def review(text: str, channel: str, resume: str, vacancy_text: str,
           facts: bool = True) -> list[str]:
    """Замечания к письму. Пусто — письмо чистое."""
    from jobsignal.agents.hh_apply import letter_problem
    from jobsignal.letter_style import invented_numbers, style_tells
    ch = CHANNELS[norm_channel(channel)]
    issues = []
    problem = letter_problem(text)
    if problem:
        issues.append(problem)
    issues += style_tells(text, max_chars=ch["ceil"])
    if len(text.strip()) < ch["min"] * 0.7:
        issues.append(f"короткое: {len(text.strip())} знаков при норме от {ch['min']}")
    nums = invented_numbers(text, resume, vacancy_text)
    if nums:
        issues.append("числа, которых нет ни в резюме, ни в вакансии: " + ", ".join(nums))
    if facts and not problem:
        for item in fact_check(text, resume, vacancy_text):
            issues.append("нет в резюме: " + item)
    return issues


def generate(role: str, company: str, vacancy_text: str, profile: str | None = None,
             channel: str = DEFAULT_CHANNEL, accent: str | None = None,
             recruiter_name: str = "") -> Letter:
    """Письмо под вакансию. profile — имя или ключ профиля; пусто — по умолчанию."""
    from jobsignal import profiles
    from jobsignal.config import get_config
    from jobsignal.llm import complete_text

    channel = norm_channel(channel)
    accent = norm_accent(accent)
    prof = profiles.resolve(profile)
    pname = prof.name if prof else (profile or "")
    resume = _resume_text(prof.key) if prof else ""
    if not resume.strip():
        # Без резюме модель вынуждена выдумывать — письма не будет вовсе.
        log.error("[letters] нет текста резюме для профиля «%s» — письмо не пишем", pname)
        return Letter("", channel, pname, accent, issues=["нет текста резюме профиля"])

    company_part = f" в {company}" if company and company not in ("—", "None") else ""
    user = (
        f"Вакансия: {role or '—'}{company_part}\n"
        f"Имя рекрутёра: {recruiter_name or 'неизвестно'}\n\n"
        + (f"Текст вакансии:\n{vacancy_text.strip()[:VAC_CAP]}\n\n"
           if (vacancy_text or "").strip() else "")
        + f"Резюме кандидата:\n{resume[:RESUME_CAP]}\n\n"
        "Напиши письмо. Факты о кандидате только из резюме. Из текста вакансии "
        "бери только их задачу и не приписывай кандидату её требования. "
        f"Длина строго до {CHANNELS[channel]['max']} знаков: лучше короче, чем длиннее."
    )
    system = system_prompt(channel, prof.letter_label if prof else pname, accent)
    model = get_config().settings.anthropic_model

    try:
        text = clean_for_channel(complete_text(system, user, model, 2000,
                                               cache_system=True))
    except Exception as exc:  # noqa: BLE001
        log.warning("[letters] модель не ответила: %s", exc)
        return Letter("", channel, pname, accent, issues=[f"модель не ответила: {exc}"])
    issues = review(text, channel, resume, vacancy_text)
    if issues:
        # Один повтор с перечнем замечаний; дальше берём вариант с меньшим
        # числом замечаний. Больше повторов дороже, а лучше почти не становится.
        try:
            retry = clean_for_channel(complete_text(
                system,
                user + "\n\nПредыдущий вариант не подошёл: " + "; ".join(issues)
                + ". Перепиши с учётом этого.\n\nПредыдущий вариант:\n" + text,
                model, 2000, cache_system=True))
            retry_issues = review(retry, channel, resume, vacancy_text)
            if len(retry_issues) <= len(issues):
                text, issues = retry, retry_issues
        except Exception as exc:  # noqa: BLE001 — есть первый вариант
            log.warning("[letters] повтор не удался: %s", exc)
    if issues:
        log.info("[letters] %s/%s: остались замечания — %s", channel, pname, "; ".join(issues))
    return Letter(text, channel, pname, accent, issues=issues)


def _db_path() -> str:
    from jobsignal.db import _default_url
    return _default_url().replace("sqlite:///", "", 1)


def vacancy_inputs(vacancy_id: int) -> dict | None:
    """Роль, компания, полный текст и лучший действующий профиль вакансии."""
    from jobsignal import profiles
    c = sqlite3.connect(_db_path(), timeout=15)
    try:
        row = c.execute(
            "SELECT v.role, v.company, coalesce(p.text, v.description, '') "
            "FROM vacancies v LEFT JOIN raw_posts p ON p.id = v.raw_post_id "
            "WHERE v.id = ?", (vacancy_id,)).fetchone()
        if not row:
            return None
        active = profiles.names()
        best = None
        if active:
            marks = ",".join("?" * len(active))
            r = c.execute(
                f"SELECT profile_key FROM match_scores WHERE vacancy_id=? "
                f"AND profile_key IN ({marks}) ORDER BY score DESC LIMIT 1",
                (vacancy_id, *active)).fetchone()
            best = r[0] if r else None
    finally:
        c.close()
    return {"role": row[0] or "", "company": row[1] or "", "text": row[2] or "",
            "profile": best or profiles.default_name()}


def for_vacancy(vacancy_id: int, channel: str = DEFAULT_CHANNEL,
                accent: str | None = None, profile: str | None = None,
                recruiter_name: str = "") -> Letter:
    inp = vacancy_inputs(vacancy_id)
    if inp is None:
        return Letter("", norm_channel(channel), profile or "", norm_accent(accent),
                      issues=[f"вакансии #{vacancy_id} нет"])
    return generate(inp["role"], inp["company"], inp["text"],
                    profile=profile or inp["profile"], channel=channel,
                    accent=accent, recruiter_name=recruiter_name)


def remember(v, letter: Letter) -> None:
    """Записать черновик в вакансию (ORM-объект) вместе с каналом, профилем и
    версией правил. Коммит — на вызывающем."""
    v.draft_text = letter.text
    v.draft_channel = letter.channel
    v.draft_profile = letter.profile
    v.draft_version = letter.version


def reusable(v, channel: str, profile: str) -> str | None:
    """Черновик вакансии, если он написан для этого канала, профиля и по
    действующим правилам и не обрывается. Иначе None — писать заново."""
    from jobsignal.agents.hh_apply import letter_problem
    text = (getattr(v, "draft_text", None) or "").strip()
    if not text:
        return None
    if (getattr(v, "draft_channel", None) != norm_channel(channel)
            or getattr(v, "draft_profile", None) != profile
            or getattr(v, "draft_version", None) != LETTER_VERSION):
        return None
    return text if letter_problem(text) is None else None


def clean_for_channel(text: str) -> str:
    """Подчистка ответа модели: markdown и кавычки вокруг всего письма."""
    t = (text or "").strip().strip("«»\"")
    return re.sub(r"\*\*(.+?)\*\*", r"\1", t)
