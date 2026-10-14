"""Собрать профильные выжимки резюме (config/cv_*.md) из master_cv.pdf.

Зачем. Оценка вакансий держится на трёх текстах — по одному на профиль: AI
PM, CPO / Head of Product, Senior PM/PO. Модель сравнивает вакансию с каждым и
ставит балл. 12.10 старый сервер перестал отвечать, а эти файлы жили только на
нём; на руках остался PDF резюме.

Как собираем. Текст PDF отдаём модели и просим три выжимки: каждая подаёт ТОТ
ЖЕ опыт под свой профиль — что выдвинуть вперёд, что сократить. Добавлять
нельзя ничего: ни цифр, ни мест работы, ни навыков, которых нет в резюме.
Выдуманный факт в выжимке хуже, чем в письме, — он тихо испортит каждую оценку.

Поэтому первый запуск только показывает. Записать — вторым запуском с
--apply, после того как глазами проверено, что модель ничего не дописала.

    cd /opt/jobsignal_local && .venv/bin/python tools/build_cv_profiles.py
    cd /opt/jobsignal_local && .venv/bin/python tools/build_cv_profiles.py --apply
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

PDF = Path("config/master_cv.pdf")
OUT = {
    "ai_pm": ("Senior AI PM", Path("config/cv_ai_pm.md")),
    "cpo": ("CPO / Head of Product", Path("config/cv_cpo.md")),
    "pm": ("Senior PM/PO", Path("config/cv_pm.md")),
}

SYSTEM = """\
Ты готовишь выжимку резюме кандидата под один карьерный профиль. По этой
выжимке другая модель будет оценивать, насколько вакансии подходят кандидату.

ГЛАВНОЕ: только факты из резюме. Не добавляй ни одной цифры, компании,
должности, даты, навыка или проекта, которых нет в тексте. Не округляй и не
усиливай цифры. Если для профиля в резюме мало материала — выжимка будет
короче, это нормально.

Задача — расставить акценты под профиль: что выдвинуть вперёд, что сократить,
какие формулировки из резюме важнее для такой роли. Сохраняй даты работы
буквально; текущее место — то, у которого стоит «по настоящее время» или
аналог.

Формат: markdown, до 3500 знаков. Сначала 2-3 строки позиционирования под
профиль, затем места работы от последнего к первому с ключевыми результатами,
затем навыки и домены. Без вступлений и пояснений — только выжимка."""

FOCUS = {
    "ai_pm": "Senior AI Product Manager: AI/LLM-продукты, RAG, агенты, "
             "скоринг, рекомендательные системы, прототипирование, метрики качества.",
    "cpo": "CPO / Head of Product: продуктовая стратегия, портфель продуктов, "
           "управление командами и продактами, P&L, бизнес-метрики, масштаб.",
    "pm": "Senior Product Manager / Product Owner: hands-on продукт, discovery, "
          "A/B-тесты, воронки, онбординг, финтех, работа с бэклогом.",
}


def main() -> int:
    apply = "--apply" in sys.argv
    if not PDF.exists():
        print(f"нет {PDF} — загрузи PDF резюме под этим именем")
        return 1

    from jobsignal.agents.resume_parser import _extract_pdf_text
    text = (_extract_pdf_text(str(PDF)) or "").strip()
    if len(text) < 500:
        print(f"из PDF извлеклось всего {len(text)} знаков — это скан или "
              f"защищённый файл, выжимку по нему не собрать")
        return 1
    print(f"текст резюме: {len(text)} знаков\n")

    from jobsignal.config import get_config
    from jobsignal.llm import complete_text
    model = get_config().settings.anthropic_model

    for key, (name, path) in OUT.items():
        user = (f"ПРОФИЛЬ: {name}\nАКЦЕНТ: {FOCUS[key]}\n\n"
                f"РЕЗЮМЕ:\n{text[:12000]}")
        out = complete_text(SYSTEM, user, model=model, max_tokens=3000,
                            tag=f"cv_{key}", cache_system=True).strip()
        print(f"===================== {path} ({name}) =====================")
        print(out)
        print()
        if apply:
            if path.exists():
                path.with_suffix(".md.bak").write_text(
                    path.read_text(encoding="utf-8"), encoding="utf-8")
            path.write_text(out + "\n", encoding="utf-8")

    if apply:
        print("записано. Проверка: .venv/bin/python run.py match 5")
    else:
        print("это предпросмотр. Проверь, что модель не дописала фактов, "
              "которых нет в резюме, и повтори с --apply")
    return 0


if __name__ == "__main__":
    sys.exit(main())
