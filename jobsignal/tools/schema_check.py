"""Сверить модели SQLAlchemy с настоящей схемой базы.

Зачем. За три дня одно и то же расхождение всплыло трижды: applications без
message_text/channel/is_draft в модели, reply_state, которого нет в старых
копиях, channels.enabled — колонка NOT NULL, о которой модель не знала, из-за
чего падала любая вставка канала.

Каждый раз это выяснялось в худший момент: когда понадобилось записать. Модель
расходится с таблицей молча, и чтение обычно работает — ломается именно
запись, причём не вся, а та, что затрагивает забытую колонку.

Проверка читает pragma table_info и сравнивает с метаданными моделей. Ничего
не меняет: решение, что делать с расхождением, зависит от его вида, и
автоматика тут навредит.

Опасное выделено отдельно: колонка NOT NULL без значения по умолчанию, которой
нет в модели, — это гарантированная ошибка при следующей вставке.

    cd /opt/jobsignal_local && .venv/bin/python tools/schema_check.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from sqlalchemy import inspect, text  # noqa: E402

from jobsignal.db import Base, get_engine  # noqa: E402


def main() -> int:
    engine = get_engine()
    insp = inspect(engine)
    real_tables = set(insp.get_table_names())
    problems = fatal = 0

    with engine.connect() as conn:
        for table in Base.metadata.sorted_tables:
            name = table.name
            if name not in real_tables:
                print(f"[{name}] таблицы в базе НЕТ — модель её создаст при "
                      f"следующем create_all")
                problems += 1
                continue

            rows = conn.execute(text(f"pragma table_info({name})")).fetchall()
            # pragma: (cid, name, type, notnull, dflt_value, pk)
            real = {r[1]: {"notnull": bool(r[3]), "default": r[4]} for r in rows}
            model = {c.name for c in table.columns}

            missing_in_model = sorted(set(real) - model)
            missing_in_db = sorted(model - set(real))

            for col in missing_in_model:
                info = real[col]
                # Вот это и ломает запись: вставка не передаёт колонку, а
                # база требует значение и своего не подставляет.
                danger = info["notnull"] and info["default"] is None
                mark = "ОПАСНО" if danger else "лишняя"
                print(f"[{name}] {mark}: колонка '{col}' есть в базе, "
                      f"нет в модели"
                      + (" (NOT NULL без default — вставка через ORM упадёт)"
                         if danger else ""))
                problems += 1
                fatal += 1 if danger else 0

            for col in missing_in_db:
                print(f"[{name}] колонка '{col}' есть в модели, нет в базе — "
                      f"запись в неё потеряется или упадёт")
                problems += 1
                fatal += 1

    if not problems:
        print("модели и схема совпадают")
        return 0
    print(f"\nрасхождений: {problems}, из них ломающих запись: {fatal}")
    # Код возврата 1 при ломающих: так проверку можно поставить в развёртывание
    # и узнавать о расхождении до того, как оно уронит рабочий прогон.
    return 1 if fatal else 0


if __name__ == "__main__":
    sys.exit(main())
