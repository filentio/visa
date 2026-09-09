from __future__ import annotations
import enum
from datetime import datetime, timezone
from sqlalchemy import (
    create_engine, Column, Integer, String, Text, Float,
    DateTime, Boolean, ForeignKey, UniqueConstraint, Index, text,
    Enum as SAEnum
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

Base = declarative_base()


class VacancyStatus(enum.Enum):
    new = "new"
    matched = "matched"
    drafted = "drafted"
    applied = "applied"
    replied = "replied"
    skipped = "skipped"
    rejected = "rejected"
    no_reply = "no_reply"
    interview = "interview"
    offer = "offer"


class Channel(Base):
    __tablename__ = "channels"
    id = Column(Integer, primary_key=True)
    handle = Column(String, unique=True, nullable=False)
    title = Column(String)
    niche = Column(String)
    active = Column(Boolean, default=True)
    last_message_id = Column(Integer, default=0)
    added_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    # manual / tgstat / telemetr / auto_search
    source = Column(String, default="manual")
    subscriber_count = Column(Integer, default=0)
    post_count_30d = Column(Integer, default=0)   # activity check
    verified_at = Column(DateTime)                # last time we confirmed it's alive
    raw_posts = relationship("RawPost", back_populates="channel")


class RawPost(Base):
    __tablename__ = "raw_posts"
    id = Column(Integer, primary_key=True)
    channel_id = Column(Integer, ForeignKey("channels.id"), nullable=False)
    tg_message_id = Column(Integer, nullable=False)
    text = Column(Text)
    raw_text = Column(Text)  # legacy alias
    post_url = Column(String)
    posted_at = Column(DateTime)
    parsed = Column(Boolean, default=False)
    channel = relationship("Channel", back_populates="raw_posts")
    __table_args__ = (UniqueConstraint("channel_id", "tg_message_id"),)


class Vacancy(Base):
    __tablename__ = "vacancies"
    id = Column(Integer, primary_key=True)
    raw_post_id = Column(Integer, ForeignKey("raw_posts.id"))
    channel_id = Column(Integer, ForeignKey("channels.id"))
    role = Column(String)
    company = Column(String)
    recruiter_handle = Column(String)   # @username for TG direct message
    salary = Column(String)
    location = Column(String)
    link = Column(String)               # apply URL (hh.ru, form, etc.)
    description = Column(Text)
    status = Column(SAEnum(VacancyStatus), default=VacancyStatus.new)
    dedup_group = Column(String)
    is_primary = Column(Boolean, default=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    draft_text = Column(Text)
    # contact_type: tg / hh / form / unknown
    contact_type = Column(String, default="unknown")
    match_scores = relationship(
        "MatchScore", back_populates="vacancy", cascade="all, delete-orphan"
    )
    applications = relationship(
        "Application", back_populates="vacancy", cascade="all, delete-orphan"
    )


class MatchScore(Base):
    __tablename__ = "match_scores"
    id = Column(Integer, primary_key=True)
    vacancy_id = Column(Integer, ForeignKey("vacancies.id"), nullable=False)
    profile_key = Column(String, nullable=False)
    score = Column(Integer, default=0)
    reason = Column(Text)
    vacancy = relationship("Vacancy", back_populates="match_scores")
    __table_args__ = (UniqueConstraint("vacancy_id", "profile_key"),)


class Application(Base):
    """Отклик: черновик (is_draft=1, sent_at пуст) или отправленный.

    Модель раньше расходилась с таблицей: здесь были draft_text/notes,
    которых в БД нет, и не было message_text/channel/is_draft/created_at,
    которые в БД объявлены NOT NULL. Из-за этого ORM-запись отклика падала
    (composer), а create_all на чистой базе создавал таблицу, в которую не
    проходил ни один из рабочих INSERT'ов. Держим модель равной схеме.

    message_text и channel умышленно без default: канал («кто отправил» —
    hh / telegram / manual) и текст должен назвать вызывающий код, иначе в
    аналитику попадёт отклик неизвестного происхождения.
    """
    __tablename__ = "applications"
    id = Column(Integer, primary_key=True)
    vacancy_id = Column(Integer, ForeignKey("vacancies.id"), nullable=False)
    message_text = Column(Text, nullable=False)
    channel = Column(String(32), nullable=False)
    is_draft = Column(Boolean, nullable=False, default=False)
    sent_at = Column(DateTime)
    created_at = Column(DateTime, nullable=False,
                        default=lambda: datetime.now(timezone.utc))
    replied_at = Column(DateTime)
    vacancy = relationship("Vacancy", back_populates="applications")

    # Частичные уникальные индексы вместо INSERT OR IGNORE: от повтора
    # защищает схема, а нарушение NOT NULL при этом остаётся видимой ошибкой.
    # Индексы частичные, потому что у вакансии законно бывают две строки —
    # черновик и отправленный отклик (в базе так у #123); запрещаем только
    # второй отправленный и второй черновик.
    __table_args__ = (
        Index("ux_applications_sent_vacancy", "vacancy_id", unique=True,
              sqlite_where=text("sent_at IS NOT NULL")),
        Index("ux_applications_draft_vacancy", "vacancy_id", unique=True,
              sqlite_where=text("is_draft = 1")),
    )


# ── Channel discovery candidates (not yet added) ─────────────────────────────

class ChannelCandidate(Base):
    """Channels found via tgstat/telemetr search, pending review/add."""
    __tablename__ = "channel_candidates"
    id = Column(Integer, primary_key=True)
    handle = Column(String, unique=True, nullable=False)
    title = Column(String)
    description = Column(Text)
    subscriber_count = Column(Integer, default=0)
    source = Column(String)          # tgstat / telemetr / manual_suggest
    search_query = Column(String)    # keyword that found it
    found_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    status = Column(String, default="pending")   # pending / added / rejected


def get_engine(url: str = "sqlite:///./data/jobsignal.db"):
    return create_engine(url, connect_args={"check_same_thread": False})


def get_session_factory(url: str = "sqlite:///./data/jobsignal.db"):
    engine = get_engine(url)
    Base.metadata.create_all(engine)
    return sessionmaker(bind=engine)

from datetime import datetime, timezone
def utcnow():
    return datetime.now(timezone.utc)
