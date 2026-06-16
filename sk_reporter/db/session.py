"""Сессии SQLAlchemy."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Iterator

from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from sk_reporter.db.config import database_url

_engine = None
_SessionLocal: sessionmaker[Session] | None = None


def _ensure_engine():
    global _engine, _SessionLocal
    url = database_url()
    if not url:
        raise RuntimeError("DATABASE_URL не задан")
    if _engine is None:
        _engine = create_engine(url, pool_pre_ping=True)
        _SessionLocal = sessionmaker(bind=_engine, autoflush=False, autocommit=False)
    return _SessionLocal


@contextmanager
def get_session() -> Iterator[Session]:
    factory = _ensure_engine()
    session = factory()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def init_db() -> bool:
    """Создать таблицы, если их ещё нет (dev / первый деплой без alembic)."""
    from sk_reporter.db.base import Base
    from sk_reporter.db import models  # noqa: F401

    url = database_url()
    if not url:
        return False
    engine = create_engine(url, pool_pre_ping=True)
    Base.metadata.create_all(engine)
    return True
