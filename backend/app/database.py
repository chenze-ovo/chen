"""数据库连接。启动时建表并灌入初始账号与示例订单。"""
from __future__ import annotations

from collections.abc import Generator
from pathlib import Path

from sqlalchemy import create_engine, event
from sqlalchemy.orm import Session, sessionmaker
from sqlalchemy.pool import NullPool

from app.config import get_settings
from app.models import Base
from app.seed import seed

engine = None
SessionLocal = None


def _sqlite_fk(dbapi_connection, _connection_record) -> None:
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()


def reset_db() -> None:
    global engine, SessionLocal
    if engine is not None:
        engine.dispose()
    engine = None
    SessionLocal = None


def init_db() -> None:
    global engine, SessionLocal
    reset_db()
    settings = get_settings()
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    url = settings.resolved_database_url
    if url.startswith("sqlite:///"):
        raw_path = url.removeprefix("sqlite:///")
        if raw_path and raw_path != ":memory:":
            Path(raw_path).parent.mkdir(parents=True, exist_ok=True)
    kwargs: dict = {"pool_pre_ping": True}
    if url.startswith("sqlite"):
        kwargs["connect_args"] = {"check_same_thread": False}
        kwargs["poolclass"] = NullPool
    engine = create_engine(url, **kwargs)
    if url.startswith("sqlite"):
        event.listen(engine, "connect", _sqlite_fk)
    SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False)
    Base.metadata.create_all(engine)
    seed(SessionLocal)


def get_db() -> Generator[Session, None, None]:
    if SessionLocal is None:
        raise RuntimeError("数据库尚未初始化")
    db = SessionLocal()
    try:
        yield db
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()
