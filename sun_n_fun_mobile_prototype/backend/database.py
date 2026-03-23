"""Database configuration and session helpers.

This module configures the SQLAlchemy engine and session used by the
application.  The current configuration uses a local SQLite database
(`sun_n_fun.db`) which is suitable for a single‑node prototype.  For
multi‑user or production deployments, consider switching to PostgreSQL
and adjusting the connection string accordingly.
"""

from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# The connection URL for SQLite.  For other databases (e.g. PostgreSQL),
# update this URL accordingly.  When using SQLite with SQLAlchemy in
# multi‑threaded environments, `check_same_thread` must be set to False.
SQLALCHEMY_DATABASE_URL = "sqlite:///./sun_n_fun.db"

engine = create_engine(
    SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False}
)

# Create a configured "SessionLocal" class.  Each instance will be a
# database session bound to our engine.  The session does not autocommit or
# autoflush; callers are responsible for committing transactions as needed.
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# Base class for declarative models.  All ORM models should inherit from
# `Base`.  SQLAlchemy uses this class to collect model metadata.
Base = declarative_base()


def get_db():
    """Yield a new database session.

    FastAPI dependency that opens a new session at the start of a request and
    closes it when the request is finished.  Use `Depends(get_db)` in your
    path operations to access the database.
    """
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()