"""SQLAlchemy models representing the core entities.

These models are intentionally close to the roadmap definitions to keep the
domain understandable.  Relationships are declared to allow navigation
between volunteers, event days and assignments.  Additional entities such as
`AuditEvent` can be added later as needed.
"""

from datetime import date, datetime
from sqlalchemy import Column, Integer, String, Boolean, ForeignKey, Date, Text, DateTime
from sqlalchemy.orm import relationship

from .database import Base


class Volunteer(Base):
    __tablename__ = "volunteers"

    id: int = Column(Integer, primary_key=True, index=True)
    first_name: str = Column(String, index=True)
    last_name: str = Column(String, index=True)
    phone: str | None = Column(String, nullable=True)
    email: str | None = Column(String, unique=True, index=True, nullable=True)
    shirt_size: str | None = Column(String, nullable=True)
    camping: bool | None = Column(Boolean, default=False)
    notes: str | None = Column(Text, nullable=True)
    emergency_contact: str | None = Column(String, nullable=True)
    tags: str | None = Column(String, nullable=True)

    # Relationship to assignments.  Setting cascade behaviour so that
    # assignments are deleted when a volunteer is removed could be done here,
    # e.g. `cascade="all, delete"`, but is omitted to keep the prototype simple.
    assignments = relationship("Assignment", back_populates="volunteer")


class EventDay(Base):
    __tablename__ = "event_days"

    id: int = Column(Integer, primary_key=True, index=True)
    date: date = Column(Date, index=True)
    phase: str | None = Column(String, nullable=True)
    hours: str | None = Column(String, nullable=True)
    status: str | None = Column(String, nullable=True)
    capacity: int | None = Column(Integer, nullable=True)

    assignments = relationship("Assignment", back_populates="event_day")


class Assignment(Base):
    __tablename__ = "assignments"

    id: int = Column(Integer, primary_key=True, index=True)
    volunteer_id: int = Column(Integer, ForeignKey("volunteers.id"))
    event_day_id: int = Column(Integer, ForeignKey("event_days.id"))
    check_in_status: str | None = Column(String, nullable=True)
    role: str | None = Column(String, nullable=True)
    notes: str | None = Column(Text, nullable=True)
    source: str | None = Column(String, nullable=True)
    created_by: str | None = Column(String, nullable=True)
    updated_by: str | None = Column(String, nullable=True)
    updated_at_utc: datetime | None = Column(DateTime, nullable=True)

    volunteer = relationship("Volunteer", back_populates="assignments")
    event_day = relationship("EventDay", back_populates="assignments")
