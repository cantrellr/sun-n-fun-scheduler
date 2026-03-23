"""Pydantic models (schemas) for request/response payloads.

These classes define the structure of data exchanged via the API.  They
mirror the SQLAlchemy models but separate internal representation from
external API contracts and provide validation.  All schemas with `Config
orm_mode = True` allow SQLAlchemy objects to be returned directly from
path operations.
"""

from datetime import date, datetime
from typing import List, Optional

from pydantic import BaseModel


class AssignmentBase(BaseModel):
    volunteer_id: int
    event_day_id: int
    check_in_status: Optional[str] = None
    role: Optional[str] = None
    notes: Optional[str] = None
    source: Optional[str] = None
    created_by: Optional[str] = None
    updated_by: Optional[str] = None
    updated_at_utc: Optional[datetime] = None


class AssignmentCreate(AssignmentBase):
    pass


class Assignment(AssignmentBase):
    id: int

    class Config:
        orm_mode = True


class VolunteerBase(BaseModel):
    first_name: str
    last_name: str
    phone: Optional[str] = None
    email: Optional[str] = None
    shirt_size: Optional[str] = None
    camping: Optional[bool] = None
    notes: Optional[str] = None
    emergency_contact: Optional[str] = None
    tags: Optional[str] = None


class VolunteerCreate(VolunteerBase):
    pass


class Volunteer(VolunteerBase):
    id: int
    assignments: List[Assignment] = []

    class Config:
        orm_mode = True


class EventDayBase(BaseModel):
    date: date
    phase: Optional[str] = None
    hours: Optional[str] = None
    status: Optional[str] = None
    capacity: Optional[int] = None


class EventDayCreate(EventDayBase):
    pass


class EventDay(EventDayBase):
    id: int
    assignments: List[Assignment] = []

    class Config:
        orm_mode = True