"""CRUD helper functions for interacting with the database.

Functions in this module encapsulate common operations on the models.
Using a separate CRUD layer makes it easy to swap the underlying database
mechanism or add business logic (e.g. capacity checks) without changing
the API routes.
"""

from typing import List, Optional

from sqlalchemy.orm import Session

from . import models, schemas


# Volunteer operations
def get_volunteers(db: Session, skip: int = 0, limit: int = 100) -> List[models.Volunteer]:
    return db.query(models.Volunteer).offset(skip).limit(limit).all()


def get_volunteer(db: Session, volunteer_id: int) -> Optional[models.Volunteer]:
    return db.query(models.Volunteer).filter(models.Volunteer.id == volunteer_id).first()


def create_volunteer(db: Session, volunteer: schemas.VolunteerCreate) -> models.Volunteer:
    db_volunteer = models.Volunteer(**volunteer.dict())
    db.add(db_volunteer)
    db.commit()
    db.refresh(db_volunteer)
    return db_volunteer


def delete_volunteer(db: Session, volunteer_id: int) -> None:
    db_volunteer = get_volunteer(db, volunteer_id)
    if db_volunteer is not None:
        db.delete(db_volunteer)
        db.commit()


def update_volunteer(db: Session, volunteer_id: int, updates: schemas.VolunteerCreate) -> Optional[models.Volunteer]:
    db_volunteer = get_volunteer(db, volunteer_id)
    if db_volunteer is None:
        return None
    for field, value in updates.dict(exclude_unset=True).items():
        setattr(db_volunteer, field, value)
    db.add(db_volunteer)
    db.commit()
    db.refresh(db_volunteer)
    return db_volunteer


# EventDay operations
def get_event_days(db: Session, skip: int = 0, limit: int = 100) -> List[models.EventDay]:
    return db.query(models.EventDay).offset(skip).limit(limit).all()


def get_event_day(db: Session, event_day_id: int) -> Optional[models.EventDay]:
    return db.query(models.EventDay).filter(models.EventDay.id == event_day_id).first()


def create_event_day(db: Session, event_day: schemas.EventDayCreate) -> models.EventDay:
    db_event_day = models.EventDay(**event_day.dict())
    db.add(db_event_day)
    db.commit()
    db.refresh(db_event_day)
    return db_event_day


def delete_event_day(db: Session, event_day_id: int) -> None:
    db_event_day = get_event_day(db, event_day_id)
    if db_event_day is not None:
        db.delete(db_event_day)
        db.commit()


def update_event_day(db: Session, event_day_id: int, updates: schemas.EventDayCreate) -> Optional[models.EventDay]:
    db_event_day = get_event_day(db, event_day_id)
    if db_event_day is None:
        return None
    for field, value in updates.dict(exclude_unset=True).items():
        setattr(db_event_day, field, value)
    db.add(db_event_day)
    db.commit()
    db.refresh(db_event_day)
    return db_event_day


# Assignment operations
def get_assignments(db: Session, skip: int = 0, limit: int = 100) -> List[models.Assignment]:
    return db.query(models.Assignment).offset(skip).limit(limit).all()


def get_assignment(db: Session, assignment_id: int) -> Optional[models.Assignment]:
    return db.query(models.Assignment).filter(models.Assignment.id == assignment_id).first()


def create_assignment(db: Session, assignment: schemas.AssignmentCreate) -> models.Assignment:
    db_assignment = models.Assignment(**assignment.dict())
    db.add(db_assignment)
    db.commit()
    db.refresh(db_assignment)
    return db_assignment


def delete_assignment(db: Session, assignment_id: int) -> None:
    db_assignment = get_assignment(db, assignment_id)
    if db_assignment is not None:
        db.delete(db_assignment)
        db.commit()


def update_assignment(db: Session, assignment_id: int, updates: schemas.AssignmentCreate) -> Optional[models.Assignment]:
    db_assignment = get_assignment(db, assignment_id)
    if db_assignment is None:
        return None
    for field, value in updates.dict(exclude_unset=True).items():
        setattr(db_assignment, field, value)
    db.add(db_assignment)
    db.commit()
    db.refresh(db_assignment)
    return db_assignment