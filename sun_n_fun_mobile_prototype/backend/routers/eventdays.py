"""API routes for managing event days."""

from typing import List

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from .. import crud, schemas, database


router = APIRouter(
    prefix="/eventdays",
    tags=["eventdays"],
)


@router.post("/", response_model=schemas.EventDay, status_code=status.HTTP_201_CREATED)
def create_event_day(event_day: schemas.EventDayCreate, db: Session = Depends(database.get_db)):
    """Create a new event day."""
    return crud.create_event_day(db=db, event_day=event_day)


@router.get("/", response_model=List[schemas.EventDay])
def read_event_days(skip: int = 0, limit: int = 100, db: Session = Depends(database.get_db)):
    """Retrieve all event days."""
    return crud.get_event_days(db, skip=skip, limit=limit)


@router.get("/{event_day_id}", response_model=schemas.EventDay)
def read_event_day(event_day_id: int, db: Session = Depends(database.get_db)):
    """Retrieve a single event day by ID."""
    db_event_day = crud.get_event_day(db, event_day_id=event_day_id)
    if db_event_day is None:
        raise HTTPException(status_code=404, detail="Event day not found")
    return db_event_day


@router.put("/{event_day_id}", response_model=schemas.EventDay)
def update_event_day(event_day_id: int, updates: schemas.EventDayCreate, db: Session = Depends(database.get_db)):
    """Update an existing event day."""
    db_event_day = crud.update_event_day(db, event_day_id=event_day_id, updates=updates)
    if db_event_day is None:
        raise HTTPException(status_code=404, detail="Event day not found")
    return db_event_day


@router.delete("/{event_day_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_event_day(event_day_id: int, db: Session = Depends(database.get_db)):
    """Delete an event day from the system."""
    crud.delete_event_day(db, event_day_id=event_day_id)
    return None