"""API routes for managing volunteers."""

from typing import List

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from .. import crud, schemas, database


router = APIRouter(
    prefix="/volunteers",
    tags=["volunteers"],
)


@router.post("/", response_model=schemas.Volunteer, status_code=status.HTTP_201_CREATED)
def create_volunteer(volunteer: schemas.VolunteerCreate, db: Session = Depends(database.get_db)):
    """Create a new volunteer entry."""
    return crud.create_volunteer(db=db, volunteer=volunteer)


@router.get("/", response_model=List[schemas.Volunteer])
def read_volunteers(skip: int = 0, limit: int = 100, db: Session = Depends(database.get_db)):
    """Retrieve a list of volunteers."""
    volunteers = crud.get_volunteers(db, skip=skip, limit=limit)
    return volunteers


@router.get("/{volunteer_id}", response_model=schemas.Volunteer)
def read_volunteer(volunteer_id: int, db: Session = Depends(database.get_db)):
    """Retrieve a single volunteer by ID."""
    db_volunteer = crud.get_volunteer(db, volunteer_id=volunteer_id)
    if db_volunteer is None:
        raise HTTPException(status_code=404, detail="Volunteer not found")
    return db_volunteer


@router.put("/{volunteer_id}", response_model=schemas.Volunteer)
def update_volunteer(volunteer_id: int, updates: schemas.VolunteerCreate, db: Session = Depends(database.get_db)):
    """Update an existing volunteer.  Fields that are not provided will be left unchanged."""
    db_volunteer = crud.update_volunteer(db, volunteer_id=volunteer_id, updates=updates)
    if db_volunteer is None:
        raise HTTPException(status_code=404, detail="Volunteer not found")
    return db_volunteer


@router.delete("/{volunteer_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_volunteer(volunteer_id: int, db: Session = Depends(database.get_db)):
    """Delete a volunteer from the system."""
    crud.delete_volunteer(db, volunteer_id=volunteer_id)
    return None