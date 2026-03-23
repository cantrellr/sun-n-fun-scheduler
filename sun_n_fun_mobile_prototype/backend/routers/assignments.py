"""API routes for managing assignments between volunteers and event days."""

from typing import List

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from .. import crud, schemas, database


router = APIRouter(
    prefix="/assignments",
    tags=["assignments"],
)


@router.post("/", response_model=schemas.Assignment, status_code=status.HTTP_201_CREATED)
def create_assignment(assignment: schemas.AssignmentCreate, db: Session = Depends(database.get_db)):
    """Assign a volunteer to an event day."""
    # TODO: Add capacity and conflict checks here.
    return crud.create_assignment(db=db, assignment=assignment)


@router.get("/", response_model=List[schemas.Assignment])
def read_assignments(skip: int = 0, limit: int = 100, db: Session = Depends(database.get_db)):
    """Retrieve all assignments."""
    return crud.get_assignments(db, skip=skip, limit=limit)


@router.get("/{assignment_id}", response_model=schemas.Assignment)
def read_assignment(assignment_id: int, db: Session = Depends(database.get_db)):
    """Retrieve a single assignment by ID."""
    db_assignment = crud.get_assignment(db, assignment_id=assignment_id)
    if db_assignment is None:
        raise HTTPException(status_code=404, detail="Assignment not found")
    return db_assignment


@router.put("/{assignment_id}", response_model=schemas.Assignment)
def update_assignment(assignment_id: int, updates: schemas.AssignmentCreate, db: Session = Depends(database.get_db)):
    """Update an existing assignment."""
    db_assignment = crud.update_assignment(db, assignment_id=assignment_id, updates=updates)
    if db_assignment is None:
        raise HTTPException(status_code=404, detail="Assignment not found")
    return db_assignment


@router.delete("/{assignment_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_assignment(assignment_id: int, db: Session = Depends(database.get_db)):
    """Delete an assignment from the system."""
    crud.delete_assignment(db, assignment_id=assignment_id)
    return None