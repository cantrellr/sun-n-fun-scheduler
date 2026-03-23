"""Entry point for the FastAPI application.

This module creates the FastAPI instance, includes routers for the core
entities, configures CORS and ensures database tables are created.  When
executed as a script (e.g. via `uvicorn main:app --reload`) it starts the
development server.  In production, the app should be served with a
production‑grade ASGI server (e.g. uvicorn or gunicorn with uvicorn workers).
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .database import Base, engine
from .routers import volunteers, eventdays, assignments

# Create database tables on startup.  In a more mature application this
# might be handled by Alembic migrations instead of automatic creation.
Base.metadata.create_all(bind=engine)

app = FastAPI(
    title="Sun 'n Fun Scheduler API",
    description="Backend API for managing volunteers, event days and assignments.",
    version="0.1.0",
)

# Allow cross‑origin requests from any origin during development.  Restrict this in production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers for each entity
app.include_router(volunteers.router)
app.include_router(eventdays.router)
app.include_router(assignments.router)


@app.get("/")
def read_root():
    """Basic root endpoint to verify the API is up."""
    return {"message": "Welcome to the Sun 'n Fun Scheduler API"}