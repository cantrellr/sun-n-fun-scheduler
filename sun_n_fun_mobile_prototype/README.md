# Sun 'n Fun Mobile Scheduler Prototype

This repository contains a **proof‑of‑concept mobile scheduling system** designed to manage volunteers for multi‑day events.  It focuses on clean separation between the server API and the client UI so future mobile frameworks (like React Native, .NET MAUI or Flutter) can be layered on top without ripping out the core logic.

## Architecture overview

- **Backend** – A [FastAPI](https://fastapi.tiangolo.com/) service with SQLite persistence.  The service exposes REST endpoints for the core entities defined in the roadmap (volunteers, event days and assignments).  SQLAlchemy models mirror the domain model and a simple CRUD layer encapsulates database interaction.  Cross‑origin requests are enabled to make the API consumable from any frontend during development.
- **Frontend** – A minimal React‑based web page served as a static file under the `frontend/` folder.  It fetches data from the API and renders lists of event days and volunteers.  While intentionally simple, the structure demonstrates how a JavaScript client can consume the API.  The client could be replaced or wrapped by React Native/Expo, Flutter or .NET MAUI to produce a true mobile application.

The goal of this prototype is to lay the groundwork for the migration path described in the roadmap: starting with a JSON‑backed scheduler, normalising into tables, exposing an API and then building a browser‑first mobile UI.

## Getting started

### Running the backend

1. Change into the backend directory:

   ```bash
   cd sun_n_fun_mobile_prototype/backend
   ```

2. Install dependencies into a virtual environment (optional but recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. Launch the API server:

   ```bash
   uvicorn main:app --reload
   ```

4. The API will now listen on `http://localhost:8000`.  You can interact with the automatically generated interactive docs at `http://localhost:8000/docs` or `http://localhost:8000/redoc`.

### Running the frontend

1. In a separate terminal window, start a simple HTTP server from the `frontend` directory:

   ```bash
   cd sun_n_fun_mobile_prototype/frontend
   python -m http.server 8001
   ```

2. Open `http://localhost:8001` in your browser.  The page will load and issue fetch requests against the API on port 8000.  If you add volunteers or event days via the API, refresh the page to see the updated data.

### Next steps

This is a **minimal foundation** for a scheduling platform.  Future enhancements should include:

- Drag‑and‑drop or tap‑to‑move assignment management in the frontend.
- Role requirements, capacity and conflict detection on the server.
- User authentication and role‑based authorisation (admin, editor, viewer).
- Richer UI with search, filters and bulk operations.
- Packaging the frontend as a progressive web app (PWA) or wrapping it in React Native or MAUI for distribution via app stores.

Feel free to extend this prototype or adapt it to your preferred technology stack.  The clean separation of concerns should make it straightforward to evolve both the backend and frontend independently.