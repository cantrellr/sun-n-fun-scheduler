# Mobile App Roadmap

## Recommended target architecture

### Backend
- ASP.NET Core Web API or FastAPI
- SQLite for single-node use, PostgreSQL for team/multi-user use
- REST API first, with a clean domain model

### Frontend
- Responsive web app first
- Then wrap with .NET MAUI, Flutter, or React Native if a true app store mobile app is needed

## Core entities

### Volunteer
- volunteerId
- firstName
- lastName
- phone
- email
- shirtSize
- camping
- notes
- emergencyContact
- tags / skills

### EventDay
- eventDayId
- date
- phase
- hours
- status
- capacity
- role requirements

### Assignment
- assignmentId
- volunteerId
- eventDayId
- checkInStatus
- role
- notes
- source
- createdBy
- updatedBy
- updatedAtUtc

### AuditEvent
- auditEventId
- action
- entityType
- entityId
- beforeJson
- afterJson
- actor
- createdAtUtc

## MVP feature set
- day-based scheduling
- drag/drop or tap-to-move assignments
- volunteer lookup
- search and filters
- bulk move
- bulk text / email export
- check-in status
- printable coordinator views
- conflict flagging
- admin/editor/viewer roles

## Smart enhancements
- automatic conflict detection
- QR or badge check-in
- text reminders
- role coverage dashboard
- weather / staffing alert banner
- offline-capable field mode

## Migration path from this prototype
1. Keep the JSON model as a seed import format
2. Normalize into database tables
3. Expose API endpoints
4. Build a browser-first mobile UI
5. Add authentication and audit
6. Add import from form submissions or spreadsheet drops
