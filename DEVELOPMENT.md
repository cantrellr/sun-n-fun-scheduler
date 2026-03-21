# Development Guide

## Repository Intent

This repository is a pragmatic Windows desktop prototype for schedule management.
It prioritizes operational usefulness and low setup overhead over architectural
purity.

## Current Architecture

### Application style

- single-script PowerShell application
- Windows Forms desktop GUI
- JSON-backed persistence
- embedded export logic for CSV, HTML, and Excel

### Main execution paths

- `Start-SunNFunScheduler.ps1`
  - loads the working schedule JSON
  - initializes the Windows Forms UI
  - supports interactive editing and export
- `Export-SunNFunSchedule.ps1`
  - delegates to the main script in export-only mode
- `Reset-SeedData.ps1`
  - restores the working dataset from the seed file

## Data Files

### Working file

`data/SunNFun-2026-Schedule.json`

This file is the active schedule used by the application.

### Seed file

`data/SunNFun-2026-Schedule.seed.json`

This file is the reset baseline and should be treated as immutable unless the
baseline schedule itself changes.

## Execution Constraints

### Windows-only UI

The desktop application uses Windows Forms and therefore requires Windows.

### STA requirement

The UI should run in Single-Threaded Apartment mode. The `.cmd` launcher handles
that requirement explicitly.

### Excel COM dependency

Excel export relies on COM automation. This means:

- Excel must be installed locally
- the machine must allow COM automation
- export failures should be expected in restricted environments

CSV and HTML exports are the resilient fallback path.

## Important Maintenance Risks

### No concurrency model

There is no locking, merge strategy, or optimistic concurrency. If multiple users
edit separate copies of the working file, manual reconciliation is required.

### No schema versioning

The JSON structure does not currently include a schema version field or migration
logic. Any future structural change should add a versioning approach before the
format evolves further.

### Sensitive data

Volunteer names, phone numbers, email addresses, and notes are present in the
working data. Logs, screenshots, exports, and issues should be handled carefully.

## Recommended Manual Validation

When making nontrivial changes, validate the parts affected by the change.

### Minimum validation set

1. Launch the GUI from `Start-SunNFunScheduler.cmd`.
2. Load at least one day view successfully.
3. Open the volunteer summary view.
4. Save the working file if data-path logic changed.
5. Run CSV/HTML export if export logic changed.
6. Run Excel export if Excel-specific logic changed and Excel is available.
7. Run `Reset-SeedData.ps1` if data reset logic changed.

## Code Organization Improvements Worth Doing Next

The next maintainability step is to split the main script into smaller units.

### Good candidate separations

- data loading and saving
- data transformation helpers
- export services
- UI composition and event handlers
- validation helpers

## Current Quality Gates

- GitHub Actions workflow for PowerShell lint and test execution
- PSScriptAnalyzer configuration for repository scripts
- Pester tests covering syntax and repository integrity checks

## Suggested Future Quality Gates

- Pester tests for non-UI functions
- schema validation for JSON files
- documented release/versioning process

## Migration Direction

This repository already has a clear boundary between current prototype behavior
and future product shape.

Recommended eventual transition:

1. preserve JSON as an import/export seed format
2. normalize data into a database-backed model
3. move scheduling rules into a service layer
4. build a browser-first responsive UI
5. add authentication, authorization, and audit logging

## Development Principles For This Repo

- keep operational workflows simple
- do not hide Windows or Excel constraints
- avoid accidental commits of generated exports
- preserve the seed-data recovery path
- treat data-bearing files as sensitive operational content
