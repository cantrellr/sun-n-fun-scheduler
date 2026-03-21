# Sun 'n Fun Scheduler

![Project Status](https://img.shields.io/badge/status-prototype-orange)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2B-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207.x-5391FE?logo=powershell&logoColor=white)
![Data Format](https://img.shields.io/badge/data-JSON-lightgrey)
![License](https://img.shields.io/badge/license-proprietary-red)

Windows PowerShell scheduler and export tool for managing the 2026 Sun 'n Fun EAA 791 volunteer schedule. This repository is intentionally lightweight: a desktop GUI for day-by-day schedule management, JSON as the system of record, and coordinator-friendly exports for operations use.

## Status

| Area | Current State |
| --- | --- |
| Product maturity | Prototype with real operational utility |
| Platform support | Windows only |
| Primary interface | Windows Forms desktop GUI |
| Data persistence | Local JSON file |
| Multi-user support | Not supported |
| Excel export | Supported when Microsoft Excel is installed |
| CSV/HTML export | Supported without Excel |
| Automated CI | GitHub Actions workflow for lint and tests |

## Quick Links

- [Getting Started](GETTING-STARTED.md)
- [Development Guide](DEVELOPMENT.md)
- [Mobile App Roadmap](docs/Mobile-App-Roadmap.md)
- [Contributing](CONTRIBUTING.md)
- [Security](SECURITY.md)
- [Changelog](CHANGELOG.md)

## Overview

This project exists to make volunteer schedule management faster and less error-prone for a specific event workflow:

- load prepared event-day and assignment data from JSON
- review staffing levels by date
- edit volunteer details directly in a desktop grid
- move volunteers between event days
- save updates back to the working schedule file
- generate coordinator exports in CSV, HTML, and optionally Excel

The current implementation is optimized for speed of use on a Windows machine, not for multi-user collaboration or cloud deployment.

## Core Features

- Day-by-day schedule view with headcount context
- Direct row editing for volunteer assignment details
- Move-selected workflow for reassigning volunteers across dates
- Save-to-JSON persistence for the current working schedule
- Export-only mode for generating fresh coordinator outputs without opening the GUI
- Reset script for restoring the working schedule from the seed file
- Volunteer summary view derived from current assignments
- Excel workbook export for users with Microsoft Excel installed
- CSV and HTML export path for environments without Excel

## Repository Layout

| Path | Purpose |
| --- | --- |
| `Start-SunNFunScheduler.ps1` | Main Windows Forms application and export engine |
| `Start-SunNFunScheduler.cmd` | Double-click launcher that starts PowerShell in STA mode |
| `Export-SunNFunSchedule.ps1` | Export-only entry point |
| `Reset-SeedData.ps1` | Restores the working JSON file from the baseline seed file |
| `data/SunNFun-2026-Schedule.json` | Working schedule data used by the app |
| `data/SunNFun-2026-Schedule.seed.json` | Baseline seed data for reset and recovery |
| `exports/` | Generated exports and sample workbook artifacts |
| `docs/Mobile-App-Roadmap.md` | Future-state product roadmap |
| `GETTING-STARTED.md` | Fast setup and first-run guide |
| `DEVELOPMENT.md` | Architecture, execution, and maintenance notes |

## Requirements

### Mandatory

- Windows 10 or newer
- PowerShell 5.1 or PowerShell 7 on Windows
- Permission to run local PowerShell scripts

### Optional

- Microsoft Excel desktop installed if `.xlsx` export is required

## First Run

### Fastest option

Double-click:

`Start-SunNFunScheduler.cmd`

### PowerShell option

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Start-SunNFunScheduler.ps1
```

The launcher exists because the GUI uses Windows Forms and behaves best when PowerShell starts in Single-Threaded Apartment mode.

## Common Workflows

### Launch the desktop scheduler

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Start-SunNFunScheduler.ps1
```

### Generate exports without opening the GUI

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Export-SunNFunSchedule.ps1
```

### Generate only CSV and HTML outputs

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Export-SunNFunSchedule.ps1 -ExportFormat CsvHtml
```

### Generate only Excel output

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Export-SunNFunSchedule.ps1 -ExportFormat Excel
```

### Reset the working schedule back to baseline

```powershell
.\Reset-SeedData.ps1
```

## Script Parameters

### `Start-SunNFunScheduler.ps1`

| Parameter | Description |
| --- | --- |
| `DataFile` | Path to the working JSON schedule file |
| `ExportDirectory` | Output directory for generated exports |
| `ExportOnly` | Runs export logic without showing the GUI |
| `ExportFormat` | `CsvHtml`, `Excel`, or `All` |

### `Export-SunNFunSchedule.ps1`

Wrapper for the same export settings as the main script, intended for direct export operations.

## Export Outputs

Depending on the selected export mode and local environment, the repository can produce:

- coordinator summary CSV
- volunteer directory CSV
- day-specific CSV files
- coordinator summary HTML report
- Excel workbook with summary, directory, day sheets, and raw assignment data

### Excel export behavior

Excel export is implemented through COM automation. That means:

- it is Windows-only
- it depends on a locally installed Excel desktop application
- it is more fragile than the CSV/HTML path if Office is missing or restricted

If Excel export fails, CSV and HTML exports remain the recommended fallback path.

## Data Model

The current JSON model is intentionally simple.

### Top-level structure

- `eventName`
- `generatedAtUtc`
- `sourceWorkbook`
- `days[]`
- `assignments[]`

### `days[]`

Each day record contains scheduling context such as:

- date
- display label
- event phase
- operating hours
- status message

### `assignments[]`

Each assignment record contains volunteer and day-specific details such as:

- assignment ID
- date
- volunteer name
- phone
- email
- shirt size
- camping flag
- notes
- original signup text
- status

### Current repository dataset

- 13 day records
- 55+ assignment records in the working file
- 21 unique volunteers in the original baseline set
- event phases covering before-event, during-event, and after-event operations
- a flagged review day for April 21 because updated hours were not provided

## Operating Model

This repository currently uses a practical prototype model:

- the assignment list is the source of truth
- volunteer directory data is derived from assignments
- the desktop GUI edits the working JSON directly
- reset behavior is file-based, not database-based

That approach keeps the implementation easy to understand and fast to operate, but it also defines the current limits.

## Known Limitations

- Windows-only due to Windows Forms and Excel COM usage
- No multi-user locking or concurrency protection
- No role-based access control
- No audit trail of who changed what and when
- No GUI automation coverage yet
- No schema versioning or migration system for data format changes
- Working data contains volunteer contact information and should be treated as sensitive

## Recommended Usage Practices

- Keep `data/SunNFun-2026-Schedule.seed.json` unchanged as the recovery baseline
- Use `data/SunNFun-2026-Schedule.json` as the active working file
- Treat exported files as operational artifacts and avoid committing unnecessary generated output
- Prefer CSV/HTML export when Excel is unavailable or unreliable
- Back up the working JSON file before major editing sessions
- Avoid concurrent edits by multiple people on copies of the same working file

## Repository Standards

This repository now includes foundational project hygiene files for maintainability:

- explicit license and usage terms
- contributor workflow guidance
- security reporting guidance
- changelog structure
- editor and line-ending normalization files
- detailed getting-started and development documentation
- GitHub Actions lint and test workflow
- baseline Pester repository tests
- baseline PSScriptAnalyzer configuration

## Development Notes

The application code is written as a single PowerShell script with embedded Windows Forms UI and export logic. That is acceptable for this prototype stage, but the next structural improvement would be separating the repository into:

- UI concerns
- schedule data access and validation
- export services
- reusable helper functions
- tests for non-UI logic

More detail is in [DEVELOPMENT.md](DEVELOPMENT.md).

## Roadmap Direction

The logical next phase is not a larger desktop script. It is a small service-backed scheduling platform with:

- normalized volunteer, day, and assignment entities
- audit logging
- search and filtering
- multi-user editing
- mobile-friendly operations workflow
- future offline-friendly field coordination options

See [docs/Mobile-App-Roadmap.md](docs/Mobile-App-Roadmap.md) for the current roadmap outline.

## Contributing

Contributions should follow the guidance in [CONTRIBUTING.md](CONTRIBUTING.md). At minimum:

- keep changes small and reviewable
- preserve Windows compatibility
- manually validate launch, export, and reset paths when relevant
- avoid accidental commits of sensitive or generated data

## Security and Data Handling

This repository contains schedule and contact information. Treat all working data and generated exports as operationally sensitive. Review [SECURITY.md](SECURITY.md) before sharing issues, logs, screenshots, or exported files.

## License

This repository is currently provided under a proprietary, all-rights-reserved license. See [LICENSE](LICENSE).
