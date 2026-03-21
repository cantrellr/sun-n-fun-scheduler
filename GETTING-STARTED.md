# Getting Started

## Purpose

This guide gets a Windows user from clone to first successful launch and export
with the least amount of setup friction.

## Prerequisites

- Windows 10 or newer
- PowerShell 5.1 or PowerShell 7 on Windows
- Local permission to run PowerShell scripts
- Microsoft Excel desktop if you need `.xlsx` export

## First Launch

### Option 1: Double-click launch

Run:

`Start-SunNFunScheduler.cmd`

This is the preferred first-run path because it launches PowerShell in STA mode,
which is the safest mode for the Windows Forms UI used by this repository.

### Option 2: PowerShell launch

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Start-SunNFunScheduler.ps1
```

## First Export

To generate exports without opening the desktop UI:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\Export-SunNFunSchedule.ps1
```

Outputs are written to the `exports` directory.

## Reset the Working Dataset

If the working JSON file needs to be restored:

```powershell
.\Reset-SeedData.ps1
```

## Recommended First Validation

1. Launch the app.
2. Confirm the event day list loads.
3. Open the volunteer summary view.
4. Run CSV/HTML export.
5. If Excel is installed, run Excel export.
6. Confirm expected output files appear in `exports`.

## Troubleshooting

### GUI does not open correctly

- Use `Start-SunNFunScheduler.cmd` instead of launching the PowerShell script
  from a non-STA session.
- Confirm you are running on Windows, not Linux, macOS, or WSL.

### Excel export fails

- Verify Microsoft Excel desktop is installed.
- Use CSV/HTML export if Excel is unavailable.

### Script execution is blocked

- Run PowerShell and use:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
```

## Next Reading

- `README.md` for repository overview and operating model
- `DEVELOPMENT.md` for architecture and maintenance notes
- `docs/Mobile-App-Roadmap.md` for future-state direction
