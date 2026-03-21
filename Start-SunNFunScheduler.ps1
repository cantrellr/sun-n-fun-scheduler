param(
    [string]$DataFile = (Join-Path $PSScriptRoot 'data\SunNFun-2026-Schedule.json'),
    [string]$ExportDirectory = (Join-Path $PSScriptRoot 'exports'),
    [switch]$ExportOnly,
    [ValidateSet('CsvHtml', 'Excel', 'All')]
    [string]$ExportFormat = 'All'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure we're running in a Single-Threaded Apartment required by Windows Forms
try {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-Warning "PowerShell is not running in STA mode. Start the script using PowerShell with -STA (the launcher cmd now sets -STA). UI may freeze otherwise."
    }
}
catch {
    # ignore environments where ApartmentState isn't available
}

function Get-ScheduleData {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Data file not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $data = $raw | ConvertFrom-Json
    return $data
}

function Save-ScheduleData {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $Data.generatedAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $json = $Data | ConvertTo-Json -Depth 20
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Get-DaySummaries {
    param([Parameter(Mandatory)]$Data)

    $counts = @{}
    foreach ($assignment in @($Data.assignments)) {
        if (-not [string]::IsNullOrWhiteSpace($assignment.date)) {
            if (-not $counts.ContainsKey($assignment.date)) {
                $counts[$assignment.date] = 0
            }
            $counts[$assignment.date]++
        }
    }

    return @(
        foreach ($day in @($Data.days | Sort-Object date)) {
            $dayCount = if ($counts.ContainsKey($day.date)) { $counts[$day.date] } else { 0 }
            [pscustomobject]@{
                Date        = $day.date
                Label       = $day.label
                DisplayName = $day.displayName
                DayName     = $day.dayName
                Phase       = $day.phase
                Hours       = $day.hours
                Status      = $day.status
                Headcount   = $dayCount
                DisplayText = "$($day.label) | $($day.phase) | $($day.hours) | Headcount: $dayCount"
            }
        }
    )
}

function Format-DayListLabel {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$Hours
    )

    $normalizedHours = [string]$Hours
    if (-not [string]::IsNullOrWhiteSpace($normalizedHours)) {
        $normalizedHours = $normalizedHours.Trim().ToLowerInvariant()
        $normalizedHours = $normalizedHours -replace '\s+', ' '
        $normalizedHours = $normalizedHours -replace '(\d)\s+(am|pm)', '$1$2'
        $normalizedHours = $normalizedHours -replace '\s*-\s*', ' - '
    }

    return "$Label | $normalizedHours"
}

function Get-DayDefinition {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Date
    )

    return @($Data.days | Where-Object { $_.date -eq $Date })[0]
}

function Get-DayAssignments {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Date
    )

    return @($Data.assignments | Where-Object { $_.date -eq $Date } | Sort-Object volunteerName, email, phone)
}

function Get-NextAssignmentId {
    param([Parameter(Mandatory)]$Data)

    $numbers = @()
    foreach ($assignment in @($Data.assignments)) {
        if ($assignment.assignmentId -match '^ASN-(\d+)$') {
            $numbers += [int]$Matches[1]
        }
    }

    $nextNumber = if ($numbers.Count -gt 0) { ($numbers | Measure-Object -Maximum).Maximum + 1 } else { 1 }
    return ('ASN-' + ([int]$nextNumber).ToString('0000'))
}

function New-AssignmentsDataTable {
    $table = New-Object System.Data.DataTable 'Assignments'
    [void]$table.Columns.Add('AssignmentId', [string])
    [void]$table.Columns.Add('Volunteer', [string])
    [void]$table.Columns.Add('Phone', [string])
    [void]$table.Columns.Add('Email', [string])
    [void]$table.Columns.Add('Shirt', [string])
    [void]$table.Columns.Add('Camping', [string])
    [void]$table.Columns.Add('Notes', [string])
    [void]$table.Columns.Add('Original Signup', [string])
    [void]$table.Columns.Add('Status', [string])
    return ,$table
}

function Convert-AssignmentsToDataTable {
    param([Parameter(Mandatory)]$Assignments)
    if (-not $Assignments) { $Assignments = @() }

    $table = New-AssignmentsDataTable
    if (-not $table) {
        $table = New-Object System.Data.DataTable 'Assignments'
        [void]$table.Columns.Add('AssignmentId', [string])
        [void]$table.Columns.Add('Volunteer', [string])
        [void]$table.Columns.Add('Phone', [string])
        [void]$table.Columns.Add('Email', [string])
        [void]$table.Columns.Add('Shirt', [string])
        [void]$table.Columns.Add('Camping', [string])
        [void]$table.Columns.Add('Notes', [string])
        [void]$table.Columns.Add('Original Signup', [string])
        [void]$table.Columns.Add('Status', [string])
    }

    foreach ($assignment in @($Assignments)) {
        if (-not $assignment) { continue }
        $row = $table.NewRow()
        $row['AssignmentId']    = [string]$assignment.assignmentId
        $row['Volunteer']       = [string]$assignment.volunteerName
        $row['Phone']           = [string]$assignment.phone
        $row['Email']           = [string]$assignment.email
        $row['Shirt']           = [string]$assignment.shirt
        $row['Camping']         = [string]$assignment.camping
        $row['Notes']           = [string]$assignment.notes
        $row['Original Signup'] = [string]$assignment.originalSignup
        $row['Status']          = [string]$assignment.status
        [void]$table.Rows.Add($row)
    }

    return ,$table
}

function New-VolunteerDirectoryDataTable {
    $table = New-Object System.Data.DataTable 'VolunteerDirectory'
    [void]$table.Columns.Add('Volunteer', [string])
    [void]$table.Columns.Add('Phone', [string])
    [void]$table.Columns.Add('Email', [string])
    [void]$table.Columns.Add('Shirt', [string])
    [void]$table.Columns.Add('Camping', [string])
    [void]$table.Columns.Add('Notes', [string])
    [void]$table.Columns.Add('Signed Up Dates', [string])
    return ,$table
}

function Convert-VolunteerDirectoryToDataTable {
    param([Parameter(Mandatory)]$Entries)

    $table = New-VolunteerDirectoryDataTable
    foreach ($entry in @($Entries)) {
        if (-not $entry) { continue }
        $row = $table.NewRow()
        $row['Volunteer'] = [string]$entry.Volunteer
        $row['Phone'] = [string]$entry.Phone
        $row['Email'] = [string]$entry.Email
        $row['Shirt'] = [string]$entry.Shirt
        $row['Camping'] = [string]$entry.Camping
        $row['Notes'] = [string]$entry.Notes
        $row['Signed Up Dates'] = [string]$entry.'Signed Up Dates'
        [void]$table.Rows.Add($row)
    }

    return ,$table
}

function Sync-DayFromGrid {
    param(
        [Parameter(Mandatory)][ref]$DataRef,
        [Parameter(Mandatory)][string]$Date,
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid
    )

    $data = $DataRef.Value
    $otherAssignments = @($data.assignments | Where-Object { $_.date -ne $Date })
    $rebuilt = New-Object System.Collections.Generic.List[object]

    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }

        $volunteer = [string]$row.Cells['Volunteer'].Value
        $phone = [string]$row.Cells['Phone'].Value
        $email = [string]$row.Cells['Email'].Value
        $shirt = [string]$row.Cells['Shirt'].Value
        $camping = [string]$row.Cells['Camping'].Value
        $notes = [string]$row.Cells['Notes'].Value
        $originalSignup = [string]$row.Cells['Original Signup'].Value
        $status = [string]$row.Cells['Status'].Value
        $assignmentId = [string]$row.Cells['AssignmentId'].Value

        if ([string]::IsNullOrWhiteSpace($volunteer) -and [string]::IsNullOrWhiteSpace($phone) -and [string]::IsNullOrWhiteSpace($email)) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($assignmentId)) {
            $assignmentId = Get-NextAssignmentId -Data $data
            $row.Cells['AssignmentId'].Value = $assignmentId
        }

        $rebuilt.Add([pscustomobject]@{
            assignmentId   = $assignmentId
            date           = $Date
            volunteerName  = $volunteer.Trim()
            phone          = $phone.Trim()
            email          = $email.Trim()
            shirt          = $shirt.Trim()
            camping        = $camping.Trim()
            notes          = $notes.Trim()
            originalSignup = $originalSignup.Trim()
            status         = $status.Trim()
        })
    }

    $combined = @()
    $combined += $otherAssignments
    $combined += $rebuilt
    $data.assignments = @($combined | Sort-Object date, volunteerName, email, phone)
    $DataRef.Value = $data
}

function Get-VolunteerDirectory {
    param([Parameter(Mandatory)]$Data)

    $map = @{}
    foreach ($assignment in @($Data.assignments)) {
        $name = [string]$assignment.volunteerName
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $key = $name.Trim().ToLowerInvariant()
        if (-not $map.ContainsKey($key)) {
            $map[$key] = [ordered]@{
                Volunteer = $name.Trim()
                Phone = [string]$assignment.phone
                Email = [string]$assignment.email
                Shirt = [string]$assignment.shirt
                Camping = [string]$assignment.camping
                Notes = [string]$assignment.notes
                Dates = New-Object System.Collections.Generic.List[string]
            }
        }

        $propertyMap = @{
            Phone   = 'phone'
            Email   = 'email'
            Shirt   = 'shirt'
            Camping = 'camping'
            Notes   = 'notes'
        }
        foreach ($field in 'Phone','Email','Shirt','Camping','Notes') {
            $propertyName = $propertyMap[$field]
            $propertyValue = [string]$assignment.PSObject.Properties[$propertyName].Value
            if ([string]::IsNullOrWhiteSpace($map[$key][$field]) -and -not [string]::IsNullOrWhiteSpace($propertyValue)) {
                $map[$key][$field] = $propertyValue
            }
        }

        if (-not $map[$key].Dates.Contains([string]$assignment.date)) {
            $map[$key].Dates.Add([string]$assignment.date)
        }
    }

    return @(
        foreach ($entry in $map.Values) {
            [pscustomobject]@{
                Volunteer = $entry.Volunteer
                Phone = $entry.Phone
                Email = $entry.Email
                Shirt = $entry.Shirt
                Camping = $entry.Camping
                Notes = $entry.Notes
                'Signed Up Dates' = (($entry.Dates | Sort-Object | ForEach-Object { ([datetime]$_).ToString('MMM dd') }) -join ', ')
            }
        }
    ) | Sort-Object Volunteer
}

function New-SheetName {
    param([Parameter(Mandatory)][string]$Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + @('[',']',':','*','?','/','\')
    $sheetName = $Name
    foreach ($char in $invalid) {
        $sheetName = $sheetName.Replace([string]$char, '')
    }
    if ($sheetName.Length -gt 31) {
        $sheetName = $sheetName.Substring(0,31)
    }
    return $sheetName
}

function Write-WorksheetHeader {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][int]$HeaderRow,
        [Parameter(Mandatory)][string[]]$Headers
    )

    $col = 1
    foreach ($header in $Headers) {
        $Worksheet.Cells.Item($HeaderRow, $col).Value2 = $header
        $col++
    }

    $headerRange = $Worksheet.Range($Worksheet.Cells.Item($HeaderRow, 1), $Worksheet.Cells.Item($HeaderRow, $Headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xD9EAF7
    $headerRange.Borders.Weight = 2
}

function AutoFit-UsedRange {
    param([Parameter(Mandatory)]$Worksheet)
    $Worksheet.UsedRange.Columns.AutoFit() | Out-Null
}

function Export-ScheduleCsvHtml {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$OutputDirectory
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $summaries = Get-DaySummaries -Data $Data
    $summaries |
        Select-Object Phase,
                      @{Name='Date';Expression={ ([datetime]$_.Date).ToString('ddd, MMM dd, yyyy') }},
                      Hours,
                      Headcount,
                      Status |
        Export-Csv -Path (Join-Path $OutputDirectory 'Coordinator_Summary.csv') -NoTypeInformation -Encoding UTF8

    $directory = Get-VolunteerDirectory -Data $Data
    $directory | Export-Csv -Path (Join-Path $OutputDirectory 'Volunteer_Directory.csv') -NoTypeInformation -Encoding UTF8

    foreach ($day in @($Data.days | Sort-Object date)) {
        $assignments = @(Get-DayAssignments -Data $Data -Date $day.date)
        $fileName = '{0}.csv' -f ($day.label -replace '\s+', '_')
        $assignments |
            Select-Object @{Name='Volunteer';Expression={$_.volunteerName}},
                          phone,
                          email,
                          shirt,
                          camping,
                          notes,
                          originalSignup,
                          status |
            Export-Csv -Path (Join-Path $OutputDirectory $fileName) -NoTypeInformation -Encoding UTF8
    }

    $style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #222; }
h1, h2 { margin-bottom: 8px; }
.meta { color: #555; margin-bottom: 10px; }
table { border-collapse: collapse; width: 100%; margin: 12px 0 24px; }
th, td { border: 1px solid #d0d7de; padding: 8px; text-align: left; vertical-align: top; }
th { background: #f3f6fa; }
.summary th { background: #1f4e78; color: white; }
.day-card { margin-top: 26px; page-break-inside: avoid; }
.status.warn { display:inline-block; background:#fff3cd; color:#7a5c00; border:1px solid #ffe08a; padding:6px 10px; border-radius:4px; margin-bottom:10px; }
.muted { color:#666; font-style: italic; }
.small { font-size: 0.92rem; color:#555; }
</style>
"@

    $summaryRows = foreach ($summary in $summaries) {
        "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>" -f `
            [System.Net.WebUtility]::HtmlEncode($summary.Phase),
            [System.Net.WebUtility]::HtmlEncode(([datetime]$summary.Date).ToString('ddd, MMM dd, yyyy')),
            [System.Net.WebUtility]::HtmlEncode($summary.Hours),
            [System.Net.WebUtility]::HtmlEncode([string]$summary.Headcount),
            [System.Net.WebUtility]::HtmlEncode($summary.Status)
    }

    $daySections = foreach ($day in @($Data.days | Sort-Object date)) {
        $assignments = @(Get-DayAssignments -Data $Data -Date $day.date)
        if ($assignments.Count -eq 0) {
            $rows = '<tr><td colspan="7" class="muted">No volunteers assigned</td></tr>'
        }
        else {
            $rows = foreach ($assignment in $assignments) {
                "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>" -f `
                    [System.Net.WebUtility]::HtmlEncode($assignment.volunteerName),
                    [System.Net.WebUtility]::HtmlEncode($assignment.phone),
                    [System.Net.WebUtility]::HtmlEncode($assignment.email),
                    [System.Net.WebUtility]::HtmlEncode($assignment.shirt),
                    [System.Net.WebUtility]::HtmlEncode($assignment.camping),
                    [System.Net.WebUtility]::HtmlEncode($assignment.notes),
                    [System.Net.WebUtility]::HtmlEncode($assignment.status)
            }
        }

        $statusBlock = if (-not [string]::IsNullOrWhiteSpace($day.status)) {
            '<div class="status warn">{0}</div>' -f [System.Net.WebUtility]::HtmlEncode($day.status)
        }
        else { '' }

        @"
<section class="day-card">
  <h2>$([System.Net.WebUtility]::HtmlEncode($day.displayName))</h2>
    <div class="meta">$([System.Net.WebUtility]::HtmlEncode($day.phase)) • $([System.Net.WebUtility]::HtmlEncode($day.hours)) • Headcount: $((@(Get-DayAssignments -Data $Data -Date $day.date)).Count)</div>
  $statusBlock
  <table>
    <thead><tr><th>Volunteer</th><th>Phone</th><th>Email</th><th>Shirt</th><th>Camping</th><th>Notes</th><th>Status</th></tr></thead>
    <tbody>
      $($rows -join [Environment]::NewLine)
    </tbody>
  </table>
</section>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Sun 'n Fun 2026 Coordinator Summary</title>
$style
</head>
<body>
<h1>$([System.Net.WebUtility]::HtmlEncode($Data.eventName))</h1>
<div class="small">Generated from the PowerShell scheduler data file.</div>

<h2>Coordinator Summary</h2>
<table class="summary">
<thead>
<tr><th>Phase</th><th>Date</th><th>Hours</th><th>Headcount</th><th>Status</th></tr>
</thead>
<tbody>
$($summaryRows -join [Environment]::NewLine)
</tbody>
</table>

$($daySections -join [Environment]::NewLine)
</body>
</html>
"@

    Set-Content -LiteralPath (Join-Path $OutputDirectory 'Sun_n_Fun_2026_Coordinator_Summary.html') -Value $html -Encoding UTF8
}

function Export-ScheduleExcel {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$OutputDirectory
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $outputPath = Join-Path $OutputDirectory 'Sun_n_Fun_2026_Coordinator_Workbook.xlsx'
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        $desiredNames = New-Object System.Collections.Generic.List[string]

        $summarySheet = $workbook.Worksheets.Item(1)
        $summarySheet.Name = 'Coordinator Summary'
        [void]$desiredNames.Add('Coordinator Summary')

        $summarySheet.Cells.Item(1,1).Value2 = $Data.eventName
        $summarySheet.Cells.Item(1,1).Font.Bold = $true
        $summarySheet.Cells.Item(1,1).Font.Size = 16

        Write-WorksheetHeader -Worksheet $summarySheet -HeaderRow 3 -Headers @('Phase','Date','Day','Hours','Headcount','Status')
        $row = 4
        foreach ($summary in (Get-DaySummaries -Data $Data)) {
            $summarySheet.Cells.Item($row,1).Value2 = $summary.Phase
            $summarySheet.Cells.Item($row,2).Value2 = [datetime]$summary.Date
            $summarySheet.Cells.Item($row,2).NumberFormat = 'ddd, mmm dd, yyyy'
            $summarySheet.Cells.Item($row,3).Value2 = $summary.DayName
            $summarySheet.Cells.Item($row,4).Value2 = $summary.Hours
            $summarySheet.Cells.Item($row,5).Value2 = [string]$summary.Headcount
            $summarySheet.Cells.Item($row,6).Value2 = $summary.Status
            $row++
        }
        $summarySheet.Application.ActiveWindow.SplitRow = 3
        $summarySheet.Application.ActiveWindow.FreezePanes = $true
        AutoFit-UsedRange -Worksheet $summarySheet

        $directorySheet = $workbook.Worksheets.Add()
        $directorySheet.Name = 'Volunteer Directory'
        [void]$desiredNames.Add('Volunteer Directory')
        $directorySheet.Cells.Item(1,1).Value2 = 'Volunteer Directory'
        $directorySheet.Cells.Item(1,1).Font.Bold = $true
        $directorySheet.Cells.Item(1,1).Font.Size = 14

        Write-WorksheetHeader -Worksheet $directorySheet -HeaderRow 3 -Headers @('Volunteer','Phone','Email','Shirt','Camping','Notes','Signed Up Dates')
        $row = 4
        foreach ($entry in (Get-VolunteerDirectory -Data $Data)) {
            $directorySheet.Cells.Item($row,1).Value2 = $entry.Volunteer
            $directorySheet.Cells.Item($row,2).Value2 = $entry.Phone
            $directorySheet.Cells.Item($row,3).Value2 = $entry.Email
            $directorySheet.Cells.Item($row,4).Value2 = $entry.Shirt
            $directorySheet.Cells.Item($row,5).Value2 = $entry.Camping
            $directorySheet.Cells.Item($row,6).Value2 = $entry.Notes
            $directorySheet.Cells.Item($row,7).Value2 = $entry.'Signed Up Dates'
            $row++
        }
        AutoFit-UsedRange -Worksheet $directorySheet

        foreach ($day in @($Data.days | Sort-Object date)) {
            $sheetName = New-SheetName -Name $day.label
            $sheet = $workbook.Worksheets.Add()
            $sheet.Name = $sheetName
            [void]$desiredNames.Add($sheetName)

            $sheet.Cells.Item(1,1).Value2 = '{0} - {1}' -f $day.displayName, $day.phase
            $sheet.Cells.Item(1,1).Font.Bold = $true
            $sheet.Cells.Item(1,1).Font.Size = 14
            $sheet.Cells.Item(3,1).Value2 = 'Date'
            $sheet.Cells.Item(3,2).Value2 = [datetime]$day.date
            $sheet.Cells.Item(3,2).NumberFormat = 'ddd, mmm dd, yyyy'
            $sheet.Cells.Item(3,4).Value2 = 'Hours'
            $sheet.Cells.Item(3,5).Value2 = $day.hours
            $sheet.Cells.Item(3,7).Value2 = 'Headcount'
            $sheet.Cells.Item(3,8).Value2 = [string](@(Get-DayAssignments -Data $Data -Date $day.date)).Count
            if (-not [string]::IsNullOrWhiteSpace($day.status)) {
                $sheet.Cells.Item(4,1).Value2 = 'Status'
                $sheet.Cells.Item(4,2).Value2 = $day.status
                $sheet.Range('A4:B4').Interior.Color = 0xCCFFFF
            }

            Write-WorksheetHeader -Worksheet $sheet -HeaderRow 6 -Headers @('Volunteer','Phone','Email','Shirt','Camping','Notes','Original Signup','Status')
            $row = 7
            foreach ($assignment in @(Get-DayAssignments -Data $Data -Date $day.date)) {
                $sheet.Cells.Item($row,1).Value2 = $assignment.volunteerName
                $sheet.Cells.Item($row,2).Value2 = $assignment.phone
                $sheet.Cells.Item($row,3).Value2 = $assignment.email
                $sheet.Cells.Item($row,4).Value2 = $assignment.shirt
                $sheet.Cells.Item($row,5).Value2 = $assignment.camping
                $sheet.Cells.Item($row,6).Value2 = $assignment.notes
                $sheet.Cells.Item($row,7).Value2 = $assignment.originalSignup
                $sheet.Cells.Item($row,8).Value2 = $assignment.status
                $row++
            }
            AutoFit-UsedRange -Worksheet $sheet
        }

        $rawSheet = $workbook.Worksheets.Add()
        $rawSheet.Name = 'Raw Assignments'
        [void]$desiredNames.Add('Raw Assignments')
        $rawSheet.Cells.Item(1,1).Value2 = 'Raw Assignments'
        $rawSheet.Cells.Item(1,1).Font.Bold = $true
        $rawSheet.Cells.Item(1,1).Font.Size = 14

        Write-WorksheetHeader -Worksheet $rawSheet -HeaderRow 3 -Headers @('AssignmentId','Date','Volunteer','Phone','Email','Shirt','Camping','Notes','Original Signup','Status')
        $row = 4
        foreach ($assignment in @($Data.assignments | Sort-Object date, volunteerName)) {
            $rawSheet.Cells.Item($row,1).Value2 = $assignment.assignmentId
            $rawSheet.Cells.Item($row,2).Value2 = [datetime]$assignment.date
            $rawSheet.Cells.Item($row,2).NumberFormat = 'yyyy-mm-dd'
            $rawSheet.Cells.Item($row,3).Value2 = $assignment.volunteerName
            $rawSheet.Cells.Item($row,4).Value2 = $assignment.phone
            $rawSheet.Cells.Item($row,5).Value2 = $assignment.email
            $rawSheet.Cells.Item($row,6).Value2 = $assignment.shirt
            $rawSheet.Cells.Item($row,7).Value2 = $assignment.camping
            $rawSheet.Cells.Item($row,8).Value2 = $assignment.notes
            $rawSheet.Cells.Item($row,9).Value2 = $assignment.originalSignup
            $rawSheet.Cells.Item($row,10).Value2 = $assignment.status
            $row++
        }
        AutoFit-UsedRange -Worksheet $rawSheet

        for ($i = $workbook.Worksheets.Count; $i -ge 1; $i--) {
            $worksheet = $workbook.Worksheets.Item($i)
            if ($desiredNames -notcontains [string]$worksheet.Name) {
                $worksheet.Delete()
            }
        }

        $workbook.SaveAs($outputPath)
        $workbook.Close($true)
        $excel.Quit()
    }
    catch {
        throw "Excel export failed. Confirm Microsoft Excel is installed on this Windows system. Details: $($_.Exception.Message)"
    }
    finally {
        if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Show-MoveDialog {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$CurrentDate
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Move to another day'
    $form.Width = 420
    $form.Height = 170
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Left = 15
    $label.Top = 20
    $label.Width = 360
    $label.Text = 'Choose the target day for the selected assignment(s):'
    $form.Controls.Add($label)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Left = 15
    $combo.Top = 50
    $combo.Width = 370
    $combo.DropDownStyle = 'DropDownList'
    $combo.DataSource = @($Data.days | Where-Object { $_.date -ne $CurrentDate } | Sort-Object date)
    $combo.DisplayMember = 'displayName'
    $combo.ValueMember = 'date'
    $form.Controls.Add($combo)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Left = 220
    $ok.Top = 85
    $ok.Width = 75
    $ok.Text = 'OK'
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Left = 310
    $cancel.Top = 85
    $cancel.Width = 75
    $cancel.Text = 'Cancel'
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $combo.SelectedItem) {
        return [string]$combo.SelectedValue
    }

    return $null
}

function Start-ScheduleGui {
    param(
        [Parameter(Mandatory)][string]$DataFile,
        [Parameter(Mandatory)][string]$ExportDirectory
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $script:ScheduleData = Get-ScheduleData -Path $DataFile
    $script:CurrentDate = $null
    $script:IsDirty = $false
    $script:SuppressDayListEvents = $false
    $script:DayListItems = @()
    $script:IsSummaryMode = $false

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Sun 'n Fun Scheduler Prototype"
    $form.Width = 1480
    $form.Height = 860
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 720)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Left = 12
    $titleLabel.Top = 12
    $titleLabel.Width = 1200
    $titleLabel.Height = 28
    $titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Text = $script:ScheduleData.eventName
    $titleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($titleLabel)

    $dataLabel = New-Object System.Windows.Forms.Label
    $dataLabel.Left = 14
    $dataLabel.Top = 42
    $dataLabel.Width = 1000
    $dataLabel.Height = 18
    $dataLabel.Text = "Data file: $DataFile"
    $dataLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($dataLabel)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Left = 12
    $split.Top = 70
    $split.Width = 1440
    $split.Height = 690
    $split.SplitterDistance = 360
    $split.IsSplitterFixed = $false
    $split.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($split)

    $dayListLabel = New-Object System.Windows.Forms.Label
    $dayListLabel.Left = 8
    $dayListLabel.Top = 8
    $dayListLabel.Width = 300
    $dayListLabel.Text = 'Event days'
    $split.Panel1.Controls.Add($dayListLabel)

    $btnSummary = New-Object System.Windows.Forms.Button
    $btnSummary.Left = 8
    $btnSummary.Top = 30
    $btnSummary.Width = 100
    $btnSummary.Height = 28
    $btnSummary.Text = 'Summary'
    $btnSummary.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $split.Panel1.Controls.Add($btnSummary)

    $dayList = New-Object System.Windows.Forms.ListBox
    $dayList.Left = 8
    $dayList.Top = 64
    $dayList.Width = 335
    $dayList.Height = 601
    $dayList.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dayList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $split.Panel1.Controls.Add($dayList)

    $selectedDayLabel = New-Object System.Windows.Forms.Label
    $selectedDayLabel.Left = 10
    $selectedDayLabel.Top = 8
    $selectedDayLabel.Width = 1000
    $selectedDayLabel.Height = 22
    $selectedDayLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $selectedDayLabel.Text = 'Select a day to begin'
    $selectedDayLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $split.Panel2.Controls.Add($selectedDayLabel)

    $selectedDayMeta = New-Object System.Windows.Forms.Label
    $selectedDayMeta.Left = 10
    $selectedDayMeta.Top = 34
    $selectedDayMeta.Width = 1000
    $selectedDayMeta.Height = 38
    $selectedDayMeta.Text = ''
    $selectedDayMeta.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $split.Panel2.Controls.Add($selectedDayMeta)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 10
    $grid.Top = 76
    $grid.Width = 1040
    $grid.Height = 580
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToDeleteRows = $true
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoGenerateColumns = $true
    $grid.AutoSizeColumnsMode = 'AllCells'
    $grid.EditMode = 'EditOnKeystrokeOrF2'
    $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $split.Panel2.Controls.Add($grid)

    # Use a BindingSource to improve DataGridView binding/rebinding behavior
    $gridBindingSource = New-Object System.Windows.Forms.BindingSource
    $grid.DataSource = $gridBindingSource

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Left = 10
    $buttonPanel.Top = 662
    $buttonPanel.Width = 1040
    $buttonPanel.Height = 34
    $buttonPanel.FlowDirection = 'LeftToRight'
    $buttonPanel.WrapContents = $false
    $buttonPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $split.Panel2.Controls.Add($buttonPanel)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add Row'
    $btnAdd.Width = 90
    $buttonPanel.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove Selected'
    $btnRemove.Width = 120
    $buttonPanel.Controls.Add($btnRemove)

    $btnMove = New-Object System.Windows.Forms.Button
    $btnMove.Text = 'Move Selected'
    $btnMove.Width = 110
    $buttonPanel.Controls.Add($btnMove)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Save'
    $btnSave.Width = 80
    $buttonPanel.Controls.Add($btnSave)

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = 'Export CSV/HTML'
    $btnExportCsv.Width = 120
    $buttonPanel.Controls.Add($btnExportCsv)

    $btnExportExcel = New-Object System.Windows.Forms.Button
    $btnExportExcel.Text = 'Export Excel'
    $btnExportExcel.Width = 100
    $buttonPanel.Controls.Add($btnExportExcel)

    $btnOpenExports = New-Object System.Windows.Forms.Button
    $btnOpenExports.Text = 'Open Exports Folder'
    $btnOpenExports.Width = 130
    $buttonPanel.Controls.Add($btnOpenExports)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'
    $btnRefresh.Width = 80
    $buttonPanel.Controls.Add($btnRefresh)

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready'
    [void]$statusStrip.Items.Add($statusLabel)
    $form.Controls.Add($statusStrip)

    function Set-StatusMessage {
        param([string]$Message)
        $statusLabel.Text = $Message
    }

    function Set-DayEditingEnabled {
        param([bool]$Enabled)

        $btnAdd.Enabled = $Enabled
        $btnRemove.Enabled = $Enabled
        $btnMove.Enabled = $Enabled
        $btnSave.Enabled = $Enabled
    }

    function Bind-GridTable {
        param([Parameter(Mandatory)][System.Data.DataTable]$Table)

        try {
            $grid.SuspendLayout()
            $gridBindingSource.DataSource = $null
            try { $grid.Columns.Clear() } catch { }
            try { $grid.DataSource = $null } catch { }

            $gridBindingSource.DataSource = $Table
            try { $grid.DataSource = $gridBindingSource } catch { }
            try { $gridBindingSource.ResetBindings($true) } catch { }
            try { $grid.Refresh(); $grid.Invalidate(); $grid.Update() } catch { }
        }
        finally {
            try { $grid.ResumeLayout() } catch { }
        }
    }

    function Show-VolunteerSummary {
        if ($script:CurrentDate -and -not $script:IsSummaryMode) {
            Save-CurrentDayInMemory
        }

        $directory = Get-VolunteerDirectory -Data $script:ScheduleData
        $table = Convert-VolunteerDirectoryToDataTable -Entries $directory
        Bind-GridTable -Table $table

        if ($grid.Columns['Volunteer']) { $grid.Columns['Volunteer'].Width = 180 }
        if ($grid.Columns['Phone']) { $grid.Columns['Phone'].Width = 110 }
        if ($grid.Columns['Email']) { $grid.Columns['Email'].Width = 190 }
        if ($grid.Columns['Shirt']) { $grid.Columns['Shirt'].Width = 60 }
        if ($grid.Columns['Camping']) { $grid.Columns['Camping'].Width = 70 }
        if ($grid.Columns['Notes']) { $grid.Columns['Notes'].Width = 280 }
        if ($grid.Columns['Signed Up Dates']) { $grid.Columns['Signed Up Dates'].Width = 140 }

        $selectedDayLabel.Text = 'Volunteer Summary'
        $selectedDayMeta.Text = "Volunteers: $($directory.Count)"
        $script:CurrentDate = $null
        $script:IsSummaryMode = $true
        Set-DayEditingEnabled -Enabled $false
        Set-StatusMessage 'Loaded volunteer summary.'
    }

    function Update-DayList {
        param([string]$PreserveDate)
        # Prevent SelectedIndexChanged re-entry while we update the list
        $script:SuppressDayListEvents = $true
        try {
            $script:DayListItems = @(Get-DaySummaries -Data $script:ScheduleData)
            $dayList.BeginUpdate()
            $dayList.Items.Clear()
            foreach ($item in $script:DayListItems) {
                [void]$dayList.Items.Add((Format-DayListLabel -Label ([string]$item.Label) -Hours ([string]$item.Hours)))
            }
            $dayList.EndUpdate()

            if ($PreserveDate) {
                for ($i = 0; $i -lt $script:DayListItems.Count; $i++) {
                    if ($script:DayListItems[$i].Date -eq $PreserveDate) {
                        if ($i -ge 0 -and $i -lt $dayList.Items.Count) {
                            try { $dayList.SelectedIndex = $i } catch { }
                        }
                        break
                    }
                }
            }
            elseif ($dayList.Items.Count -gt 0 -and $dayList.SelectedIndex -lt 0) {
                try { if ($dayList.Items.Count -gt 0) { $dayList.SelectedIndex = 0 } } catch { }
            }
        }
        finally {
            $script:SuppressDayListEvents = $false
        }
    }

    function Load-DayIntoGrid {
        param([Parameter(Mandatory)][string]$Date)

        $day = Get-DayDefinition -Data $script:ScheduleData -Date $Date
        $assignments = @(Get-DayAssignments -Data $script:ScheduleData -Date $Date)
        $table = Convert-AssignmentsToDataTable -Assignments $assignments

        Bind-GridTable -Table $table
        if ($grid.Columns['AssignmentId']) { $grid.Columns['AssignmentId'].Visible = $false }
        if ($grid.Columns['Volunteer'])     { $grid.Columns['Volunteer'].Width = 180 }
        if ($grid.Columns['Phone'])         { $grid.Columns['Phone'].Width = 110 }
        if ($grid.Columns['Email'])         { $grid.Columns['Email'].Visible = $false }
        if ($grid.Columns['Shirt'])         { $grid.Columns['Shirt'].Visible = $false }
        if ($grid.Columns['Camping'])       { $grid.Columns['Camping'].Visible = $false }
        if ($grid.Columns['Notes'])         { $grid.Columns['Notes'].Visible = $false }
        if ($grid.Columns['Original Signup']) { $grid.Columns['Original Signup'].Visible = $false }
        if ($grid.Columns['Status'])        { $grid.Columns['Status'].Visible = $false }

        $selectedDayLabel.Text = "$($day.displayName) - $($day.phase)"
        $meta = "Hours: $($day.hours)   |   Headcount: $($assignments.Count)"
        if (-not [string]::IsNullOrWhiteSpace($day.status)) {
            $meta = "$meta   |   Status: $($day.status)"
        }
        $selectedDayMeta.Text = $meta
        $script:CurrentDate = $Date
        $script:IsSummaryMode = $false
        Set-DayEditingEnabled -Enabled $true
        Set-StatusMessage "Loaded $($day.displayName)"
    }

    function Save-CurrentDayInMemory {
        if ($script:IsSummaryMode) { return }
        if (-not $script:CurrentDate) { return }
        # Prefer syncing from the DataTable bound to the BindingSource to avoid relying on UI row state
        try {
            $table = [System.Data.DataTable]$gridBindingSource.DataSource
            if ($table -and $table.Rows.Count -ge 0) {
                $rebuilt = New-Object System.Collections.Generic.List[object]
                foreach ($r in $table.Rows) {
                    $assignmentId = [string]$r['AssignmentId']
                    $volunteer = [string]$r['Volunteer']
                    $phone = [string]$r['Phone']
                    $email = [string]$r['Email']
                    $shirt = [string]$r['Shirt']
                    $camping = [string]$r['Camping']
                    $notes = [string]$r['Notes']
                    $originalSignup = [string]$r['Original Signup']
                    $status = [string]$r['Status']

                    if ([string]::IsNullOrWhiteSpace($volunteer) -and [string]::IsNullOrWhiteSpace($phone) -and [string]::IsNullOrWhiteSpace($email)) { continue }
                    if ([string]::IsNullOrWhiteSpace($assignmentId)) { $assignmentId = Get-NextAssignmentId -Data $script:ScheduleData }

                    $rebuilt.Add([pscustomobject]@{
                        assignmentId   = $assignmentId
                        date           = $script:CurrentDate
                        volunteerName  = $volunteer.Trim()
                        phone          = $phone.Trim()
                        email          = $email.Trim()
                        shirt          = $shirt.Trim()
                        camping        = $camping.Trim()
                        notes          = $notes.Trim()
                        originalSignup = $originalSignup.Trim()
                        status         = $status.Trim()
                    })
                }

                $data = $script:ScheduleData
                $existingCount = @($data.assignments | Where-Object { $_.date -eq $script:CurrentDate }).Count
                if ($rebuilt.Count -eq 0 -and $existingCount -gt 0) {
                    return
                }

                $otherAssignments = @($data.assignments | Where-Object { $_.date -ne $script:CurrentDate })
                $combined = @()
                $combined += $otherAssignments
                $combined += $rebuilt
                $data.assignments = @($combined | Sort-Object date, volunteerName, email, phone)
                $script:ScheduleData = $data
                return
            }
        }
        catch { }

        # Fallback to the original grid-based sync if no DataTable available
        Sync-DayFromGrid -DataRef ([ref]$script:ScheduleData) -Date $script:CurrentDate -Grid $grid
    }

    $grid.add_CellValueChanged({
        $script:IsDirty = $true
    })

    $grid.add_UserDeletedRow({
        $script:IsDirty = $true
    })

    $dayList.add_SelectedIndexChanged({
        if ($script:SuppressDayListEvents) { return }
        if ($dayList.SelectedIndex -lt 0) { return }
        if ($dayList.SelectedIndex -ge $script:DayListItems.Count) { return }

        $newDate = [string]$script:DayListItems[$dayList.SelectedIndex].Date
        if ($script:CurrentDate -and $script:CurrentDate -ne $newDate) {
            Save-CurrentDayInMemory
        }

        Load-DayIntoGrid -Date $newDate
        Update-DayList -PreserveDate $newDate
    })

    $btnSummary.add_Click({
        Show-VolunteerSummary
    })

    $btnAdd.add_Click({
        if ($script:IsSummaryMode) { return }
        if (-not $script:CurrentDate) { return }
        $table = [System.Data.DataTable]$gridBindingSource.DataSource
        $row = $table.NewRow()
        $row['AssignmentId'] = Get-NextAssignmentId -Data $script:ScheduleData
        [void]$table.Rows.Add($row)
        $grid.CurrentCell = $grid.Rows[$grid.Rows.Count - 2].Cells['Volunteer']
        $grid.BeginEdit($true)
        $script:IsDirty = $true
        Set-StatusMessage 'Added blank row.'
    })

    $btnRemove.add_Click({
        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Select one or more rows to remove.', 'Nothing selected') | Out-Null
            return
        }

        foreach ($selectedRow in @($grid.SelectedRows | Sort-Object Index -Descending)) {
            if (-not $selectedRow.IsNewRow) {
                $grid.Rows.Remove($selectedRow)
            }
        }
        $script:IsDirty = $true
        Set-StatusMessage 'Removed selected rows.'
        Update-DayList -PreserveDate $script:CurrentDate
    })

    $btnMove.add_Click({
        if ($script:IsSummaryMode) { return }
        if (-not $script:CurrentDate) { return }
        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Select one or more rows to move.', 'Nothing selected') | Out-Null
            return
        }

        Save-CurrentDayInMemory
        $targetDate = Show-MoveDialog -Data $script:ScheduleData -CurrentDate $script:CurrentDate
        if (-not $targetDate) { return }

        $selectedIds = @()
        foreach ($selectedRow in @($grid.SelectedRows)) {
            if (-not $selectedRow.IsNewRow) {
                $selectedIds += [string]$selectedRow.Cells['AssignmentId'].Value
            }
        }

        foreach ($assignment in @($script:ScheduleData.assignments | Where-Object { $selectedIds -contains $_.assignmentId })) {
            $assignment.date = $targetDate
        }

        Update-DayList -PreserveDate $script:CurrentDate
        Load-DayIntoGrid -Date $script:CurrentDate
        $script:IsDirty = $true
        $target = Get-DayDefinition -Data $script:ScheduleData -Date $targetDate
        Set-StatusMessage "Moved selected assignment(s) to $($target.displayName)."
    })

    $btnSave.add_Click({
        try {
            if ($script:IsSummaryMode) {
                Set-StatusMessage 'Summary view has no day-specific edits to save.'
                return
            }
            Save-CurrentDayInMemory
            Save-ScheduleData -Data $script:ScheduleData -Path $DataFile
            $script:IsDirty = $false
            Update-DayList -PreserveDate $script:CurrentDate
            Load-DayIntoGrid -Date $script:CurrentDate
            Set-StatusMessage "Saved changes to $DataFile"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Save failed') | Out-Null
            Set-StatusMessage 'Save failed.'
        }
    })

    $btnExportCsv.add_Click({
        try {
            Save-CurrentDayInMemory
            Save-ScheduleData -Data $script:ScheduleData -Path $DataFile
            Export-ScheduleCsvHtml -Data $script:ScheduleData -OutputDirectory $ExportDirectory
            $script:IsDirty = $false
            [System.Windows.Forms.MessageBox]::Show("CSV and HTML exports were written to:`n$ExportDirectory", 'Export complete') | Out-Null
            Set-StatusMessage 'CSV and HTML export complete.'
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Export failed') | Out-Null
            Set-StatusMessage 'CSV/HTML export failed.'
        }
    })

    $btnExportExcel.add_Click({
        try {
            Save-CurrentDayInMemory
            Save-ScheduleData -Data $script:ScheduleData -Path $DataFile
            Export-ScheduleExcel -Data $script:ScheduleData -OutputDirectory $ExportDirectory
            $script:IsDirty = $false
            [System.Windows.Forms.MessageBox]::Show("Excel export was written to:`n$ExportDirectory", 'Export complete') | Out-Null
            Set-StatusMessage 'Excel export complete.'
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Excel export failed') | Out-Null
            Set-StatusMessage 'Excel export failed.'
        }
    })

    $btnOpenExports.add_Click({
        if (-not (Test-Path -LiteralPath $ExportDirectory)) {
            New-Item -ItemType Directory -Path $ExportDirectory -Force | Out-Null
        }
        Start-Process explorer.exe $ExportDirectory
    })

    $btnRefresh.add_Click({
        if (-not $script:IsSummaryMode) {
            Save-CurrentDayInMemory
        }
        Update-DayList -PreserveDate $script:CurrentDate
        if ($script:IsSummaryMode) {
            Show-VolunteerSummary
        }
        elseif ($script:CurrentDate) {
            Load-DayIntoGrid -Date $script:CurrentDate
        }
        Set-StatusMessage 'Refreshed in-memory schedule.'
    })

    $form.add_FormClosing({
        if (-not $script:IsDirty) { return }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            'You have unsaved changes. Save before exiting?',
            'Unsaved changes',
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) {
            $_.Cancel = $true
            return
        }

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Save-CurrentDayInMemory
                Save-ScheduleData -Data $script:ScheduleData -Path $DataFile
                $script:IsDirty = $false
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Save failed') | Out-Null
                $_.Cancel = $true
            }
        }
    })

    Update-DayList
    if ($dayList.Items.Count -gt 0) {
        Load-DayIntoGrid -Date ([string]$script:DayListItems[0].Date)
        Update-DayList -PreserveDate ([string]$script:DayListItems[0].Date)
    }

    [void]$form.ShowDialog()
}

if ($ExportOnly) {
    $scheduleData = Get-ScheduleData -Path $DataFile

    switch ($ExportFormat) {
        'CsvHtml' {
            Export-ScheduleCsvHtml -Data $scheduleData -OutputDirectory $ExportDirectory
        }
        'Excel' {
            Export-ScheduleExcel -Data $scheduleData -OutputDirectory $ExportDirectory
        }
        default {
            Export-ScheduleCsvHtml -Data $scheduleData -OutputDirectory $ExportDirectory
            Export-ScheduleExcel -Data $scheduleData -OutputDirectory $ExportDirectory
        }
    }

    Write-Host "Export complete. Files written to: $ExportDirectory"
}
else {
    Start-ScheduleGui -DataFile $DataFile -ExportDirectory $ExportDirectory
}
