param(
    [string]$DataFile = (Join-Path $PSScriptRoot 'data\SunNFun-2026-Schedule.json'),
    [string]$ExportDirectory = (Join-Path $PSScriptRoot 'exports'),
    [ValidateSet('CsvHtml', 'Excel', 'All')]
    [string]$ExportFormat = 'All'
)

Set-ExecutionPolicy -Scope Process Bypass -Force
& (Join-Path $PSScriptRoot 'Start-SunNFunScheduler.ps1') -DataFile $DataFile -ExportDirectory $ExportDirectory -ExportOnly -ExportFormat $ExportFormat
