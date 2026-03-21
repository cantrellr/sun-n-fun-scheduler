$seed = Join-Path $PSScriptRoot 'data\SunNFun-2026-Schedule.seed.json'
$working = Join-Path $PSScriptRoot 'data\SunNFun-2026-Schedule.json'

if (-not (Test-Path -LiteralPath $seed)) {
    throw "Seed file not found: $seed"
}

Copy-Item -LiteralPath $seed -Destination $working -Force
Write-Host "Working data reset to seed: $working"
