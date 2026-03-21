$script:repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPaths = @(
    (Join-Path $script:repoRoot 'Start-SunNFunScheduler.ps1'),
    (Join-Path $script:repoRoot 'Export-SunNFunSchedule.ps1'),
    (Join-Path $script:repoRoot 'Reset-SeedData.ps1')
)
$jsonPaths = @(
    (Join-Path $script:repoRoot 'data\SunNFun-2026-Schedule.json'),
    (Join-Path $script:repoRoot 'data\SunNFun-2026-Schedule.seed.json')
)

Describe 'Repository script quality' {
    It 'parses each PowerShell entry script without syntax errors' -ForEach $scriptPaths {
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($_, [ref]$tokens, [ref]$errors)

        $errors | Should -BeNullOrEmpty
    }

    It 'keeps the Windows launcher in STA mode' {
        $launcherPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Start-SunNFunScheduler.cmd'
        $launcherContent = Get-Content -LiteralPath $launcherPath -Raw

        $launcherContent | Should -Match '-STA'
    }
}

Describe 'Repository data integrity' {
    It 'loads each schedule JSON file successfully' -ForEach $jsonPaths {
        { Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json } | Should -Not -Throw
    }

    It 'contains the required top-level properties in each schedule file' -ForEach $jsonPaths {
        $data = Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json

        $data.PSObject.Properties.Name | Should -Contain 'eventName'
        $data.PSObject.Properties.Name | Should -Contain 'days'
        $data.PSObject.Properties.Name | Should -Contain 'assignments'
    }

    It 'contains at least one day and one assignment in each schedule file' -ForEach $jsonPaths {
        $data = Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json

        @($data.days).Count | Should -BeGreaterThan 0
        @($data.assignments).Count | Should -BeGreaterThan 0
    }

    It 'stores the expected day fields in the seed dataset' {
        $seedPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'data\SunNFun-2026-Schedule.seed.json'
        $seedData = Get-Content -LiteralPath $seedPath -Raw | ConvertFrom-Json
        $firstDay = @($seedData.days)[0]

        $firstDay.PSObject.Properties.Name | Should -Contain 'date'
        $firstDay.PSObject.Properties.Name | Should -Contain 'label'
        $firstDay.PSObject.Properties.Name | Should -Contain 'displayName'
        $firstDay.PSObject.Properties.Name | Should -Contain 'phase'
        $firstDay.PSObject.Properties.Name | Should -Contain 'hours'
    }

    It 'stores the expected assignment fields in the seed dataset' {
        $seedPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'data\SunNFun-2026-Schedule.seed.json'
        $seedData = Get-Content -LiteralPath $seedPath -Raw | ConvertFrom-Json
        $firstAssignment = @($seedData.assignments)[0]

        $firstAssignment.PSObject.Properties.Name | Should -Contain 'assignmentId'
        $firstAssignment.PSObject.Properties.Name | Should -Contain 'date'
        $firstAssignment.PSObject.Properties.Name | Should -Contain 'volunteerName'
        $firstAssignment.PSObject.Properties.Name | Should -Contain 'email'
        $firstAssignment.PSObject.Properties.Name | Should -Contain 'status'
    }
}