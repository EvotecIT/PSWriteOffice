param(
    [string] $ModuleRoot = (Split-Path -Parent $PSScriptRoot),
    [switch] $SkipBuild,
    [switch] $SkipPester,
    [switch] $SkipExamples,
    [switch] $FailOnMissingArtifacts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Step {
    param(
        [Parameter(Mandatory)]
        [string] $Message
    )

    Write-Host "==> $Message" -ForegroundColor Cyan
}

function New-CheckResult {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [ValidateSet('Passed', 'Skipped', 'Failed')]
        [string] $Status,

        [string] $Details = ''
    )

    [PSCustomObject]@{
        Name    = $Name
        Status  = $Status
        Details = $Details
    }
}

function Get-NativeCommandErrorMessage {
    param(
        [Parameter(Mandatory)]
        [string] $Command
    )

    $exitCode = if ($null -ne $global:LASTEXITCODE) { $global:LASTEXITCODE } else { -1 }
    return "'$Command' exited with code $exitCode."
}

function Invoke-ExampleCheck {
    param(
        [Parameter(Mandatory)]
        [string] $ScriptPath,

        [Parameter(Mandatory)]
        [string[]] $Outputs,

        [switch] $MissingIsFailure
    )

    if (-not (Test-Path -Path $ScriptPath)) {
        if ($MissingIsFailure) {
            return New-CheckResult -Name (Split-Path -Leaf $ScriptPath) -Status Failed -Details "Example not found: $ScriptPath"
        }

        return New-CheckResult -Name (Split-Path -Leaf $ScriptPath) -Status Skipped -Details "Example not found on this branch: $ScriptPath"
    }

    foreach ($output in $Outputs) {
        if (Test-Path -Path $output) {
            Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
        }
    }

    try {
        & $ScriptPath | Out-Null

        $missingOutputs = foreach ($output in $Outputs) {
            if (-not (Test-Path -Path $output)) {
                $output
            }
        }

        if ($missingOutputs) {
            return New-CheckResult -Name (Split-Path -Leaf $ScriptPath) -Status Failed -Details ("Missing outputs: " + ($missingOutputs -join ', '))
        }

        return New-CheckResult -Name (Split-Path -Leaf $ScriptPath) -Status Passed -Details 'Example executed and expected outputs were created.'
    } catch {
        return New-CheckResult -Name (Split-Path -Leaf $ScriptPath) -Status Failed -Details $_.Exception.Message
    } finally {
        foreach ($output in $Outputs) {
            if (Test-Path -Path $output) {
                Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

$resolvedRoot = (Resolve-Path -Path $ModuleRoot).Path
$manifestPath = Join-Path $resolvedRoot 'PSWriteOffice.psd1'
if (-not (Test-Path -Path $manifestPath)) {
    throw "Unable to find PSWriteOffice.psd1 in '$resolvedRoot'."
}

$releaseArtifactChecks = @(
    'Sources/PSWriteOffice/Cmdlets/Word/UpdateOfficeWordTextCommand.cs'
    'Sources/PSWriteOffice/Cmdlets/Word/AddOfficeWordTableCellCommand.cs'
    'Sources/PSWriteOffice/Cmdlets/Word/AddOfficeWordChartCommand.cs'
)
$missingReleaseArtifacts = @(
    foreach ($artifact in $releaseArtifactChecks) {
        $artifactPath = Join-Path $resolvedRoot $artifact
        if (-not (Test-Path -Path $artifactPath)) {
            $artifact
        }
    }
)

$results = [System.Collections.Generic.List[object]]::new()

Push-Location $resolvedRoot
try {
    if (-not $SkipBuild) {
        Write-Step 'Building PSWriteOffice solution'
        try {
            dotnet build 'Sources/PSWriteOffice.sln'
            if ($LASTEXITCODE -ne 0) {
                $results.Add((New-CheckResult -Name 'dotnet build Sources/PSWriteOffice.sln' -Status Failed -Details (Get-NativeCommandErrorMessage -Command 'dotnet build Sources/PSWriteOffice.sln')))
            } else {
                $results.Add((New-CheckResult -Name 'dotnet build Sources/PSWriteOffice.sln' -Status Passed))
            }
        } catch {
            $results.Add((New-CheckResult -Name 'dotnet build Sources/PSWriteOffice.sln' -Status Failed -Details $_.Exception.Message))
        }
    } else {
        $results.Add((New-CheckResult -Name 'dotnet build Sources/PSWriteOffice.sln' -Status Skipped -Details 'Skipped by caller.'))
    }

    if (-not $SkipPester) {
        $missingArtifactDetails = "Missing required release artifacts: " + ($missingReleaseArtifacts -join ', ')
        if ($missingReleaseArtifacts.Count -gt 0 -and -not $FailOnMissingArtifacts) {
            $results.Add((New-CheckResult -Name 'Word release-candidate Pester tests' -Status Skipped -Details ("Run on the merged release branch. Missing artifacts: " + ($missingReleaseArtifacts -join ', '))))
        } elseif ($missingReleaseArtifacts.Count -gt 0) {
            $results.Add((New-CheckResult -Name 'Release artifacts' -Status Failed -Details $missingArtifactDetails))
        } else {
            Write-Step 'Running Word release-candidate Pester tests'
            $testFiles = @(
                'Tests/WordDsl.Tests.ps1',
                'Tests/WordReader.Tests.ps1'
            )

            foreach ($testFile in $testFiles) {
                if (-not (Test-Path -Path $testFile)) {
                    $status = if ($FailOnMissingArtifacts) { 'Failed' } else { 'Skipped' }
                    $details = if ($FailOnMissingArtifacts) { "Missing test file: $testFile" } else { "Missing test file on this branch: $testFile" }
                    $results.Add((New-CheckResult -Name $testFile -Status $status -Details $details))
                    continue
                }

                try {
                    $pesterResult = Invoke-Pester $testFile -Output Detailed -PassThru
                    if ($null -eq $pesterResult) {
                        $results.Add((New-CheckResult -Name $testFile -Status Failed -Details 'Invoke-Pester returned no result object.'))
                    } elseif ($pesterResult.FailedCount -gt 0) {
                        $results.Add((New-CheckResult -Name $testFile -Status Failed -Details ("Failed tests: " + $pesterResult.FailedCount)))
                    } else {
                        $results.Add((New-CheckResult -Name $testFile -Status Passed))
                    }
                } catch {
                    $results.Add((New-CheckResult -Name $testFile -Status Failed -Details $_.Exception.Message))
                }
            }
        }
    } else {
        $results.Add((New-CheckResult -Name 'Word Pester tests' -Status Skipped -Details 'Skipped by caller.'))
    }

    if (-not $SkipExamples) {
        Write-Step 'Running release examples'
        $exampleChecks = @(
            @{
                ScriptPath = Join-Path $resolvedRoot 'Examples/Word/Example-WordLineBreaks.ps1'
                Outputs    = @((Join-Path $resolvedRoot 'Examples/Documents/Example-WordLineBreaks.docx'))
            }
            @{
                ScriptPath = Join-Path $resolvedRoot 'Examples/Word/Example-WordCharts.ps1'
                Outputs    = @((Join-Path $resolvedRoot 'Examples/Documents/Word-Charts.docx'))
            }
            @{
                ScriptPath = Join-Path $resolvedRoot 'Examples/Word/Example-WordTableCalculatedColumns.ps1'
                Outputs    = @((Join-Path $resolvedRoot 'Examples/Documents/Word-CalculatedColumns.docx'))
            }
            @{
                ScriptPath = Join-Path $resolvedRoot 'Examples/Word/Example-WordTableCells.ps1'
                Outputs    = @(
                    (Join-Path $resolvedRoot 'Examples/Word/Example-WordTableCells.docx')
                    (Join-Path $resolvedRoot 'Examples/Word/Example-WordTableCells.png')
                )
            }
            @{
                ScriptPath = Join-Path $resolvedRoot 'Examples/Word/Example-WordReplaceText.ps1'
                Outputs    = @((Join-Path $resolvedRoot 'Examples/Word/Example-WordReplaceText.docx'))
            }
        )

        foreach ($exampleCheck in $exampleChecks) {
            $result = Invoke-ExampleCheck -ScriptPath $exampleCheck.ScriptPath -Outputs $exampleCheck.Outputs -MissingIsFailure:$FailOnMissingArtifacts
            $results.Add($result)
        }
    } else {
        $results.Add((New-CheckResult -Name 'Release examples' -Status Skipped -Details 'Skipped by caller.'))
    }
} finally {
    Pop-Location
}

Write-Host ''
Write-Host 'Release smoke test summary:' -ForegroundColor Yellow
$results | Format-Table -AutoSize | Out-Host

$failed = @($results | Where-Object Status -eq 'Failed')
if ($failed.Count -gt 0) {
    throw ("Release smoke checks failed: " + (($failed | ForEach-Object Name) -join ', '))
}
