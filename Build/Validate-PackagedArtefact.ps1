param(
    [string] $ArtefactModulePath = "$PSScriptRoot\..\Artefacts\Unpacked\Modules\PSWriteOffice",
    [switch] $SkipBuild
)

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path -Path "$PSScriptRoot\..").Path
$ResolvedArtefactModulePath = Resolve-Path -Path $ArtefactModulePath -ErrorAction SilentlyContinue
if ($ResolvedArtefactModulePath) {
    $ArtefactModulePath = $ResolvedArtefactModulePath.Path
}
$ArtefactManifest = Join-Path $ArtefactModulePath 'PSWriteOffice.psd1'

function Import-PSPublishModule {
    $candidatePaths = @(
        $env:PSPUBLISHMODULE_MANIFEST
    ) | Where-Object { $_ }

    $candidateModules = @(
        foreach ($candidate in $candidatePaths) {
            if (Test-Path -LiteralPath $candidate) {
                Get-Item -LiteralPath $candidate
            }
        }
        Get-Module -ListAvailable -Name PSPublishModule | Sort-Object Version -Descending
    )

    foreach ($candidate in $candidateModules) {
        Remove-Module -Name PSPublishModule -Force -ErrorAction SilentlyContinue
        $candidateModulePath = if ($candidate.FullName) { $candidate.FullName } elseif ($candidate.Path) { $candidate.Path } else { [string] $candidate }
        Import-Module -Name $candidateModulePath -Force -ErrorAction Stop
        $newConfigurationBuild = Get-Command -Name New-ConfigurationBuild -ErrorAction SilentlyContinue
        if ($newConfigurationBuild -and $newConfigurationBuild.Parameters.ContainsKey('NETAssemblyLoadContext')) {
            return
        }
    }

    throw 'PSPublishModule with New-ConfigurationBuild -NETAssemblyLoadContext support was not found. Install the current PSPublishModule or set $env:PSPUBLISHMODULE_MANIFEST.'
}

if (-not $SkipBuild -or -not (Test-Path -LiteralPath $ArtefactManifest)) {
    Import-PSPublishModule
    $previousOfficeIMORoot = $env:OfficeIMORoot
    try {
        $env:OfficeIMORoot = Join-Path $RepoRoot '.missing-officeimo'
        & "$RepoRoot\Build\Manage-PSWriteOffice.ps1"
    } finally {
        if ($null -eq $previousOfficeIMORoot) {
            Remove-Item env:OfficeIMORoot -ErrorAction SilentlyContinue
        } else {
            $env:OfficeIMORoot = $previousOfficeIMORoot
        }
    }
}

if (-not (Test-Path -LiteralPath $ArtefactManifest)) {
    throw "Artefact manifest not found: $ArtefactManifest"
}

$ArtefactManifestLiteral = $ArtefactManifest.Replace("'", "''")

$ImportChecks = @(
    @{
        Name    = 'PowerShell 7'
        Command = (@'
Import-Module '__ARTEFACT__' -Force -ErrorAction Stop
$required = @('New-OfficeWord', 'New-OfficeExcel', 'New-OfficePowerPoint', 'ConvertTo-OfficeCsv', 'Get-OfficeMarkdown')
$missing = foreach ($name in $required) {
    if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
        $name
    }
}
if ($missing) {
    throw ('Missing required commands: ' + ($missing -join ', '))
}
(Get-Command -Module PSWriteOffice | Measure-Object).Count
'@).Replace('__ARTEFACT__', $ArtefactManifestLiteral)
        Engine  = 'pwsh'
    }
    @{
        Name    = 'Windows PowerShell'
        Command = (@'
Import-Module '__ARTEFACT__' -Force -ErrorAction Stop
$required = @('New-OfficeWord', 'New-OfficeExcel', 'New-OfficePowerPoint', 'ConvertTo-OfficeCsv', 'Get-OfficeMarkdown')
$missing = foreach ($name in $required) {
    if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
        $name
    }
}
if ($missing) {
    throw ('Missing required commands: ' + ($missing -join ', '))
}
(Get-Command -Module PSWriteOffice | Measure-Object).Count
'@).Replace('__ARTEFACT__', $ArtefactManifestLiteral)
        Engine  = 'powershell'
    }
)

foreach ($Check in $ImportChecks) {
    $Count = & $Check.Engine -NoLogo -NoProfile -Command $Check.Command
    if (-not $Count) {
        throw "Artefact import returned no command count for $($Check.Name)."
    }
    Write-Host "$($Check.Name): imported packaged module with $Count command(s)."
}
