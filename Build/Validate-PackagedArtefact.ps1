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

if (-not $SkipBuild -or -not (Test-Path -LiteralPath $ArtefactManifest)) {
    Import-Module PSPublishModule -Force -ErrorAction Stop
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

$RequiredCommands = @(
    'New-OfficeWord',
    'New-OfficeExcel',
    'New-OfficePowerPoint',
    'ConvertTo-OfficeCsv',
    'Get-OfficeMarkdown'
)

$RequiredCommandsLiteral = "'" + ($RequiredCommands -join "', '") + "'"

$ImportChecks = @(
    @{
        Name    = 'PowerShell 7'
        Command = (@'
Import-Module '__ARTEFACT__' -Force -ErrorAction Stop
$required = @(__REQUIRED_COMMANDS__)
$missing = foreach ($name in $required) {
    if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
        $name
    }
}
if ($missing) {
    throw ('Missing required commands: ' + ($missing -join ', '))
}
(Get-Command -Module PSWriteOffice | Measure-Object).Count
'@).Replace('__ARTEFACT__', $ArtefactManifestLiteral).Replace('__REQUIRED_COMMANDS__', $RequiredCommandsLiteral)
        Engine  = 'pwsh'
    }
    @{
        Name    = 'Windows PowerShell'
        Command = (@'
Import-Module '__ARTEFACT__' -Force -ErrorAction Stop
$required = @(__REQUIRED_COMMANDS__)
$missing = foreach ($name in $required) {
    if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
        $name
    }
}
if ($missing) {
    throw ('Missing required commands: ' + ($missing -join ', '))
}
(Get-Command -Module PSWriteOffice | Measure-Object).Count
'@).Replace('__ARTEFACT__', $ArtefactManifestLiteral).Replace('__REQUIRED_COMMANDS__', $RequiredCommandsLiteral)
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
