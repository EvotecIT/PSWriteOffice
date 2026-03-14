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
$PSPublishModuleManifest = 'C:\Support\GitHub\PSPublishModule\Module\Artefacts\Unpacked\v3.0.1\Modules\PSPublishModule\PSPublishModule.psd1'

if (-not $SkipBuild -or -not (Test-Path -LiteralPath $ArtefactManifest)) {
    Import-Module -Name $PSPublishModuleManifest -Force
    & "$RepoRoot\Build\Manage-PSWriteOffice.ps1"
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
