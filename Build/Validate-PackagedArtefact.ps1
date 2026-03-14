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
        $env:PSPUBLISHMODULE_MANIFEST,
        'C:\Support\GitHub\PSPublishModule\Module\Artefacts\Unpacked\v3.0.1\Modules\PSPublishModule\PSPublishModule.psd1',
        'C:\Support\GitHub\PSPublishModule\Module\Artefacts\Unpacked\Modules\PSPublishModule\PSPublishModule.psd1'
    ) | Where-Object { $_ }

    foreach ($candidate in $candidatePaths) {
        if (Test-Path -LiteralPath $candidate) {
            Import-Module -Name $candidate -Force
            return
        }
    }

    $availableModule = Get-Module -ListAvailable -Name PSPublishModule |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $availableModule) {
        throw 'PSPublishModule was not found. Install it from PSGallery or set $env:PSPUBLISHMODULE_MANIFEST to the module manifest path.'
    }

    Import-Module -Name $availableModule.Path -Force
}

if (-not $SkipBuild -or -not (Test-Path -LiteralPath $ArtefactManifest)) {
    Import-PSPublishModule
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
