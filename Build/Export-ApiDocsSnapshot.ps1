param(
    [string] $ModuleRoot = "$PSScriptRoot\..\Artefacts\Unpacked\Modules\PSWriteOffice",
    [string] $OutputRoot = "$PSScriptRoot\..\Docs\Generated",
    [string] $ArtifactsRoot = "$PSScriptRoot\..\WebsiteArtifacts",
    [switch] $SkipBuild
)

$ErrorActionPreference = 'Stop'

$moduleName = 'PSWriteOffice'
$slug = 'pswriteoffice'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$validateScript = Join-Path $PSScriptRoot 'Validate-PackagedArtefact.ps1'

if (-not (Test-Path -LiteralPath $validateScript -PathType Leaf)) {
    throw "Validation script not found: $validateScript"
}

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [string] $BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Write-CommandMetadata {
    param(
        [Parameter(Mandatory)]
        [string] $ModuleManifestPath,
        [Parameter(Mandatory)]
        [string] $OutputPath
    )

    Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
    try {
        Import-Module $ModuleManifestPath -Force -ErrorAction Stop | Out-Null
        $allCommands = @(Get-Command -Module $moduleName -ErrorAction Stop)
        $aliasesByTarget = @{}

        foreach ($aliasCommand in $allCommands | Where-Object CommandType -EQ 'Alias') {
            $targetName = $aliasCommand.ResolvedCommandName
            if ([string]::IsNullOrWhiteSpace($targetName)) {
                continue
            }

            if (-not $aliasesByTarget.ContainsKey($targetName)) {
                $aliasesByTarget[$targetName] = [System.Collections.Generic.List[string]]::new()
            }

            if (-not $aliasesByTarget[$targetName].Contains($aliasCommand.Name)) {
                $null = $aliasesByTarget[$targetName].Add($aliasCommand.Name)
            }
        }

        $commandMetadata = foreach ($command in $allCommands | Where-Object CommandType -In @('Function', 'Cmdlet', 'Filter', 'ExternalScript') | Sort-Object Name) {
            [ordered]@{
                name    = $command.Name
                kind    = if ($command.CommandType -EQ 'Cmdlet') { 'Cmdlet' } else { 'Function' }
                aliases = @($aliasesByTarget[$command.Name] | Sort-Object -Unique)
            }
        }

        $payload = [ordered]@{
            moduleName  = $moduleName
            generatedAt = (Get-Date).ToString('o')
            commands    = @($commandMetadata)
        }

        $payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    } finally {
        Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
    }
}

$resolvedModuleRoot = Resolve-Path -LiteralPath $ModuleRoot -ErrorAction SilentlyContinue
if ($resolvedModuleRoot) {
    $ModuleRoot = $resolvedModuleRoot.Path
}

if (-not $SkipBuild) {
    & $validateScript -ArtefactModulePath $ModuleRoot
} elseif (-not (Test-Path -LiteralPath (Join-Path $ModuleRoot "$moduleName.psd1") -PathType Leaf)) {
    & $validateScript -ArtefactModulePath $ModuleRoot
}

$helpPath = Join-Path $ModuleRoot "en-US\$moduleName-help.xml"
if (-not (Test-Path -LiteralPath $helpPath -PathType Leaf)) {
    throw "Generated help XML not found: $helpPath"
}

$moduleManifestPath = Join-Path $ModuleRoot "$moduleName.psd1"
if (-not (Test-Path -LiteralPath $moduleManifestPath -PathType Leaf)) {
    throw "Packaged module manifest not found: $moduleManifestPath"
}

$examplesSource = Join-Path $repoRoot 'Examples'
$resolvedOutputRoot = Resolve-AbsolutePath -Path $OutputRoot -BasePath $repoRoot
$resolvedArtifactsRoot = Resolve-AbsolutePath -Path $ArtifactsRoot -BasePath $repoRoot

New-Item -ItemType Directory -Path $resolvedOutputRoot -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedArtifactsRoot -Force | Out-Null

$outputHelpPath = Join-Path $resolvedOutputRoot "$moduleName-help.xml"
Copy-Item -LiteralPath $helpPath -Destination $outputHelpPath -Force

$apiRoot = Join-Path $resolvedArtifactsRoot 'apidocs\powershell'
$examplesTarget = Join-Path $apiRoot 'examples'
if (Test-Path -LiteralPath $apiRoot) {
    Remove-Item -LiteralPath $apiRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $apiRoot -Force | Out-Null
Copy-Item -LiteralPath $helpPath -Destination (Join-Path $apiRoot "$moduleName-help.xml") -Force
Copy-Item -LiteralPath $moduleManifestPath -Destination (Join-Path $apiRoot "$moduleName.psd1") -Force
Write-CommandMetadata -ModuleManifestPath $moduleManifestPath -OutputPath (Join-Path $apiRoot 'command-metadata.json')

if (Test-Path -LiteralPath $examplesSource -PathType Container) {
    Copy-Item -LiteralPath $examplesSource -Destination $examplesTarget -Recurse -Force
}

$version = (Import-PowerShellDataFile -Path $moduleManifestPath).ModuleVersion.ToString()
$commit = (& git -C $repoRoot rev-parse HEAD).Trim()
$manifest = [ordered]@{
    slug        = $slug
    name        = $moduleName
    description = 'PSWriteOffice website artifacts for the Evotec multi-project hub.'
    mode        = 'hub-full'
    contentMode = 'hybrid'
    status      = 'active'
    listed      = $false
    version     = $version
    generatedAt = (Get-Date).ToString('o')
    commit      = $commit
    links       = [ordered]@{
        source = 'https://github.com/EvotecIT/PSWriteOffice'
    }
    surfaces    = [ordered]@{
        docs         = $false
        apiPowerShell = $true
        apiDotNet    = $false
        examples     = $false
    }
    artifacts   = [ordered]@{
        api      = 'WebsiteArtifacts/apidocs'
        examples = 'Examples'
    }
}

$manifestPath = Join-Path $resolvedArtifactsRoot 'project-manifest.json'
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

[PSCustomObject]@{
    OutputPath          = $outputHelpPath
    SourceHelpPath      = $helpPath
    SnapshotSizeBytes   = (Get-Item -LiteralPath $outputHelpPath).Length
    WebsiteArtifactsRoot = $resolvedArtifactsRoot
    ApiRoot             = $apiRoot
    ManifestPath        = $manifestPath
    ExamplesCopied      = Test-Path -LiteralPath $examplesTarget -PathType Container
}
