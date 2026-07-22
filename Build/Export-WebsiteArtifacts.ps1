param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $ArtifactsRoot = '',
    [switch] $SkipBuild
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ArtifactsRoot)) {
    $ArtifactsRoot = Join-Path $RepositoryRoot 'WebsiteArtifacts'
}

$moduleName = 'PSWriteOffice'
$manifestPath = Join-Path $RepositoryRoot "$moduleName.psd1"
$manifest = Import-PowerShellDataFile -LiteralPath $manifestPath
$expectedCommands = @($manifest.CmdletsToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
    Sort-Object -Unique

function Get-HelpCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    [xml] $help = Get-Content -LiteralPath $Path -Raw
    @($help.helpItems.command) |
        ForEach-Object { [string] $_.details.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Find-CompleteBuiltModule {
    $unpackedRoot = Join-Path $RepositoryRoot 'Artefacts\Unpacked'
    if (-not (Test-Path -LiteralPath $unpackedRoot -PathType Container)) {
        return $null
    }

    foreach ($helpFile in Get-ChildItem -LiteralPath $unpackedRoot -Recurse -Filter "$moduleName-help.xml" -File |
        Sort-Object LastWriteTimeUtc -Descending) {
        $moduleRoot = Split-Path -Parent (Split-Path -Parent $helpFile.FullName)
        $builtManifestPath = Join-Path $moduleRoot "$moduleName.psd1"
        if (-not (Test-Path -LiteralPath $builtManifestPath -PathType Leaf)) {
            continue
        }

        $helpCommands = @(Get-HelpCommandNames -Path $helpFile.FullName)
        if ($helpCommands.Count -ne $expectedCommands.Count) {
            continue
        }

        if (@(Compare-Object -ReferenceObject $expectedCommands -DifferenceObject $helpCommands).Count -eq 0) {
            return [PSCustomObject]@{
                HelpPath = $helpFile.FullName
                ManifestPath = $builtManifestPath
                ModuleRoot = $moduleRoot
            }
        }
    }

    return $null
}

$builtModule = Find-CompleteBuiltModule
if (-not $builtModule -and -not $SkipBuild) {
    & (Join-Path $RepositoryRoot 'Build\Build-Module.ps1') -RunMode Build
    if ($LASTEXITCODE -ne 0) {
        throw "PSWriteOffice module build failed with exit code $LASTEXITCODE."
    }
    $builtModule = Find-CompleteBuiltModule
}

if (-not $builtModule) {
    throw "No complete built PSWriteOffice API bundle was found. Run Build\Build-Module.ps1 -RunMode Build first."
}

$sourceFilesByType = @{}
foreach ($sourceFile in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'Sources\PSWriteOffice\Cmdlets') -Recurse -Filter '*.cs' -File) {
    $sourceFilesByType[$sourceFile.BaseName] = $sourceFile.FullName
    $sourceText = Get-Content -LiteralPath $sourceFile.FullName -Raw
    foreach ($typeMatch in [regex]::Matches($sourceText, '\b(?:sealed\s+)?class\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\b')) {
        $sourceFilesByType[$typeMatch.Groups['name'].Value] = $sourceFile.FullName
    }
}

Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
try {
    Import-Module $builtModule.ManifestPath -Force -ErrorAction Stop | Out-Null
    $allCommands = @(Get-Command -Module $moduleName -ErrorAction Stop)
    $commandsByName = @{}
    $aliasesByTarget = @{}

    foreach ($command in $allCommands) {
        $commandsByName[$command.Name] = $command
        if ($command.CommandType -ne 'Alias' -or [string]::IsNullOrWhiteSpace($command.ResolvedCommandName)) {
            continue
        }

        if (-not $aliasesByTarget.ContainsKey($command.ResolvedCommandName)) {
            $aliasesByTarget[$command.ResolvedCommandName] = [System.Collections.Generic.List[string]]::new()
        }
        [void] $aliasesByTarget[$command.ResolvedCommandName].Add($command.Name)
    }

    $missingCommands = @($expectedCommands | Where-Object { -not $commandsByName.ContainsKey($_) })
    if ($missingCommands.Count -gt 0) {
        throw "The built module does not export every manifest cmdlet: $($missingCommands -join ', ')"
    }

    $commandMetadata = foreach ($commandName in $expectedCommands) {
        $command = $commandsByName[$commandName]
        $sourcePath = $null
        $sourceLine = $null
        $implementingType = $command.ImplementingType
        if ($implementingType -and $sourceFilesByType.ContainsKey($implementingType.Name)) {
            $sourceFullPath = $sourceFilesByType[$implementingType.Name]
            $sourcePath = [System.IO.Path]::GetRelativePath($RepositoryRoot, $sourceFullPath).Replace('\', '/')
            $typePattern = '\bclass\s+' + [regex]::Escape($implementingType.Name) + '\b'
            $typeMatch = Select-String -LiteralPath $sourceFullPath -Pattern $typePattern | Select-Object -First 1
            if ($typeMatch) {
                $sourceLine = $typeMatch.LineNumber
            }
        }

        $entry = [ordered]@{
            name = $commandName
            kind = 'Cmdlet'
            aliases = @($aliasesByTarget[$commandName] | Sort-Object -Unique)
        }
        if ($sourcePath) {
            $entry.sourcePath = $sourcePath
            $entry.sourceLine = $sourceLine
            $entry.sourceUrl = "https://github.com/EvotecIT/PSWriteOffice/blob/main/$sourcePath#L$sourceLine"
        }
        $entry
    }
} finally {
    Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
}

$apiRoot = Join-Path $ArtifactsRoot 'apidocs\powershell'
New-Item -ItemType Directory -Path $apiRoot -Force | Out-Null

Copy-Item -LiteralPath $builtModule.HelpPath -Destination (Join-Path $apiRoot "$moduleName-help.xml") -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $apiRoot "$moduleName.psd1") -Force

$metadata = [ordered]@{
    schemaVersion = 1
    moduleName = $moduleName
    commands = @($commandMetadata)
}
$metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $apiRoot 'command-metadata.json') -Encoding utf8

[PSCustomObject]@{
    OutputPath = (Resolve-Path -LiteralPath $apiRoot).Path
    CommandCount = $commandMetadata.Count
    SourceMappedCount = @($commandMetadata | Where-Object sourcePath).Count
    HelpPath = $builtModule.HelpPath
}
