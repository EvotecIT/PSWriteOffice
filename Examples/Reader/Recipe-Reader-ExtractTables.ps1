param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Reader')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Reader-Table-Source.md'
$sidecars = Join-Path $OutputDirectory 'Reader-Table-Sidecars'
Set-Content -Path $path -Value @(
    '# Service scores'
    ''
    '| Service | Score |'
    '| --- | ---: |'
    '| Identity | 98 |'
    '| Messaging | 94 |'
) -Encoding UTF8

$tables = @(Get-OfficeDocumentTable -Path $path)
$exports = @(Get-OfficeDocumentTable -Path $path -OutputDirectory $sidecars)
$chunks = @(Get-OfficeDocumentChunk -Path $path)

[pscustomobject]@{
    Path         = $path
    Tables       = $tables.Count
    Rows         = $tables[0].Rows.Count
    Sidecars     = @($exports | Where-Object Written).Count
    Chunks       = $chunks.Count
} | Format-List
