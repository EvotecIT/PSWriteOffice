param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Reader')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$folder = Join-Path $OutputDirectory 'Ingest-Corpus'
New-Item -Path $folder -ItemType Directory -Force | Out-Null
Set-Content -Path (Join-Path $folder 'service.html') -Value '<html><body><h1>Service</h1><p>Indexed HTML evidence.</p></body></html>' -Encoding UTF8
Set-Content -Path (Join-Path $folder 'control.json') -Value '{"control":"Indexed JSON evidence"}' -Encoding UTF8
Set-Content -Path (Join-Path $folder 'owner.yaml') -Value 'owner: Indexed YAML evidence' -Encoding UTF8
Set-Content -Path (Join-Path $folder 'ignored.txt') -Value 'This file is outside the selected extension set.' -Encoding UTF8

$result = Get-OfficeDocumentIngest -FolderPath $folder -Extension html,json,yaml -NoRecurse
[pscustomobject]@{
    Folder         = $folder
    FilesScanned   = $result.FilesScanned
    FilesParsed    = $result.FilesParsed
    ChunksProduced = $result.ChunksProduced
} | Format-List
