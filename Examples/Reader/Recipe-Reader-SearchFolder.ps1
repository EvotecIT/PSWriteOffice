param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Reader')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$folder = Join-Path $OutputDirectory 'Search-Corpus'
New-Item -Path $folder -ItemType Directory -Force | Out-Null
$wordPath = Join-Path $folder 'policy.docx'
$excelPath = Join-Path $folder 'register.xlsx'
$markdownPath = Join-Path $folder 'notes.md'
WordNew -Path $wordPath { WordSection { WordParagraph -Text 'Retention policy requires seven years.' } }
ExcelNew -Path $excelPath { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Retention owner'; ExcelCell -Address B1 -Value 'Compliance' } }
Set-Content -Path $markdownPath -Value '# Notes', 'Retention evidence is reviewed quarterly.' -Encoding UTF8

$matches = @(Search-OfficeDocument -Path $folder -Recurse -Query 'Retention' -MaxDocuments 20 -AllResults)
$matches | Select-Object DocumentType,Path,Match | Format-Table -AutoSize
