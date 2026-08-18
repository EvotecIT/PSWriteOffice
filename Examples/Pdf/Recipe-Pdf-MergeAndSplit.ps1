param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$first = Join-Path $OutputDirectory 'Pdf-Pack-Part-A.pdf'
$second = Join-Path $OutputDirectory 'Pdf-Pack-Part-B.pdf'
$merged = Join-Path $OutputDirectory 'Pdf-Combined-Pack.pdf'
$splitDirectory = Join-Path $OutputDirectory 'Pdf-Combined-Pack-Pages'

PdfNew -Path $first { PdfHeading 'Part A'; PdfParagraph 'Operations summary.' }
PdfNew -Path $second { PdfHeading 'Part B'; PdfParagraph 'Detailed evidence.' }
Join-OfficePdf -Path $first,$second -OutputPath $merged -PageSize A4 -ResizeMode Fit -ResizeMargin 18
$parts = @(Split-OfficePdf -Path $merged -OutputDirectory $splitDirectory -Prefix 'page' -PagesPerDocument 1 -PadIndex -IndexWidth 2)

[pscustomobject]@{
    Path       = $merged
    Pages      = (Get-OfficePdfInfo -Path $merged).PageCount
    SplitFiles = $parts.Count
} | Format-List
