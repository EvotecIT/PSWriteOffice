param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Pdf-Inspection-Source.pdf'
PdfNew -Path $path -CreateOutlineFromHeadings {
    PdfMetadata -Title 'Inspection source' -Author 'PSWriteOffice'
    PdfHeading 'Summary'; PdfParagraph 'The service is ready for review.'
    PdfPageBreak
    PdfHeading 'Evidence'; PdfParagraph 'Evidence is retained for 90 days.'
}

$info = Get-OfficePdfInfo -Path $path
$preflight = Get-OfficePdfPreflight -Path $path
$pages = @(Get-OfficePdfText -Path $path -ByPage)

[pscustomobject]@{
    Path       = $path
    Pages      = $info.PageCount
    Title      = $info.Metadata.Title
    CanRead    = $preflight.CanRead
    TextPages  = $pages.Count
    Characters = (($pages.Text -join '').Length)
} | Format-List
