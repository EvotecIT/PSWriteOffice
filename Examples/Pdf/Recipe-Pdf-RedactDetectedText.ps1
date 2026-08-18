param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$source = Join-Path $OutputDirectory 'Pdf-Redaction-Source.pdf'
$redacted = Join-Path $OutputDirectory 'Pdf-Redacted.pdf'
PdfNew -Path $source {
    PdfHeading 'Incident record'
    PdfParagraph 'Visible owner: Operations'
    PdfParagraph 'Secret account: 123-45-6789'
    PdfParagraph 'Visible status: Closed'
}

$block = Get-OfficePdfText -Path $source -AsTextBlock |
    Where-Object { $_.Text -match 'Secret account' } |
    Select-Object -First 1
if (-not $block) { throw 'The expected text block was not found.' }

$x = [math]::Min($block.XStart, $block.XEnd) - 2
$width = [math]::Abs($block.XEnd - $block.XStart) + 4
$y = $block.BaselineY - 14
ConvertTo-OfficePdfRedacted -Path $source -OutputPath $redacted -PageNumber $block.PageNumber `
    -X $x -Y $y -Width $width -Height 20 -FillColor '#111111'

$text = Get-OfficePdfText -Path $redacted
[pscustomobject]@{
    Path          = $redacted
    SecretRemoved = $text -notmatch '123-45-6789'
    VisibleKept   = $text -match 'Visible status'
} | Format-List
