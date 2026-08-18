param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$source = Join-Path $OutputDirectory 'Pdf-Canvas-Source.pdf'
$positioned = Join-Path $OutputDirectory 'Pdf-Positioned-Canvas.pdf'
PdfNew -Path $source {
    PdfHeading 'Fixed-position review copy'
    PdfParagraph 'Normal PdfText and PdfParagraph content remains in document flow.'
    PdfPageBreak
    PdfParagraph 'The canvas callback can place content differently on every page.'
}

$runs = @(
    New-OfficeTextRun -Text 'Owner: ' -Bold | ConvertTo-OfficePdfTextRun
    New-OfficeTextRun -Text 'Platform' -Color '#0F766E' | ConvertTo-OfficePdfTextRun
    New-OfficeTextRun -Text '  |  REVIEW COPY' -Italic | ConvertTo-OfficePdfTextRun
)
# Build a runtime-typed array after the module assembly is loaded. This also works when the script is parsed before Import-Module runs.
$nativeRuns = [Array]::CreateInstance($runs[0].GetType(), $runs.Count)
for ($index = 0; $index -lt $runs.Count; $index++) {
    $nativeRuns.SetValue($runs[$index], $index)
}

Add-OfficePdfCanvas -Path $source -OutputPath $positioned -Content {
    param($canvas, $page)
    # Canvas coordinates use PDF points from the visual top-left of the page.
    $null = $canvas.Text($nativeRuns, 36, 24, $page.Width - 72, 24, $null, 'Left', 10, 12)
    $null = $canvas.Text("Page $($page.PageNumber) of $($page.PageCount)", $page.Width - 136, $page.Height - 36, 100, 18, 9)
}

[pscustomobject]@{
    Path      = $positioned
    Pages     = (Get-OfficePdfInfo -Path $positioned).PageCount
    HasHeader = (Get-OfficePdfText -Path $positioned) -match 'REVIEW COPY'
} | Format-List
