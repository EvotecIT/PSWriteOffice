param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$source = Join-Path $OutputDirectory 'Pdf-Delivery-Source.pdf'
$sanitized = Join-Path $OutputDirectory 'Pdf-Delivery-Sanitized.pdf'
$optimized = Join-Path $OutputDirectory 'Pdf-Delivery-Optimized.pdf'
PdfNew -Path $source {
    PdfMetadata -Title 'Delivery copy' -Author 'PSWriteOffice'
    PdfHeading 'Delivery copy'
    PdfParagraph ('Repeated delivery evidence. ' * 60)
}

ConvertTo-OfficePdfSanitized -Path $source -OutputPath $sanitized
$report = ConvertTo-OfficePdfOptimized -Path $sanitized -OutputPath $optimized -AllowLarger -PassThruReport

[pscustomobject]@{
    Path          = $optimized
    SourceBytes   = (Get-Item $source).Length
    SanitizedBytes = (Get-Item $sanitized).Length
    OutputBytes   = (Get-Item $optimized).Length
    OptimizationActions = $report.ActionCount
    CanRead       = (Get-OfficePdfPreflight -Path $optimized).CanRead
} | Format-List
