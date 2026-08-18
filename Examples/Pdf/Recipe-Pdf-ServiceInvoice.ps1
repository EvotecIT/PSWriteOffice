param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Pdf')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Service-Invoice.pdf'
$lines = @(
    [pscustomobject]@{ Description = 'Automation assessment'; Quantity = 1; UnitPrice = '$1,200.00'; Amount = '$1,200.00' }
    [pscustomobject]@{ Description = 'Report implementation'; Quantity = 3; UnitPrice = '$850.00'; Amount = '$2,550.00' }
    [pscustomobject]@{ Description = 'Handover workshop'; Quantity = 1; UnitPrice = '$600.00'; Amount = '$600.00' }
)

New-OfficePdf -Path $path {
    PdfTheme Report
    PdfMetadata -Title 'Service invoice INV-2026-0818' -Author 'Northwind Automation'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfHeader 'NORTHWIND AUTOMATION'
    PdfFooter 'Invoice INV-2026-0818 | Page {page}/{pages}'

    PdfHeading 'SERVICE INVOICE' -Level 1 -Color '#0F766E'
    PdfText -Run @(
        @{ Text = 'Invoice: '; Bold = $true }
        @{ Text = 'INV-2026-0818' }
        @{ Text = '    Issue date: '; Bold = $true }
        @{ Text = '2026-08-18' }
    )
    PdfText -Run @(
        @{ Text = 'Bill to: '; Bold = $true }
        @{ Text = 'Contoso Ltd., 1 Example Street, London' }
    )
    PdfHr -Color '#0F766E' -SpacingBefore 12 -SpacingAfter 14

    PdfTable -InputObject $lines -Property Description,Quantity,UnitPrice,Amount -Header 'Description','Qty','Unit price','Amount' -HeaderFill '#0F766E' -HeaderTextColor '#FFFFFF' -RightAlignNumeric -AutoFitColumns

    PdfText -Text 'Subtotal: $4,350.00' -Align Right -Bold
    PdfText -Text 'Tax (20%): $870.00' -Align Right
    PdfText -Text 'Total due: $5,220.00' -Align Right -Bold -FontSize 14 -Color '#0F766E'

    PdfPanel 'Payment terms: 14 days. Include the invoice number with the transfer.'
    PdfHeading 'Questions' -Level 2
    PdfText -Run @(
        @{ Text = 'Contact ' }
        @{ Text = 'billing@example.com'; LinkUri = 'mailto:billing@example.com'; Color = '#2563EB' }
        @{ Text = ' before the due date if any line needs correction.' }
    )
} | Out-Null

Write-Host "PDF service invoice saved to $path"
