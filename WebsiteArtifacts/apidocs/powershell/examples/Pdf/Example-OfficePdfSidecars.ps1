$ErrorActionPreference = 'Stop'

$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1'
}

if (-not (Get-Module -Name PSWriteOffice)) {
    Import-Module $modulePath -ErrorAction Stop
}

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$wordPath = Join-Path $documents 'Example-OfficePdfSidecar.docx'
$wordPdf = Join-Path $documents 'Example-OfficePdfSidecar-Word.pdf'
$excelPath = Join-Path $documents 'Example-OfficePdfSidecar.xlsx'
$excelPdf = Join-Path $documents 'Example-OfficePdfSidecar-Excel.pdf'
$markdownPath = Join-Path $documents 'Example-OfficePdfSidecar.md'
$markdownPdf = Join-Path $documents 'Example-OfficePdfSidecar-Markdown.pdf'
$powerPointPath = Join-Path $documents 'Example-OfficePdfSidecar.pptx'
$powerPointPdf = Join-Path $documents 'Example-OfficePdfSidecar-PowerPoint.pdf'

$rows = @(
    [pscustomobject]@{ Name = 'Alpha'; Status = 'Ready'; Count = 12 }
    [pscustomobject]@{ Name = 'Beta'; Status = 'Review'; Count = 7 }
)

New-OfficeWord -Path $wordPath -PdfPath $wordPdf {
    WordParagraph -Text 'Word PDF sidecar' -Style Heading1
    WordParagraph 'The Word document and PDF sidecar are saved in one command.'
    WordTable -InputObject $rows -Layout AutoFitToWindow
} | Out-Null

New-OfficeExcel -Path $excelPath -PdfPath $excelPdf {
    ExcelSheet -Name 'Summary' {
        ExcelTable -Data $rows
        ExcelAutoFit
    }
} | Out-Null

New-OfficeMarkdown -Path $markdownPath -PdfPath $markdownPdf {
    MarkdownHeading -Level 1 -Text 'Markdown PDF sidecar'
    MarkdownParagraph 'Markdown keeps the same authoring mindset with format-appropriate simplification.'
    MarkdownTable -InputObject $rows
} | Out-Null

New-OfficePowerPoint -Path $powerPointPath -PdfPath $powerPointPdf {
    PptSlide {
        PptTitle -Title 'PowerPoint PDF sidecar'
        PptBullets -Bullets 'Create the deck', 'Save the PDF sidecar', 'Inspect generated output'
    }
} | Out-Null

Get-Item -LiteralPath $wordPdf, $excelPdf, $markdownPdf, $powerPointPdf |
    Select-Object FullName, Length |
    Format-Table -AutoSize
