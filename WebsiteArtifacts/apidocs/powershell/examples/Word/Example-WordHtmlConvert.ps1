Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$docPath = Join-Path $documents 'Word-HtmlSource.docx'
$htmlPath = Join-Path $documents 'Word-HtmlSource.html'
$roundtripPath = Join-Path $documents 'Word-HtmlRoundtrip.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Text 'Hello from HTML conversion.' -Style Heading2
        Add-OfficeWordParagraph -Text 'This document will round-trip to HTML.'
    }
} | Out-Null

ConvertTo-OfficeWordHtml -Path $docPath -OutputPath $htmlPath -PassThru | Out-Null
ConvertFrom-OfficeWordHtml -Path $htmlPath -OutputPath $roundtripPath -PassThru | Out-Null

Write-Host "HTML saved to $htmlPath"
Write-Host "Round-trip document saved to $roundtripPath"
