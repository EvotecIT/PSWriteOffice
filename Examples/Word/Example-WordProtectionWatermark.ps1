Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Word-ProtectedWatermark.docx'

New-OfficeWord -Path $path {
    Add-OfficeWordParagraph -Text 'Confidential report'
    Add-OfficeWordWatermark -Text 'CONFIDENTIAL'
    Protect-OfficeWordDocument -Password 'secret'
} | Out-Null

Write-Host "Document saved to $path"
