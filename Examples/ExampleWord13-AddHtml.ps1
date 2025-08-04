Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$docAsIs = New-OfficeWord -FilePath "$PSScriptRoot\Documents\HtmlAsIs.docx"
[PSWriteOffice.Services.Word.WordDocumentService]::AddHtml($docAsIs, '<p>Hello <b>World</b></p>', [PSWriteOffice.Services.Word.HtmlImportMode]::AsIs)
Save-OfficeWord -Document $docAsIs -Show

$docParsed = New-OfficeWord -FilePath "$PSScriptRoot\Documents\HtmlParsed.docx"
[PSWriteOffice.Services.Word.WordDocumentService]::AddHtml($docParsed, '<p>Hello <b>World</b></p>')
Save-OfficeWord -Document $docParsed -Show
