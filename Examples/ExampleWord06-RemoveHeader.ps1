Import-Module .\PSWriteOffice.psd1 -Force

$Document = Get-OfficeWord -FilePath "$PSScriptRoot\Documents\BasicDocument.docx"
Remove-OfficeWordFooter -Document $Document
Save-OfficeWord -Document $Document -Show -FilePath "$PSScriptRoot\Documents\BasicDocumentWithoutHeader.docx" -Retry 1