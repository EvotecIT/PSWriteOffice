Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\NewDocument.docx"
$document = New-OfficeWord -FilePath $path
Save-OfficeWord -Document $document
Close-OfficeWord -Document $document
