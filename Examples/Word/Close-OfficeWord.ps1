Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\CloseDocument.docx"
$doc = New-OfficeWord -FilePath $path
Close-OfficeWord -Document $doc
