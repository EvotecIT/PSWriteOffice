Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\ExistingDocument.docx"
$doc = New-OfficeWord -FilePath $path
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc

$loaded = Get-OfficeWord -FilePath $path
Close-OfficeWord -Document $loaded
