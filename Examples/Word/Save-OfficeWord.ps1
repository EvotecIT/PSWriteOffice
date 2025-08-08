Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\SaveDocument.docx"
$doc = New-OfficeWord -FilePath $path
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
