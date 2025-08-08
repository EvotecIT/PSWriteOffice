Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\RemoveHeader.docx"
$doc = New-OfficeWord -FilePath $path
Remove-OfficeWordHeader -Document $doc
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
