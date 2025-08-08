Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\RemoveFooter.docx"
$doc = New-OfficeWord -FilePath $path
Remove-OfficeWordFooter -Document $doc
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
