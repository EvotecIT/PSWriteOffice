Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\TextDocument.docx"
$doc = New-OfficeWord -FilePath $path
New-OfficeWordText -Document $doc -Text 'Hello from PSWriteOffice'
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
