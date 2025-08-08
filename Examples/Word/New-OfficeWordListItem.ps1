Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\ListItemDocument.docx"
$doc = New-OfficeWord -FilePath $path
$list = New-OfficeWordList -Document $doc
New-OfficeWordListItem -List $list -Text 'First item'
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
