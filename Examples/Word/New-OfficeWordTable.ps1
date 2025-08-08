Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\TableDocument.docx"
$doc = New-OfficeWord -FilePath $path
$data = @(
    [PSCustomObject]@{ Name = 'A'; Value = 1 },
    [PSCustomObject]@{ Name = 'B'; Value = 2 }
)
New-OfficeWordTable -Document $doc -DataTable $data
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
