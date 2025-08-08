Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\ListDocument.docx"
$doc = New-OfficeWord -FilePath $path
$list = New-OfficeWordList -Document $doc {
    New-OfficeWordListItem -Text 'Item1'
    New-OfficeWordListItem -Text 'Item2' -Level 1
}
Save-OfficeWord -Document $doc
Close-OfficeWord -Document $doc
