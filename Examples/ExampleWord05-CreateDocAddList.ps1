Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\DocList.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bol' -Bold $null, $true -Underline $null, 'Dashed'

# Not working yet
New-OfficeWordList -Document $Document {
    New-OfficeWordListItem -Text 'Test1'
    New-OfficeWordListItem -Text 'Test2'
    New-OfficeWordListItem -Text 'Test3'
}

Save-OfficeWord -Document $Document -Show -FilePath $PSScriptRoot\Documents\Doc1Updated.docx