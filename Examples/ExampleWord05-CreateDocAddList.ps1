Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\DocList.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bol' -Bold $null, $true -Underline $null, 'Dashed'

# Not working yet
$List = New-OfficeWordList -Document $Document {
    New-OfficeWordListItem -Text 'Test1'
    New-OfficeWordListItem -Text 'Test2'
    New-OfficeWordListItem -Text 'Test3'
} -Style Heading1ai

$P1 = New-OfficeWordListItem -Text 'Test4' -List $List

New-OfficeWordText -Document $Document -Text "But lists don't really have to be next to each other" -Bold $true -Alignment Center -Color RoseBud

$P2 = New-OfficeWordListItem -Text 'Test5' -List $List

$null = $Document.AddPageBreak()

$P3 = New-OfficeWordListItem -Text 'Test6' -List $List
$P3.Bold = $true
$P3.FontSize = 20

Save-OfficeWord -Document $Document -Show -FilePath $PSScriptRoot\Documents\Doc1Updated.docx