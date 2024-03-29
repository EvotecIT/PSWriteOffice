﻿Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\DocList.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bol' -Bold $null, $true -Underline $null, 'Dashed'

$List = New-OfficeWordList -Document $Document {
    New-OfficeWordListItem -Text 'Test1'
    New-OfficeWordListItem -Text 'Test2'
    New-OfficeWordListItem -Text 'Test3' -Level 2
} -Style Heading1ai

$Document.AddParagraph()
$Document.AddParagraph()
$Document.AddParagraph()
$Document.AddParagraph()

$P1 = New-OfficeWordListItem -Text 'Test4' -List $List

New-OfficeWordText -Document $Document -Text "But lists don't really have to be next to each other" -Bold $true -Alignment Center -Color RoseBud

$P2 = New-OfficeWordListItem -Text 'Test5' -List $List

New-OfficeWordText -Document $Document -Text "Here's another way to define list" -Bold $true -Alignment Center -Color RoseBud

$List1 = $Document.AddList([OfficeIMO.Word.WordListStyle]::Headings111)
$Paragraph1 = $List1.AddItem("Test", 2)
$Paragraph1.FontSize = 24
$Paragraph2 = $List1.AddItem("Test1", 1)
$Paragraph2.Strike = $true
$Paragraph3 = $List1.AddItem("Test2", 0)
$Paragraph3 = $Paragraph3.SetColor([SixLabors.ImageSharp.Color]::Red)

$null = $Document.AddPageBreak()

$P3 = New-OfficeWordListItem -Text 'Test6' -List $List
$P3.Bold = $true
$P3.FontSize = 20

# Get all listitems and find the listitem you want to change
$List.ListItems | Format-Table
$List.ListItems[0].Bold = $true

for ($i = 1; $i -le 3; $i++) {
    New-OfficeWordText -Document $Document -Text "Software Category #$i" -Style Heading1
    New-OfficeWordList {
        New-OfficeWordListItem -Text 'Test1'
        New-OfficeWordListItem -Text 'Test2'
        New-OfficeWordListItem -Text 'Test3' -Level 2
    } -Style Headings111 -Document $Document
}

Save-OfficeWord -Document $Document -Show -FilePath $PSScriptRoot\Documents\Doc1Updated.docx