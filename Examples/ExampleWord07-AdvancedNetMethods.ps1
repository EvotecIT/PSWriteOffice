# This is just a show what can be quickly done using .NET before I get to do it's PowerShell version

Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx
$Document.BuiltinDocumentProperties.Title = "This is title"
$Document.BuiltinDocumentProperties.Subject = "This is subject aka subtitle"

$null = $Document.AddCoverPage([OfficeIMO.Word.CoverPageTemplate]::Austin)

$null = $Document.AddTableOfContent()

$null = $Document.AddPageBreak()

$ListTOC = $Document.AddTableOfContentList([OfficeIMO.Word.WordListStyle]::Headings111)

$null = $ListTOC.AddItem("Heading 1")

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

$null = $ListTOC.AddItem("Heading 2")

New-OfficeWordText -Document $Document -Text 'This is a test, very big test', 'ooops' -Color Blue, Gold -Alignment Right

$null = $ListTOC.AddItem("Heading 2.1", 1)

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject

$null = $ListTOC.AddItem("Heading 3")

New-OfficeWordText -Document $Document -Text ' Attached to existing paragraph', ' continue' -Paragraph $Paragraph -Color Blue

$null = $ListTOC.AddItem("Heading 3.1", 1)

$Document.TableOfContent.Update()

Save-OfficeWord -Document $Document -Show