

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test, very big test', 'ooops' -Color Blue, Gold -Alignment Right

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject
$Paragraph = $Paragraph.AddBreak()
$Paragraph = New-OfficeWordText -Document $Document -Text ' Attached to existing paragraph', ' continue' -Paragraph $Paragraph -Color Blue -ReturnObject
$Paragraph.Bold = $true
$Paragraph = $Paragraph.SetItalic().AddText("More txt").SetBold()
$Paragraph = $Paragraph.AddText("Even more text").AddBreak()
$Paragraph.AddText("Mix and match").SetBold().SetItalic()
$Paragraph = $Paragraph.SetUnderline([DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::Dash)
Save-OfficeWord -Document $Document -Show