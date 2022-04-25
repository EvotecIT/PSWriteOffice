# This is just a show what can be quickly done using .NET before I get to do it's PowerShell version

Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocumentWithSomeText.docx
$Document.BuiltinDocumentProperties.Title = "This is title"
$Document.BuiltinDocumentProperties.Subject = "This is subject aka subtitle"

New-OfficeWordText -Document $Document -Text 'This is a test 1 ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test 2', ' and we continue using different color' -Color Blue, Gold -Alignment Right
New-OfficeWordText -Document $Document -Text 'This is a test 3 ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test 4', ' something else' -Color Green, Gold -Alignment Right
New-OfficeWordText -Document $Document -Text 'This is a test 5 ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test 6', ' even more text' -Color Blue, Gold -Alignment Right

foreach ($Paragraph in $Document.Paragraphs) {
    if ($Paragraph.Text) {
        if ($Paragraph.Text -like "*test 4*") {
            $Paragraph.Text = "We replace that value, but we keep the same color"
        }
    }
}

Save-OfficeWord -Document $Document -Show