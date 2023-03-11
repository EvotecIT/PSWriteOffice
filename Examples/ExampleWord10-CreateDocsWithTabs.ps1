#Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test, \t\t very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text "This is a test, `t`t very big test ", 'and this should be bold' -Bold $null, $true -Underline Dash, $null

$p3 = New-OfficeWordText -Document $Document -Bold 1, 0 -Underline $null, Dash -Text "We are ", "`t`tJohn Doe and Jane Doe`t`t" -ReturnObject

Save-OfficeWord -Document $Document -Show