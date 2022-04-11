#Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test, very big test', 'ooops' -Color Blue, Gold -Alignment Right

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject

New-OfficeWordText -Document $Document -Text ' Attached to existing paragraph', ' continue' -Paragraph $Paragraph -Color Blue

Save-OfficeWord -Document $Document -Show