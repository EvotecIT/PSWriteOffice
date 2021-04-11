Clear-Host
Import-Module .\PSOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\Test.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bol' -Bold $null, $true -UnderLine $null, $true -Space Preserve, Preserve

New-OfficeWordText -Document $Document -Text 'This is a test, very big test', 'ooops' -Color Blue, Gold -Alignment Right
New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center

Save-OfficeWord -Document $Document -Show