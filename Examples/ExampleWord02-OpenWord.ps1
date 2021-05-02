Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = Get-OfficeWord -FilePath $PSScriptRoot\Documents\Test.docx

New-OfficeWordText -Document $Document -Text 'Add more things!' -Bold $null, $true -Underline $null, $true -Space Preserve, Preserve
New-OfficeWordText -Document $Document -Text 'Add more things!' -Bold $null, $true -Underline $null, $true -Space Preserve, Preserve
New-OfficeWordText -Document $Document -Text 'Add more things!' -Bold $null, $true -Underline $null, $true -Space Preserve, Preserve

Save-OfficeWord -Document $Document -Show -FilePath $PSScriptRoot\Documents\Test6.docx