Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\Test1.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bol' -Bold $null, $true -Underline $null, $true -Space Preserve, Preserve

$DataTable = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'Nope' }
)

New-OfficeWordTable -DataTable $DataTable -TableLayout Autofit

Save-OfficeWord -Document $Document -Show