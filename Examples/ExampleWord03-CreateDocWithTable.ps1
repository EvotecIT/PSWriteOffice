Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\Test5.docx
New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline $null, 'Double'

$DataTable = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'Nope' }
)

New-OfficeWordTable -Document $Document -DataTable $DataTable -TableLayout Autofit -Style GridTable1LightAccent1

New-OfficeWordText -Document $Document
New-OfficeWordText -Document $Document -Text "Another table below" -Bold $true -Color RedBerry -Alignment Center

New-OfficeWordTable -Document $Document -DataTable $DataTable -TableLayout Autofit -Style GridTable2 -SkipHeader

Save-OfficeWord -Document $Document #-Show

Invoke-Item -LiteralPath "$PSScriptRoot\Documents\Test5.docx"