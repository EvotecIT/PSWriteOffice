Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\Test5.docx
New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline $null, 'Double'

$DataTable = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'Nope' }
)

# not working yet
New-OfficeWordTable -DataTable $DataTable -TableLayout Autofit #-SkipHeader

Save-OfficeWord -Document $Document #-Show

Invoke-Item -LiteralPath "$PSScriptRoot\Documents\Test5.docx"