Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\Doc1.docx
New-OfficeWordText -Document $Document -Text 'This is document' -Bold $true -Underline 
New-OfficeWordText -Document $Document -Text 'This is document' -Bold $true -Underline ([DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::Dash)
Save-OfficeWord -Document $Document

$Document = Get-OfficeWord -FilePath $PSScriptRoot\Documents\Doc1.docx

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline $null, 'Dash'

$DataTable = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'Nope' }
)

$Table = New-OfficeWordTable -Document $Document -DataTable $DataTable -TableLayout Autofit
$Table.Style = 'PlainTable5'

Save-OfficeWord -Document $Document -Show