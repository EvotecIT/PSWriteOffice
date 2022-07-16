Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\TestTable.docx
New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline $null, 'Double'

$DataTable = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'Nope' }
)

$Table = New-OfficeWordTable -Document $Document -DataTable $DataTable -TableLayout Autofit #-Style GridTable1LightAccent1
foreach ($Cell in $Table.Rows[0].Cells) {
    $Cell.Paragraphs[0].Color = '#FF0000'
    $Cell.Paragraphs[0].Highlight = [DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues]::Magenta
}

New-OfficeWordText -Document $Document
New-OfficeWordText -Document $Document -Text "Another table below" -Bold $true -Color RedBerry -Alignment Center

New-OfficeWordTable -Document $Document -DataTable $DataTable -TableLayout Autofit -Style GridTable2 -SkipHeader

Save-OfficeWord -Document $Document -Show