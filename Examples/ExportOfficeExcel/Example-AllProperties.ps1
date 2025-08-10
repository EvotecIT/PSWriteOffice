Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = Join-Path $PSScriptRoot 'AllProperties.xlsx'
$data = @(
    [PSCustomObject]@{ Name = 'Alice'; Age = 30 },
    [PSCustomObject]@{ Name = 'Bob' }
)
$data | Export-OfficeExcel -FilePath $path -AllProperties -Show
