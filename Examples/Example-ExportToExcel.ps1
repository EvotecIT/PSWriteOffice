Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Data1 = Get-Process | Select-Object -First 15
$Objects = @(
    [PSCustomObject] @{ Test = 1; DateTime = (Get-Date); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
    [PSCustomObject] @{ Test = 1; }
    [PSCustomObject] @{ Test = 3; DateTime = (Get-Date).AddDays(1); TimeSpan = (New-TimeSpan -Minutes 10); TestString = 'string' }
)

$Objects | Export-OfficeExcel -FilePath "$PSScriptRoot\Documents\ExportToExcel.xlsx" -WorksheetName "Sheet3" 
$Data1 | Export-OfficeExcel -FilePath "$PSScriptRoot\Documents\ExportToExcel.xlsx" -WorksheetName "Sheet4" -Show