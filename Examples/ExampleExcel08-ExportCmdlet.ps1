Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = Join-Path $PSScriptRoot 'Documents/ExportCmdlet.xlsx'
$data = Get-Process | Select-Object -First 5
$data | Export-OfficeExcel -FilePath $path -WorksheetName 'Processes' -Show
