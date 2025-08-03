Import-Module PSWriteOffice -Force

$path = Join-Path $PSScriptRoot 'WorkbookExample.xlsx'

$workbook = New-OfficeExcel
New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace | Out-Null
Save-OfficeExcel -Workbook $workbook -FilePath $path

$loaded = Get-OfficeExcel -FilePath $path
Close-OfficeExcel -Workbook $loaded
