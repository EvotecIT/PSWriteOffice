Import-Module PSWriteOffice -Force

$workbook = New-OfficeExcel
New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'First' -Option Replace | Out-Null
New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Second' -Option Replace | Out-Null

# Retrieve by name
Get-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'First'

# Retrieve by index
Get-OfficeExcelWorkSheet -Workbook $workbook -Index 2

# Retrieve all worksheet names
Get-OfficeExcelWorkSheet -Workbook $workbook -NameOnly
