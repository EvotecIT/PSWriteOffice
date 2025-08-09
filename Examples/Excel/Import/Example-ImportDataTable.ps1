Import-Module ../../PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot/DataTable.xlsx"
$workbook = New-OfficeExcel
$sheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 1 -Value 'Name'
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 2 -Value 'Age'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 1 -Value 'Jane'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 2 -Value 31
Save-OfficeExcel -Workbook $workbook -FilePath $path

$table = Import-OfficeExcel -FilePath $path -AsDataTable
$table | Format-Table
