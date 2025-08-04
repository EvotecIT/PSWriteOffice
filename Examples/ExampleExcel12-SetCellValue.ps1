$workbook = New-OfficeExcel
$worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
New-OfficeExcelValue -Worksheet $worksheet -Row 2 -Column 1 -Value 'Example'
