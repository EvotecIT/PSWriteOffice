$workbook = New-OfficeExcel
$worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
$worksheet.Cell(1,1).Value = 'Name'
$worksheet.Cell(1,2).Value = 'Age'
$worksheet.Cell(2,1).Value = 'Alice'
$worksheet.Cell(2,2).Value = 30
Get-OfficeExcelWorkSheetData -Worksheet $worksheet
