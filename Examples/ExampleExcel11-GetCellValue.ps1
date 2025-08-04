$workbook = New-OfficeExcel
$worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
$worksheet.Cell(1, 1).Value = 'Example'

$cell = Get-OfficeExcelValue -Worksheet $worksheet -Row 1 -Column 1
$cell.Value
