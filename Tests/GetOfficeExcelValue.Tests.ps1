Describe 'Get-OfficeExcelValue cmdlet' {
    It 'returns cell object for given coordinates' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        $worksheet.Cell(2, 3).Value = 'data'
        $cell = Get-OfficeExcelValue -Worksheet $worksheet -Row 2 -Column 3
        $cell.Value | Should -Be 'data'
    }
}
