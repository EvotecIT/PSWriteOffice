Describe 'New-OfficeExcelValue cmdlet' {
    It 'writes value to a cell' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        New-OfficeExcelValue -Worksheet $worksheet -Row 3 -Column 2 -Value 'hello'
        $worksheet.Cell(3,2).Value | Should -Be 'hello'
    }
}
