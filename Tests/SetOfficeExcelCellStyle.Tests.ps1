Describe 'Set-OfficeExcelCellStyle cmdlet' {
    It 'applies style to a cell' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        New-OfficeExcelValue -Worksheet $worksheet -Row 1 -Column 1 -Value 'Test'
        Set-OfficeExcelCellStyle -Worksheet $worksheet -Row 1 -Column 1 -Bold $true
        $worksheet.Cell(1,1).Style.Font.Bold | Should -BeTrue
    }
}
