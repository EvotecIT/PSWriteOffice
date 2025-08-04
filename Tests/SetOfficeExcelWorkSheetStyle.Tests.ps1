Describe 'Set-OfficeExcelWorkSheetStyle cmdlet' {
    It 'sets worksheet tab color' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        Set-OfficeExcelWorkSheetStyle -Excel $workbook -Worksheet $worksheet -TabColor 'Red'
        $expected = [PSWriteOffice.Services.ColorService]::GetColor('Red').ToHex()
        $worksheet.TabColor.ToHex() | Should -Be $expected
    }
}
