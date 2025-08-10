Describe 'Set-OfficeExcelWorkSheetStyle cmdlet' {
    It 'sets worksheet tab color' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        Set-OfficeExcelWorkSheetStyle -Excel $workbook -Worksheet $worksheet -TabColor 'Red'
        # XLColor doesn't have ToHex(), but we can check if the color was set
        $worksheet.TabColor | Should -Not -BeNullOrEmpty
        # Check that it's red (RGB: 255,0,0)
        $worksheet.TabColor.Color.R | Should -Be 255
        $worksheet.TabColor.Color.G | Should -Be 0
        $worksheet.TabColor.Color.B | Should -Be 0
    }
}
