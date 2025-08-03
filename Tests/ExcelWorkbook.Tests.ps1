Describe 'Excel workbook cmdlets' {
    It 'creates, saves, and loads workbook' {
        $path = Join-Path $TestDrive 'test.xlsx'
        $workbook = New-OfficeExcel
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace | Out-Null
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $loaded = Get-OfficeExcel -FilePath $path
        $loaded.Worksheets.Count | Should -Be 1
        Close-OfficeExcel -Workbook $loaded
    }
}
