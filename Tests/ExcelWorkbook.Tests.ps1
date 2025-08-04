Describe 'Excel workbook cmdlets' {
    It 'creates, saves, and loads workbook' {
        $path = Join-Path $TestDrive 'test.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $workbook = New-OfficeExcel
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace | Out-Null
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $loaded = Get-OfficeExcel -FilePath $path
        $loaded.Worksheets.Count | Should -Be 1
        Close-OfficeExcel -Workbook $loaded
    }

    It 'throws when saving to invalid path' {
        $workbook = New-OfficeExcel
        { Save-OfficeExcel -Workbook $workbook -FilePath (Join-Path $TestDrive 'missing.xlsx') } | Should -Throw
    }

    It 'throws when loading from invalid path' {
        { Get-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') } | Should -Throw
    }
}
