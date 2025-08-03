Describe 'Export-OfficeExcel cmdlet' {
    It 'creates an Excel file with a worksheet and table' {
        $path = Join-Path $TestDrive 'test.xlsx'
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        Test-Path $path | Should -BeTrue
    }
}
