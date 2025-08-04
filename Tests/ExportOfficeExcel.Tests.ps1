Describe 'Export-OfficeExcel cmdlet' {
    It 'creates an Excel file with a worksheet and table' {
        $path = Join-Path $TestDrive 'test.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        Test-Path $path | Should -BeTrue
    }

    It 'throws for invalid path' {
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        { $data | Export-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') } | Should -Throw
    }
}
