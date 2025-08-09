Describe 'Export-OfficeExcel cmdlet' {
    It 'creates an Excel file with a worksheet and table' {
        $path = Join-Path $TestDrive 'test.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        Test-Path $path | Should -BeTrue
    }

    It 'appends to an existing table when Append is used' {
        $path = Join-Path $TestDrive 'append.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $first = 1..2 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $first | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        $second = 3..4 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $second | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -Append
        $rows = Import-OfficeExcel -FilePath $path -WorkSheetName 'Data'
        $rows.Count | Should -Be 4
        $rows[-1].Value | Should -Be 4
    }

    It 'includes all properties when AllProperties is used' {
        $path = Join-Path $TestDrive 'allprops.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $data = @(
            [PSCustomObject]@{ First = 1; Second = 'A' },
            [PSCustomObject]@{ First = 2 }
        )
        $data | Export-OfficeExcel -FilePath $path -AllProperties
        $rows = Import-OfficeExcel -FilePath $path
        $rows[0].PSObject.Properties.Name | Should -Contain 'Second'
        $rows[1].PSObject.Properties.Name | Should -Contain 'Second'
        $rows[1].Second | Should -BeNullOrEmpty
    }

    It 'throws for invalid path' {
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        { $data | Export-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') } | Should -Throw
    }
}
