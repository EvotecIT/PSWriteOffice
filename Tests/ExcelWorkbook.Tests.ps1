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

    It 'throws when saving to invalid path' {
        $workbook = New-OfficeExcel
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' | Out-Null
        { Save-OfficeExcel -Workbook $workbook -FilePath 'C:\InvalidPath<>*|?"\missing.xlsx' -ErrorAction Stop } | Should -Throw
    }

    It 'throws when loading from invalid path' {
        { Get-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') -ErrorAction Stop } | Should -Throw
    }
        
    It 'supports -WhatIf parameter' {
        $path = Join-Path $TestDrive 'whatif.xlsx'
        $workbook = New-OfficeExcel
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' | Out-Null
        Save-OfficeExcel -Workbook $workbook -FilePath $path -WhatIf
        Test-Path $path | Should -BeFalse
        Close-OfficeExcel -Workbook $workbook -WhatIf
        { $workbook.Worksheets.Count } | Should -Not -Throw
    }
}
