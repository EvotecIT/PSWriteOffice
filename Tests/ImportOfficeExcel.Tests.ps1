Describe 'Import-OfficeExcel cmdlet' {
    It 'imports specified worksheet data' {
        $path = Join-Path $TestDrive 'import.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Name'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 'Age'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'Jane'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 31
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Other' | Out-Null
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $rows = Import-OfficeExcel -FilePath $path -WorkSheetName 'Data'
        $rows[0].Name | Should -Be 'Jane'
        $rows[0].Age | Should -Be 31
    }

    It 'imports data using specified culture' {
        $path = Join-Path $TestDrive 'culture.xlsx'
        New-Item -Path $path -ItemType File | Out-Null
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Number'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value '1,23'
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $culture = [System.Globalization.CultureInfo]'pl-PL'
        $rows = Import-OfficeExcel -FilePath $path -Culture $culture
        $rows[0].Number | Should -Be 1.23
    }

    It 'throws for invalid path' {
        { Import-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') } | Should -Throw
    }
}
