Describe 'Import-OfficeExcel cmdlet' {
    It 'imports specified worksheet data' {
        $path = Join-Path $TestDrive 'import.xlsx'
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
}
