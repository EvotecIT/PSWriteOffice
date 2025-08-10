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

    It 'imports data using specified culture' {
        $path = Join-Path $TestDrive 'culture.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Number'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value '1,23'
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $culture = [System.Globalization.CultureInfo]'pl-PL'
        $rows = Import-OfficeExcel -FilePath $path -Culture $culture
        $rows[0].Number | Should -Be 1.23
    }

    It 'imports data within specified range' {
        $path = Join-Path $TestDrive 'range.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Name'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 'Age'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'John'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 30
        New-OfficeExcelValue -Worksheet $sheet1 -Row 3 -Column 1 -Value 'Jane'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 3 -Column 2 -Value 25
        New-OfficeExcelValue -Worksheet $sheet1 -Row 4 -Column 1 -Value 'Bob'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 4 -Column 2 -Value 40
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $rows = Import-OfficeExcel -FilePath $path -StartRow 2 -EndRow 3 -StartColumn 1 -EndColumn 2
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'John'
        $rows[1].Name | Should -Be 'Jane'
    }

    It 'imports data using custom header row' {
        $path = Join-Path $TestDrive 'headerrow.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Skip'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 'Skip'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'Name'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 'Age'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 3 -Column 1 -Value 'John'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 3 -Column 2 -Value 30
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $rows = Import-OfficeExcel -FilePath $path -StartRow 2 -EndRow 3 -HeaderRow 2
        $rows[0].Name | Should -Be 'John'
        $rows[0].Age | Should -Be 30
    }

    It 'imports data without header row' {
        $path = Join-Path $TestDrive 'noheader.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'John'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 30
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'Jane'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 25
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $rows = Import-OfficeExcel -FilePath $path -NoHeader -StartRow 1 -EndRow 2 -StartColumn 1 -EndColumn 2
        $rows[0].Column1 | Should -Be 'John'
        $rows[0].Column2 | Should -Be 30
    }

    It 'imports data as specified type' {
        class Person {
            [string]$Name
            [int]$Age
        }

        $path = Join-Path $TestDrive 'typed.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Name'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 'Age'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'John'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 30
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $rows = Import-OfficeExcel -FilePath $path -Type ([Person])
        $rows[0].GetType().FullName | Should -Be ([Person]).FullName
        $rows[0].Name | Should -Be 'John'
        $rows[0].Age | Should -Be 30
    }

    It 'imports data as DataTable' {
        $path = Join-Path $TestDrive 'datatable.xlsx'
        $workbook = New-OfficeExcel
        $sheet1 = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 1 -Value 'Name'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 1 -Column 2 -Value 'Age'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 1 -Value 'Jane'
        New-OfficeExcelValue -Worksheet $sheet1 -Row 2 -Column 2 -Value 31
        Save-OfficeExcel -Workbook $workbook -FilePath $path
        $table = Import-OfficeExcel -FilePath $path -AsDataTable
        $table.GetType().FullName | Should -Be ([System.Data.DataTable]).FullName
        $table.Rows.Count | Should -Be 1
        $table.Rows[0].Name | Should -Be 'Jane'
    }

    It 'throws for invalid path' {
        { Import-OfficeExcel -FilePath (Join-Path $TestDrive 'missing.xlsx') -ErrorAction Stop } | Should -Throw
    }
}
