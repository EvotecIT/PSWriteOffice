Describe 'New-OfficeExcelTable cmdlet' {
    It 'creates table with data' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        $data = @(
            [pscustomobject]@{ Name = 'A'; Value = 1 },
            [pscustomobject]@{ Name = 'B'; Value = 2 }
        )
        $table = New-OfficeExcelTable -Worksheet $worksheet -DataTable $data -StartRow 1 -StartColumn 1
        $table.RowCount() | Should -Be 2
        $table.ColumnCount() | Should -Be 2
    }
}
