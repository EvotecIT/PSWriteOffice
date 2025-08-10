Describe 'New-OfficeExcelTable cmdlet' {
    It 'creates table with data' {
        $workbook = New-OfficeExcel
        $worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace
        $data = @(
            [pscustomobject]@{ Name = 'A'; Value = 1 },
            [pscustomobject]@{ Name = 'B'; Value = 2 }
        )
        $table = New-OfficeExcelTable -Worksheet $worksheet -DataTable $data -StartRow 1 -StartColumn 1
        # RowCount includes header row, so 2 data rows + 1 header = 3
        $table.RowCount() | Should -Be 3
        $table.ColumnCount() | Should -Be 2
    }
}
