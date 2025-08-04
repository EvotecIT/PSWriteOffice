Describe 'Get-OfficeExcelWorkSheet cmdlet' {
    It 'retrieves worksheets by name, index, and all' {
        $workbook = New-OfficeExcel
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace | Out-Null
        New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet2' -Option Replace | Out-Null

        $wsByName = Get-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1'
        $wsByName.Name | Should -Be 'Sheet1'

        $wsByIndex = Get-OfficeExcelWorkSheet -Workbook $workbook -Index 2
        $wsByIndex.Name | Should -Be 'Sheet2'

        $all = Get-OfficeExcelWorkSheet -Workbook $workbook
        ($all | Measure-Object).Count | Should -Be 2

        $names = Get-OfficeExcelWorkSheet -Workbook $workbook -NameOnly
        $names | Should -Contain 'Sheet1'
        $names | Should -Contain 'Sheet2'
    }
}
