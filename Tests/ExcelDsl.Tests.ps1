BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'Excel DSL surface' {
    It 'creates a workbook with canonical cmdlets' {
        $path = Join-Path $TestDrive 'DslExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Add-OfficeExcelTable -Data $rows -TableName 'Sales'
            }
        }

        Test-Path $path | Should -BeTrue

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -BeGreaterThan 0
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports alias-only syntax' {
        $path = Join-Path $TestDrive 'DslExcelAlias.xlsx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Laptop'; Qty = 5 }
            [PSCustomObject]@{ Item = 'Tablet'; Qty = 12 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Orders' -Content {
                ExcelCell -Address 'A1' -Value 'Item'
                ExcelCell -Address 'B1' -Value 'Qty'
                ExcelTable -Data $rows -TableName 'OrdersTable'
            }
        }

        Test-Path $path | Should -BeTrue
    }
}
