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

    It 'supports autofit and validation list helpers' {
        $path = Join-Path $TestDrive 'DslExcelExtras.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'New' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'Done' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Items' -AutoFit
                Add-OfficeExcelValidationList -Range 'C2:C3' -Values 'New','In Progress','Done'
                Invoke-OfficeExcelAutoFit -Columns
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports row/column helpers and reader metadata' {
        $path = Join-Path $TestDrive 'DslExcelReaders.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Value = 10 }
            [PSCustomObject]@{ Name = 'Beta'; Value = 20 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelColumn -Column 1 -StartRow 2 -Values 'Alpha', 'Beta'
                Set-OfficeExcelColumn -Column 2 -StartRow 2 -Values 10, 20
                Set-OfficeExcelNamedRange -Name 'ManualRange' -Range 'A1:B3'
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -StartRow 5
            }
        } | Out-Null

        $named = Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
        $named | Should -Not -BeNullOrEmpty

        $tables = Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales'
        $tables | Should -Not -BeNullOrEmpty

        $doc = Get-OfficeExcel -Path $path
        try {
            $doc | Save-OfficeExcel | Out-Null
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'sets named ranges, formulas, and header/footer' {
        $path = Join-Path $TestDrive 'DslExcelMeta.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 10
                Set-OfficeExcelCell -Address 'B1' -Value 20
                Set-OfficeExcelFormula -Address 'C1' -Formula 'SUM(A1:B1)'
                Set-OfficeExcelNamedRange -Name 'Totals' -Range 'A1:C1'
                Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N'
            }
        }

        Test-Path $path | Should -BeTrue
    }
}
