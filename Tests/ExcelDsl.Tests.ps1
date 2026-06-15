BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    function Get-TestLoadedType {
        param(
            [Parameter(Mandatory)]
            [string] $Name
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($Name, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        if ($null -eq $type) {
            throw "Unable to find loaded type '$Name'."
        }

        $type
    }

    function Test-OfficeLoadedMethod {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,

            [Parameter(Mandatory)]
            [string] $MethodName
        )

        $type = Get-TestLoadedType -Name $TypeName
        @($type.GetMethods() | Where-Object Name -eq $MethodName).Count -gt 0
    }
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
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
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

    It 'supports transposed Excel tables' {
        $path = Join-Path $TestDrive 'TransposedExcelTable.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'Europe'; Revenue = 21704714 }
            [PSCustomObject]@{ Region = 'Asia'; Revenue = 8774099 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -View Transpose -TableName 'TransposedSales'
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C3')
        $imported[0].Property | Should -Be 'Region'
        $imported[0].Row1 | Should -Be 'Europe'
        $imported[0].Row2 | Should -Be 'Asia'
        $imported[1].Property | Should -Be 'Revenue'
        $imported[1].Row1 | Should -Be 21704714
        $imported[1].Row2 | Should -Be 8774099
    }

    It 'supports transposed Excel tables from IDataReader input' {
        $path = Join-Path $TestDrive 'TransposedExcelReaderTable.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('One', 1)
        [void] $table.Rows.Add('Two', 2)
        $reader = $table.CreateDataReader()

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $reader -View Transpose -TableName 'TransposedReader'
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C3')
        $imported[0].Property | Should -Be 'Name'
        $imported[0].Row1 | Should -Be 'One'
        $imported[0].Row2 | Should -Be 'Two'
        $imported[1].Property | Should -Be 'Value'
        $imported[1].Row1 | Should -Be 1
        $imported[1].Row2 | Should -Be 2
    }

    It 'round-trips encrypted workbooks through lifecycle cmdlets' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.Excel.ExcelDocument' -MethodName 'LoadEncrypted')) {
            (Get-Command New-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Save-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Get-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedExcel.xlsx'

        New-OfficeExcel -Path $path -Password 'secret' -SafePreflight {
            Set-OfficeExcelExecutionPolicy -Mode Sequential -ParallelThreshold 5 -WorksheetValidation Always -Diagnostics
            Add-OfficeExcelSheet -Name 'Secure' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Encrypted value'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw

        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'Secure'
            $value = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $value) | Should -BeTrue
            $value | Should -Be 'Encrypted value'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'configures Excel execution policy from PowerShell' {
        $path = Join-Path $TestDrive 'ExcelExecutionPolicy.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Policy'
        }
        $doc = Get-OfficeExcel -Path $path
        try {
            $result = $doc | Set-OfficeExcelExecutionPolicy -Mode Parallel -ParallelThreshold 3 -MaxDegreeOfParallelism 2 -WorksheetValidation Disabled -Diagnostics -DisableAutoFitImmediateSave -PassThru

            $result | Should -Be $doc
            $doc.Execution.Mode.ToString() | Should -Be 'Parallel'
            $doc.Execution.ParallelThreshold | Should -Be 3
            $doc.Execution.MaxDegreeOfParallelism | Should -Be 2
            $doc.Execution.WorksheetValidation.ToString() | Should -Be 'Disabled'
            $doc.Execution.DiagnosticsRequested | Should -BeTrue
            $doc.Execution.SaveWorksheetAfterAutoFit | Should -BeFalse
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
                ExcelTable -InputObject $rows -TableName 'OrdersTable'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'preserves legacy Excel table data parameter aliases' {
        $path = Join-Path $TestDrive 'DslExcelTableDataAliases.xlsx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Laptop'; Qty = 5 }
            [PSCustomObject]@{ Item = 'Tablet'; Qty = 12 }
        )

        $table = [System.Data.DataTable]::new('Stock')
        [void] $table.Columns.Add('Item', [string])
        [void] $table.Columns.Add('Qty', [int])
        [void] $table.Rows.Add('Dock', 3)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'DataAlias' -Content {
                ExcelTable -Data $rows -TableName 'RowsAlias'
            }
            Add-OfficeExcelSheet -Name 'DataTableAlias' -Content {
                ExcelTable -DataTable $table -TableName 'TableAlias'
            }
        }

        $rowsAlias = @(Import-OfficeExcel -Path $path -WorksheetName 'DataAlias' -Range 'A1:B3')
        $tableAlias = @(Import-OfficeExcel -Path $path -WorksheetName 'DataTableAlias' -Range 'A1:B2')
        $rowsAlias[0].Item | Should -Be 'Laptop'
        $rowsAlias[1].Qty | Should -Be 12
        $tableAlias[0].Item | Should -Be 'Dock'
        $tableAlias[0].Qty | Should -Be 3
    }

    It 'writes a DataTable directly as an Excel table' {
        $path = Join-Path $TestDrive 'DslExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('People')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Score', [int])
        [void] $table.Rows.Add('Ada', 10)
        [void] $table.Rows.Add('Grace', 20)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $table -TableName 'PeopleTable' -AutoFit
            }
        }

        Test-Path $path | Should -BeTrue
        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'PeopleTable')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:B3'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Ada'
        $imported[0].Score | Should -Be 10
    }

    It 'appends rows to an existing Excel table outside the DSL' {
        $path = Join-Path $TestDrive 'ExcelExistingTableAppend.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        $doc = Get-OfficeExcel -Path $path
        try {
            $table = $doc | Add-OfficeExcelTableRow -Sheet Data -TableName Sales -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 }) -PassThru
            $table.Range | Should -Be 'A1:B4'

            Close-OfficeExcel -Document $doc -Save
            $doc = $null
        } finally {
            if ($null -ne $doc) {
                Close-OfficeExcel -Document $doc
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $imported.Count | Should -Be 3
        $imported[2].Region | Should -Be 'APAC'
        $imported[2].Revenue | Should -Be 300
    }

    It 'finds a named Excel table on later sheets when appending without a sheet filter' {
        $path = Join-Path $TestDrive 'ExcelExistingTableAppendWithoutSheet.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Overview'
            }
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        Add-OfficeExcelTableRow -Path $path -TableName Sales -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 })

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[1].Region | Should -Be 'APAC'
        $imported[1].Revenue | Should -Be 300
    }

    It 'appends explicit input to each piped Excel table target' {
        $path1 = Join-Path $TestDrive 'ExcelPipedTableAppend1.xlsx'
        $path2 = Join-Path $TestDrive 'ExcelPipedTableAppend2.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        foreach ($path in @($path1, $path2)) {
            New-OfficeExcel -Path $path {
                Add-OfficeExcelSheet -Name 'Data' -Content {
                    Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                }
            }
        }

        $doc1 = Get-OfficeExcel -Path $path1
        $doc2 = Get-OfficeExcel -Path $path2
        try {
            $table1 = $doc1.Sheets[0].Table('Sales')
            $table2 = $doc2.Sheets[0].Table('Sales')

            @($table1, $table2) |
                Add-OfficeExcelTableRow -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 })

            Close-OfficeExcel -Document $doc1 -Save
            Close-OfficeExcel -Document $doc2 -Save
            $doc1 = $null
            $doc2 = $null
        } finally {
            if ($null -ne $doc1) {
                Close-OfficeExcel -Document $doc1
            }
            if ($null -ne $doc2) {
                Close-OfficeExcel -Document $doc2
            }
        }

        foreach ($path in @($path1, $path2)) {
            $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
            $imported.Count | Should -Be 2
            $imported[1].Region | Should -Be 'APAC'
            $imported[1].Revenue | Should -Be 300
        }
    }

    It 'writes a DataSet as one worksheet per table' {
        $path = Join-Path $TestDrive 'DslExcelDataSet.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')

        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $sales.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($sales)

        $notes = [System.Data.DataTable]::new('Notes')
        [void] $notes.Columns.Add('Text', [string])
        [void] $notes.Rows.Add('Checked')
        [void] $dataSet.Tables.Add($notes)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelDataSet -DataSet $dataSet -AutoFit
        }

        Test-Path $path | Should -BeTrue
        $salesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $salesRows.Count | Should -Be 2
        $salesRows[1].Region | Should -Be 'EMEA'
        $salesRows[1].Revenue | Should -Be 200

        $notesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Notes' -Range 'A1:A2')
        $notesRows.Count | Should -Be 1
        $notesRows[0].Text | Should -Be 'Checked'
    }

    It 'exports and imports objects through operator cmdlets' {
        $path = Join-Path $TestDrive 'ExportOfficeExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100; Internal = 'skip' }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200; Internal = 'skip' }
        )

        $file = $rows |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -AutoFit -FreezeTopRow -BoldTopRow -ExcludeProperty Internal -PassThru

        $file.FullName | Should -Be $path
        Test-Path $path | Should -BeTrue

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A2:B4'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A2:B4')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'Internal'

        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 4 -EndRow 2 -StartColumn 1 -EndColumn 2 } |
            Should -Throw '*StartRow must be less than or equal to EndRow*'
        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 2 -EndRow 4 -StartColumn 3 -EndColumn 1 } |
            Should -Throw '*StartColumn must be less than or equal to EndColumn*'
    }

    It 'exports plain objects through the default table path' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDefaultObjects.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100; Created = [DateTime] '2026-01-01'; Enabled = $true }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200; Created = [DateTime] '2026-01-02'; Enabled = $false }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales'

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:D3'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:D3')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[1].Region | Should -Be 'EMEA'
        $imported[1].Revenue | Should -Be 200
        $imported[1].Enabled | Should -BeFalse
    }

    It 'appends rows without rewriting headers' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppend.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -AutoFit
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -AutoFit

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $imported.Count | Should -Be 3
        $imported[2].Region | Should -Be 'APAC'
        $imported[2].Revenue | Should -Be 150

        $excelSheetType = Get-TestLoadedType -Name 'OfficeIMO.Excel.ExcelSheet'
        $hasTableAppend = @($excelSheetType.GetMethods() | Where-Object Name -eq 'AppendDataTableToTable').Count -gt 0
        if ($hasTableAppend) {
            $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
            $tables.Count | Should -Be 1
            $tables[0].Range | Should -Be 'A1:B4'
        }
    }

    It 'exports DataTable input without exposing DataRow metadata' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('Sales')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $table.Rows.Add('EMEA', 200)

        Export-OfficeExcel -Path $path -InputObject $table -WorksheetName 'Data' -TableName 'Sales' -AutoFit

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'RowError'
    }

    It 'exports IDataReader input without requiring callers to buffer it first' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReader.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('A', 1)
        [void] $table.Rows.Add('B', 2)
        $reader = $table.CreateDataReader()

        Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -TableName 'SqlRows' -AutoFit -FreezeTopRow

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'A'
        $rows[1].Value | Should -Be 2
    }

    It 'exports HTML-parser DataTable output with companion link URL columns' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelHtmlDataTable.xlsx'
        $table = [System.Data.DataTable]::new('HtmlLinks')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('NameUrl', [string])
        [void] $table.Columns.Add('Status', [string])
        [void] $table.Rows.Add('Alpha', 'https://example.com/a', 'Ready')
        [void] $table.Rows.Add('Beta', 'https://example.com/b', 'Hold')

        $table | Export-OfficeExcel -Path $path -WorksheetName 'Links' -TableName 'HtmlLinks' -AutoFit -FreezeTopRow

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Links' -Range 'A1:C3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Alpha'
        $imported[0].NameUrl | Should -Be 'https://example.com/a'
        $imported[1].Status | Should -Be 'Hold'

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'HtmlLinks')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:C3'
    }

    It 'exports DataSet input as one worksheet per table' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSet.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')

        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($sales)

        $inventory = [System.Data.DataTable]::new('Inventory')
        [void] $inventory.Columns.Add('Item', [string])
        [void] $inventory.Columns.Add('Count', [int])
        [void] $inventory.Rows.Add('Laptop', 5)
        [void] $dataSet.Tables.Add($inventory)

        Export-OfficeExcel -Path $path -InputObject $dataSet -TableName 'IgnoredForDataSet' -AutoFit

        $salesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $inventoryRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Inventory' -Range 'A1:B2')

        $salesRows.Count | Should -Be 1
        $salesRows[0].Region | Should -Be 'NA'
        $salesRows[0].Revenue | Should -Be 100
        $inventoryRows.Count | Should -Be 1
        $inventoryRows[0].Item | Should -Be 'Laptop'
        $inventoryRows[0].Count | Should -Be 5
    }

    It 'exports DataSet tables one at a time without mutating source tables' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetNamespaceDuplicates.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')
        $dataSet.Namespace = 'urn:report'

        $inheritedNamespace = [System.Data.DataTable]::new('T')
        [void] $inheritedNamespace.Columns.Add('Name', [string])
        [void] $inheritedNamespace.Columns.Add('Secret', [string])
        [void] $inheritedNamespace.Rows.Add('Inherited', 'one')
        [void] $dataSet.Tables.Add($inheritedNamespace)

        $emptyNamespace = [System.Data.DataTable]::new('T')
        $emptyNamespace.Namespace = ''
        [void] $emptyNamespace.Columns.Add('Name', [string])
        [void] $emptyNamespace.Columns.Add('Secret', [string])
        [void] $emptyNamespace.Rows.Add('Empty', 'two')
        [void] $dataSet.Tables.Add($emptyNamespace)

        Export-OfficeExcel -Path $path -InputObject $dataSet -ExcludeProperty Secret

        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'T' -Range 'A1:A2')
        $secondRows = @(Import-OfficeExcel -Path $path -WorksheetName 'T (2)' -Range 'A1:A2')

        $firstRows.Count | Should -Be 1
        $firstRows[0].Name | Should -Be 'Inherited'
        $secondRows.Count | Should -Be 1
        $secondRows[0].Name | Should -Be 'Empty'
        $inheritedNamespace.Columns.Contains('Secret') | Should -BeTrue
        $emptyNamespace.Columns.Contains('Secret') | Should -BeTrue
    }

    It 'appends and clears DataSet sheets using sanitized worksheet names' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSanitizedAppend.xlsx'

        $dataSet = [System.Data.DataSet]::new('Report')
        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($sales)

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendSales = [System.Data.DataTable]::new('Sales:2026')
        [void] $appendSales.Columns.Add('Region', [string])
        [void] $appendSales.Columns.Add('Revenue', [int])
        [void] $appendSales.Rows.Add('EMEA', 200)
        [void] $appendSet.Tables.Add($appendSales)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $rows.Count | Should -Be 2
        $rows[1].Region | Should -Be 'EMEA'
        $rows[1].Revenue | Should -Be 200

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementSales = [System.Data.DataTable]::new('Sales:2026')
        [void] $replacementSales.Columns.Add('Region', [string])
        [void] $replacementSales.Columns.Add('Revenue', [int])
        [void] $replacementSales.Rows.Add('APAC', 300)
        [void] $replacementSet.Tables.Add($replacementSales)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replaced = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $replaced.Count | Should -Be 1
        $replaced[0].Region | Should -Be 'APAC'
        $replaced[0].Revenue | Should -Be 300
    }

    It 'keeps sanitized symbol DataSet sheet names distinct from existing workbook sheets' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetFallbackCollision.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Sheet1' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Existing'
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $table = [System.Data.DataTable]::new(':')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($table)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $existingText = $null
            $doc['Sheet1'].TryGetCellText(1, 1, [ref] $existingText) | Should -BeTrue
            $existingText | Should -Be 'Existing'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B2')
        $rows.Count | Should -Be 1
        $rows[0].Region | Should -Be 'NA'
        $rows[0].Revenue | Should -Be 100

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendTable = [System.Data.DataTable]::new(':')
        [void] $appendTable.Columns.Add('Region', [string])
        [void] $appendTable.Columns.Add('Revenue', [int])
        [void] $appendTable.Rows.Add('EMEA', 200)
        [void] $appendSet.Tables.Add($appendTable)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $appendedRows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B3')
        $appendedRows.Count | Should -Be 2
        $appendedRows[1].Region | Should -Be 'EMEA'
        $appendedRows[1].Revenue | Should -Be 200

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementTable = [System.Data.DataTable]::new(':')
        [void] $replacementTable.Columns.Add('Region', [string])
        [void] $replacementTable.Columns.Add('Revenue', [int])
        [void] $replacementTable.Rows.Add('APAC', 300)
        [void] $replacementSet.Tables.Add($replacementTable)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replacedRows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B2')
        $replacedRows.Count | Should -Be 1
        $replacedRows[0].Region | Should -Be 'APAC'
        $replacedRows[0].Revenue | Should -Be 300

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $existingText = $null
            $doc['Sheet1'].TryGetCellText(1, 1, [ref] $existingText) | Should -BeTrue
            $existingText | Should -Be 'Existing'
            $doc.Sheets.Name | Should -Not -Contain '_ (2)'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'reuses existing suffixed DataSet sheets when appending and clearing sanitized duplicates' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetDuplicateSanitizedAppend.xlsx'

        $dataSet = [System.Data.DataSet]::new('Report')
        $salesColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $salesColon.Columns.Add('Region', [string])
        [void] $salesColon.Columns.Add('Revenue', [int])
        [void] $salesColon.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($salesColon)

        $salesSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $salesSlash.Columns.Add('Region', [string])
        [void] $salesSlash.Columns.Add('Revenue', [int])
        [void] $salesSlash.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($salesSlash)

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $appendColon.Columns.Add('Region', [string])
        [void] $appendColon.Columns.Add('Revenue', [int])
        [void] $appendColon.Rows.Add('APAC', 300)
        [void] $appendSet.Tables.Add($appendColon)

        $appendSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $appendSlash.Columns.Add('Region', [string])
        [void] $appendSlash.Columns.Add('Revenue', [int])
        [void] $appendSlash.Rows.Add('LATAM', 400)
        [void] $appendSet.Tables.Add($appendSlash)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $secondRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026 (2)' -Range 'A1:B3')
        $firstRows.Count | Should -Be 2
        $secondRows.Count | Should -Be 2
        $firstRows[1].Region | Should -Be 'APAC'
        $secondRows[1].Region | Should -Be 'LATAM'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Not -Contain 'Sales_2026 (3)'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $replacementColon.Columns.Add('Region', [string])
        [void] $replacementColon.Columns.Add('Revenue', [int])
        [void] $replacementColon.Rows.Add('NA', 500)
        [void] $replacementSet.Tables.Add($replacementColon)

        $replacementSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $replacementSlash.Columns.Add('Region', [string])
        [void] $replacementSlash.Columns.Add('Revenue', [int])
        [void] $replacementSlash.Rows.Add('EMEA', 600)
        [void] $replacementSet.Tables.Add($replacementSlash)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replacedFirst = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $replacedSecond = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026 (2)' -Range 'A1:B2')
        $replacedFirst.Count | Should -Be 1
        $replacedSecond.Count | Should -Be 1
        $replacedFirst[0].Revenue | Should -Be 500
        $replacedSecond[0].Revenue | Should -Be 600
    }

    It 'preserves underscore-distinct DataSet sheet names and sparse suffix matches' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSparseSuffixes.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Q1_Ops' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Existing'
            }
            Add-OfficeExcelSheet -Name 'Sparse' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
            }
            Add-OfficeExcelSheet -Name 'Sparse (10)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $underscored = [System.Data.DataTable]::new('Q1__Ops')
        [void] $underscored.Columns.Add('Region', [string])
        [void] $underscored.Columns.Add('Revenue', [int])
        [void] $underscored.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($underscored)

        $sparseFirst = [System.Data.DataTable]::new('Sparse ')
        [void] $sparseFirst.Columns.Add('Region', [string])
        [void] $sparseFirst.Columns.Add('Revenue', [int])
        [void] $sparseFirst.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($sparseFirst)

        $sparseSecond = [System.Data.DataTable]::new(' Sparse')
        [void] $sparseSecond.Columns.Add('Region', [string])
        [void] $sparseSecond.Columns.Add('Revenue', [int])
        [void] $sparseSecond.Rows.Add('APAC', 300)
        [void] $dataSet.Tables.Add($sparseSecond)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Contain 'Q1__Ops'
            $doc.Sheets.Name | Should -Not -Contain 'Q1_Ops (2)'
            $q1Text = $null
            $doc['Q1_Ops'].TryGetCellText(1, 1, [ref] $q1Text) | Should -BeTrue
            $q1Text | Should -Be 'Existing'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $underscoredRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Q1__Ops' -Range 'A1:B2')
        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sparse' -Range 'A1:B2')
        $sparseRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sparse (10)' -Range 'A1:B2')

        $underscoredRows.Count | Should -Be 1
        $underscoredRows[0].Revenue | Should -Be 100
        $firstRows.Count | Should -Be 1
        $firstRows[0].Revenue | Should -Be 200
        $sparseRows.Count | Should -Be 1
        $sparseRows[0].Revenue | Should -Be 300
    }

    It 'matches the lowest suffixed DataSet sheet independent of workbook order' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetLowestSuffix.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data (3)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Set-OfficeExcelCell -Address 'A2' -Value 'APAC'
                Set-OfficeExcelCell -Address 'B2' -Value 300
            }
            Add-OfficeExcelSheet -Name 'Data (2)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Set-OfficeExcelCell -Address 'A2' -Value 'EMEA'
                Set-OfficeExcelCell -Address 'B2' -Value 200
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $table = [System.Data.DataTable]::new('Data')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($table)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $lowestRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data (2)' -Range 'A1:B3')
        $higherRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data (3)' -Range 'A1:B2')

        $lowestRows.Count | Should -Be 2
        $lowestRows[1].Region | Should -Be 'NA'
        $lowestRows[1].Revenue | Should -Be 100
        $higherRows.Count | Should -Be 1
        $higherRows[0].Revenue | Should -Be 300
    }

    It 'preserves symbol-only DataSet sheet names and sanitizes control characters' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSymbolNames.xlsx'
        $controlName = "Bad$([char]1)Name"

        $dataSet = [System.Data.DataSet]::new('Report')
        foreach ($name in @('---', '___', $controlName)) {
            $table = [System.Data.DataTable]::new($name)
            [void] $table.Columns.Add('Region', [string])
            [void] $table.Columns.Add('Revenue', [int])
            [void] $table.Rows.Add('NA', 100)
            [void] $dataSet.Tables.Add($table)
        }

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Contain '---'
            $doc.Sheets.Name | Should -Contain '___'
            $doc.Sheets.Name | Should -Contain 'Bad_Name'
            $doc.Sheets.Name | Should -Not -Contain 'Sheet1'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'adds a DataTable inside the Excel DSL table command' {
        $path = Join-Path $TestDrive 'DslExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('Items')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Quantity', [int])
        [void] $table.Rows.Add('Laptop', 5)
        [void] $table.Rows.Add('Tablet', 12)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $table -TableName 'Items' -AutoFit
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Laptop'
        $imported[0].Quantity | Should -Be 5
    }

    It 'lets OfficeIMO generate table names when the DSL caller omits them' {
        $path = Join-Path $TestDrive 'DslExcelGeneratedTableNames.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'First' -Content {
                Add-OfficeExcelTable -InputObject $rows
            }
            Add-OfficeExcelSheet -Name 'Second' -Content {
                Add-OfficeExcelTable -InputObject $rows
            }
        }

        $tables = @(Get-OfficeExcelTable -Path $path)
        $tables.Count | Should -Be 2
        @($tables.Name | Select-Object -Unique).Count | Should -Be 2
    }

    It 'keeps append freeze panes anchored to the existing table header' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppendFreeze.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -FreezeTopRow
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -FreezeTopRow

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $pane = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetViews']/*[local-name()='sheetView']/*[local-name()='pane']")

        $pane.GetAttribute('ySplit') | Should -Be '2'
    }

    It 'supports autofit and validation list helpers' {
        $path = Join-Path $TestDrive 'DslExcelExtras.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'New' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'Done' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Items' -AutoFit
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
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -StartRow 5
            }
        } | Out-Null

        $named = Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
        $named | Should -Not -BeNullOrEmpty

        $namedRangeType = Get-TestLoadedType -Name 'PSWriteOffice.Cmdlets.Excel.GetOfficeExcelNamedRangeCommand'
        $normalizeRange = $namedRangeType.GetMethod('NormalizeRange', [System.Reflection.BindingFlags] 'NonPublic, Static')
        $normalizeRange.Invoke($null, @("'Budget`$2026'!`$A`$1:`$B`$2")) | Should -Be "'Budget`$2026'!A1:B2"

        $tables = Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales'
        $tables | Should -Not -BeNullOrEmpty

        $namedRows = @($named | Import-OfficeExcel)
        $namedRows.Count | Should -Be 2
        $namedRows[0].Name | Should -Be 'Alpha'
        $namedRows[0].Value | Should -Be 10

        $tableRows = @($tables | Import-OfficeExcel)
        $tableRows.Count | Should -Be 2
        $tableRows[1].Name | Should -Be 'Beta'
        $tableRows[1].Value | Should -Be 20

        $doc = Get-OfficeExcel -Path $path
        try {
            $documentRows = @($doc | Import-OfficeExcel -Sheet 'Data' -Range 'A1:B3')
            $documentRows.Count | Should -Be 2
            $documentRows[0].Name | Should -Be 'Alpha'

            $doc | Save-OfficeExcel | Out-Null
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $server = Start-TestHttpFileServer -FilePath $path -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' -RequestCount 8
        try {
            $uri = [uri] $server.Url

            $remoteRows = @(Import-OfficeExcel -Uri $uri -AllowHttp -Sheet 'Data' -Range 'A1:B3')
            $remoteRows.Count | Should -Be 2
            $remoteRows[1].Name | Should -Be 'Beta'

            $remoteRange = @(Get-OfficeExcelRange -Uri $uri -AllowHttp -Sheet 'Data' -Range 'A1:B3')
            $remoteRange.Count | Should -Be 2
            $remoteRange[0].Value | Should -Be 10

            $remoteUsedRange = Get-OfficeExcelUsedRange -Uri $uri -AllowHttp -Sheet 'Data' -AsDataTable
            $remoteUsedRange.Rows.Count | Should -Be 6

            $remoteTables = Get-OfficeExcelTable -Uri $uri -AllowHttp | Where-Object Name -eq 'Sales'
            $remoteTableRows = @($remoteTables | Import-OfficeExcel -AllowHttp)
            $remoteTableRows.Count | Should -Be 2

            $remoteNamed = Get-OfficeExcelNamedRange -Uri $uri -AllowHttp -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
            $remoteNamedRows = @($remoteNamed | Import-OfficeExcel -AllowHttp)
            $remoteNamedRows.Count | Should -Be 2

            $remoteDoc = Get-OfficeExcel -Uri $uri -AllowHttp -ReadOnly
            try {
                $remoteDocRows = @($remoteDoc | Import-OfficeExcel -Sheet 'Data' -Range 'A1:B3')
                $remoteDocRows.Count | Should -Be 2
            } finally {
                Close-OfficeExcel -Document $remoteDoc
            }
        } finally {
            Stop-TestHttpFileServer -Server $server
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

    It 'supports advanced Excel data helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedData.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelAutoFilter -Range 'A1:F4'
                Invoke-OfficeExcelSort -Header 'Region'
                Set-OfficeExcelFreeze -TopRows 1
                Add-OfficeExcelConditionalRule -Range 'B2:B4' -Operator GreaterThan -Formula1 '150'
                Add-OfficeExcelConditionalDataBar -Range 'B2:B4' -Color '#92D050'
                Add-OfficeExcelConditionalColorScale -Range 'C2:C4' -StartColor '#FEE599' -EndColor '#6AA84F'
                Add-OfficeExcelConditionalIconSet -Range 'C2:C4'
                Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type ColumnClustered -Title 'Sales'
                Add-OfficeExcelImage -Address 'I1' -Path $imagePath -WidthPixels 64 -HeightPixels 64
                Set-OfficeExcelHyperlink -Address 'A2' -Url 'https://example.org' -Display 'Example'
                Add-OfficeExcelComment -Address 'B2' -Text 'Check sales'
                Add-OfficeExcelSparkline -DataRange 'B2:B4' -LocationRange 'H2:H4' -Type Column
            }
        }

        Test-Path $path | Should -BeTrue

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $sheet = $doc.Sheets[0]
            $sheet.Name | Should -Be 'Data'
            $sheet.HasComment(2, 2) | Should -BeTrue

            $cellText = $null
            $sheet.TryGetCellText(2, 1, [ref] $cellText) | Should -BeTrue
            $cellText | Should -Be 'Example'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports advanced Excel pivot, validation, and protection helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedPivot.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelPivotTable -SourceRange 'A1:F4' -DestinationCell 'J1' -RowField 'Region' -DataField 'Sales' -DataDisplayName 'Total Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000 -AllowBlank:$false
                Add-OfficeExcelValidationDecimal -Range 'C2:C4' -Operator Between -Formula1 0.0 -Formula2 1.0
                Add-OfficeExcelValidationDate -Range 'D2:D4' -Operator GreaterThan -Formula1 ([datetime]'2024-01-01')
                Add-OfficeExcelValidationTime -Range 'E2:E4' -Operator GreaterThan -Formula1 ([TimeSpan]'08:00:00')
                Add-OfficeExcelValidationTextLength -Range 'F2:F4' -Operator Between -Formula1 1 -Formula2 20
                Add-OfficeExcelValidationCustomFormula -Range 'G2:G4' -Formula 'LEN(A2)>0'
                Protect-OfficeExcelSheet
                Unprotect-OfficeExcelSheet
                Protect-OfficeExcelSheet
            }
        }

        Test-Path $path | Should -BeTrue

        $pivotTables = @(Get-OfficeExcelPivotTable -Path $path -Name 'PivotTable')
        $pivotTables.Count | Should -Be 1

        $pivot = $pivotTables[0]
        $pivot.SourceRange | Should -Be 'A1:F4'
        $pivot.Location | Should -Match '^J1:[A-Z]+\d+$'
        $pivot.RowFields | Should -Contain 'Region'
        @($pivot.DataFields).Count | Should -BeGreaterThan 0
        $pivot.DataFields[0].FieldName | Should -Be 'Sales'
        $pivot.DataFields[0].DisplayName | Should -Be 'Total Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $doc.Sheets[0].IsProtected | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'uses OfficeIMO pivot field options and captions' {
        $path = Join-Path $TestDrive 'DslExcelPivotOptions.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Product = 'Standard'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Product = 'Standard'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Product = 'Legacy'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelPivotTable -SourceRange 'A1:C4' -DestinationCell 'E1' -RowField 'Region' -PageField 'Product' -DataField 'Sales' -DataNumberFormat '#,##0' -GrandTotalCaption 'Overall' -FieldSort @{ Region = 'Ascending' } -FieldHiddenItems @{ Region = @('APAC') } -PageFieldSelection @{ Product = 'Standard' }
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports advanced Excel page setup and visibility helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedLayout.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0
                Set-OfficeExcelMargins -Preset Narrow
                Set-OfficeExcelOrientation -Orientation Landscape
                Set-OfficeExcelGridlines -Hide
                Set-OfficeExcelSheetVisibility -Hide
            }
        }

        Test-Path $path | Should -BeTrue

        $workbookXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/workbook.xml'
        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'

        $workbookSheet = $workbookXml.SelectSingleNode("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']")
        $workbookSheet.GetAttribute('name') | Should -Be 'Data'
        $workbookSheet.GetAttribute('state') | Should -Be 'hidden'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.HiddenSheetCount | Should -Be 1
        $summary.Sheets[0].State | Should -Be 'Hidden'

        $pageSetup = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageSetup']")
        $pageSetup.GetAttribute('fitToWidth') | Should -Be '1'
        $pageSetup.GetAttribute('fitToHeight') | Should -Be '0'
        $pageSetup.GetAttribute('orientation') | Should -Be 'landscape'

        $pageMargins = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageMargins']")
        $pageMargins.GetAttribute('left') | Should -Be '0.25'
        $pageMargins.GetAttribute('right') | Should -Be '0.25'
        $pageMargins.GetAttribute('top') | Should -Be '0.5'
        $pageMargins.GetAttribute('bottom') | Should -Be '0.5'
    }

    It 'wraps OfficeIMO worksheet operations and print definitions' {
        $path = Join-Path $TestDrive 'ExcelWorksheetOperations.xlsx'
        $sourcePath = Join-Path $TestDrive 'ExcelWorksheetOperationsSource.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
            Add-OfficeExcelSheet -Name 'More' -Content {
                Add-OfficeExcelTable -InputObject $moreRows -TableName 'MoreSales'
            }
        }
        New-OfficeExcel -Path $sourcePath {
            Add-OfficeExcelSheet -Name 'External' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Name'
                Set-OfficeExcelCell -Address 'A2' -Value 'Imported'
            }
        }

        Copy-OfficeExcelSheet -Path $path -SourceSheet 'Data' -NewName 'DataCopy' | Should -Not -BeNullOrEmpty
        Move-OfficeExcelSheet -Path $path -Sheet 'DataCopy' -Index 0
        Copy-OfficeExcelSheet -Path $path -SourcePath $sourcePath -SourceSheet 'External' -NewName 'ExternalCopy' | Should -Not -BeNullOrEmpty
        $join = Join-OfficeExcelSheet -Path $path -TargetSheet 'Data' -SourceSheet 'More' -MatchColumnsByHeader
        Set-OfficeExcelPrintArea -Path $path -Sheet 'Data' -Range 'A1:B4'
        Set-OfficeExcelPrintTitles -Path $path -Sheet 'Data' -FirstRow 1 -LastRow 1

        $join.RowsCopied | Should -Be 1
        $join.TargetSheetName | Should -Be 'Data'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.Sheets[0].Name | Should -Be 'DataCopy'
        $summary.Sheets.Name | Should -Contain 'ExternalCopy'

        $external = @(Import-OfficeExcel -Path $path -WorksheetName 'ExternalCopy' -Range 'A1:A2')
        $external.Count | Should -Be 1
        $external[0].Name | Should -Be 'Imported'

        $merged = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $merged.Count | Should -Be 3
        $merged[2].Region | Should -Be 'APAC'

        $differences = @(Compare-OfficeExcelRange -Path $path -LeftSheet 'Data' -RightSheet 'DataCopy')
        $differences.Count | Should -BeGreaterThan 0

        $names = @(Get-OfficeExcelNamedRange -Path $path -Sheet 'Data')
        @($names | Where-Object Name -eq '_xlnm.Print_Area').Count | Should -Be 1
        @($names | Where-Object Name -eq '_xlnm.Print_Titles').Count | Should -Be 1
    }

    It 'finds, replaces, and edits Excel row values' {
        $path = Join-Path $TestDrive 'ExcelFindReplaceEditRows.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Ada'; Status = 'Draft' }
            [PSCustomObject]@{ Name = 'Grace'; Status = 'Draft' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'People'
            }
        }

        @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text 'Draft').Count | Should -Be 2
        Update-OfficeExcelText -Path $path -Sheet 'Data' -OldValue 'Draft' -NewValue 'Ready' | Should -Be 2
        Edit-OfficeExcelRow -Path $path -Sheet 'Data' -ScriptBlock {
            param($row)
            if ($row.CellByHeader('Name').Value -eq 'Ada') {
                $row.Set('Status', 'Done')
            }
        }

        $updated = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $updated[0].Status | Should -Be 'Done'
        $updated[1].Status | Should -Be 'Ready'
        @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text '^Done$' -Regex).Count | Should -Be 1
    }

    It 'counts threaded comments in workbook summaries' {
        $path = Join-Path $TestDrive 'DslExcelThreadedComments.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        $spreadsheetDocumentType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.SpreadsheetDocument'
        $worksheetPartType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.WorksheetPart'
        $threadedPartType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.WorksheetThreadedCommentsPart'
        $threadedCommentsType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComments'
        $threadedCommentType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComment'
        $threadedCommentTextType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedCommentText'

        $openMethod = $spreadsheetDocumentType.GetMethod('Open', [type[]] @([string], [bool]))
        $openArguments = [object[]] @($path.ToString(), $true)
        $document = $openMethod.Invoke($null, $openArguments)
        try {
            $worksheetPart = @($document.WorkbookPart.WorksheetParts)[0]
            $addPartMethod = $worksheetPartType.GetMethods() |
                Where-Object { $_.Name -eq 'AddNewPart' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0 } |
                Select-Object -First 1
            $threadedPart = $addPartMethod.MakeGenericMethod($threadedPartType).Invoke($worksheetPart, @())
            $threadedComments = [Activator]::CreateInstance($threadedCommentsType)
            $threadedComment = [Activator]::CreateInstance($threadedCommentType)
            $threadedComment.Ref = 'A2'
            $threadedComment.PersonId = '{00000000-0000-0000-0000-000000000001}'
            $threadedComment.Id = '{00000000-0000-0000-0000-000000000002}'
            $threadedCommentTextConstructor = $threadedCommentTextType.GetConstructor([type[]] @([string]))
            $threadedComment.AppendChild($threadedCommentTextConstructor.Invoke([object[]] @('Modern note'))) | Out-Null
            $threadedComments.AppendChild($threadedComment) | Out-Null
            $threadedComments.Save($threadedPart)
        } finally {
            $document.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.CommentCount | Should -Be 1
        $summary.Sheets[0].CommentCount | Should -Be 1
    }

    It 'adds a table of contents and reads ranges with the new Excel readers' {
        $path = Join-Path $TestDrive 'DslExcelNavigation.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelNamedRange -Name 'SalesData' -Range 'A1:B3'
            }
            Add-OfficeExcelSheet -Name 'Notes' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Label', 'Value'
                Set-OfficeExcelRow -Row 2 -Values 'Generated', 'Yes'
            }
        } | Out-Null

        $usedRange = Get-OfficeExcelUsedRange -Path $path -Sheet 'Data' -AsDataTable
        $usedRange.Rows.Count | Should -Be 2
        $usedRange.Columns[0].ColumnName | Should -Be 'Region'
        $usedRange.Rows[0]['Region'] | Should -Be 'NA'

        Add-OfficeExcelTableOfContents -Path $path -IncludeNamedRanges -AddBackLinks

        $tocRows = @(Get-OfficeExcelRange -Path $path -Sheet 'TOC' -Range 'A3:C5' -AsHashtable)
        $tocRows.Count | Should -Be 2
        $tocRows[0]['Sheet'] | Should -Be 'Data'
        $tocRows[0]['Named Ranges'] | Should -Match 'SalesData'
        $tocRows[1]['Sheet'] | Should -Be 'Notes'

        $noteRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Notes' -Range 'A1:B2')
        $noteRows.Count | Should -Be 1
        $noteRows[0].Label | Should -Be 'Generated'
        $noteRows[0].Value | Should -Be 'Yes'

        $dataRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Data' -Range 'A1:B3')
        $dataRows.Count | Should -Be 2
        $dataRows[0].Region | Should -Be 'NA'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 3
        $summary.VisibleSheetCount | Should -Be 3
        $summary.TableCount | Should -Be 2
        $summary.NamedRangeCount | Should -Be 1
        $summary.HyperlinkCount | Should -BeGreaterThan 0
        $summary.Sheets.Name | Should -Contain 'Data'
        ($summary.Sheets | Where-Object Name -eq 'Data').UsedRange | Should -Be 'A1:B5'
        ($summary.Sheets | Where-Object Name -eq 'Data').Tables.Name | Should -Contain 'Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'TOC'

            $backLink = $null
            $doc['Data'].TryGetCellText(5, 1, [ref] $backLink) | Should -BeTrue
            $backLink | Should -Be "$([char]0x2190) TOC"
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'includes chartsheet charts in workbook summaries' {
        $path = Join-Path $TestDrive 'WorkbookWithChartSheet.xlsx'
        $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            function Add-ZipTextEntry {
                param(
                    [Parameter(Mandatory)]
                    [System.IO.Compression.ZipArchive] $Archive,

                    [Parameter(Mandatory)]
                    [string] $EntryName,

                    [Parameter(Mandatory)]
                    [string] $Content
                )

                $entry = $Archive.CreateEntry($EntryName)
                $stream = $entry.Open()
                try {
                    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
                    try {
                        $writer.Write($Content)
                    } finally {
                        $writer.Dispose()
                    }
                } finally {
                    $stream.Dispose()
                }
            }

            Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>
'@
            Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Revenue Chart" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/worksheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Region</v></c><c r="B1" t="str"><v>Revenue</v></c></row>
    <row r="2"><c r="A2" t="str"><v>EMEA</v></c><c r="B2"><v>42</v></c></row>
  </sheetData>
</worksheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <drawing r:id="rId1"/>
</chartsheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/_rels/sheet1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/drawing1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/><xdr:ext cx="6000000" cy="4000000"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/_rels/drawing1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/charts/chart1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>
</c:chartSpace>
'@
        } finally {
            $archive.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 2
        $summary.ChartCount | Should -Be 1
        ($summary.Sheets | Where-Object Name -eq 'Revenue Chart').ChartCount | Should -Be 1
    }

    It 'formats Excel charts with legend, labels, and style presets' {
        $path = Join-Path $TestDrive 'DslExcelChartFormatting.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Pie -Title 'Revenue Mix' -PassThru
                $formattedChart = $chart |
                    Set-OfficeExcelChartLegend -Position Right |
                    Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -Position OutsideEnd -NumberFormat '0.0%' -SourceLinked:$false |
                    Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10

                $formattedChart | Should -Not -BeNullOrEmpty
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/style*.xml' }).Count | Should -BeGreaterThan 0
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/colors*.xml' }).Count | Should -BeGreaterThan 0

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $legendPosition = $chartXml.SelectSingleNode("/*[local-name()='chartSpace']/*[local-name()='chart']/*[local-name()='legend']/*[local-name()='legendPos']")
        $legendPosition | Should -Not -BeNullOrEmpty
        $legendPosition.GetAttribute('val') | Should -Be 'r'

        $dataLabels = $chartXml.SelectSingleNode("//*[local-name()='dLbls']")
        $dataLabels | Should -Not -BeNullOrEmpty
        $dataLabels.SelectSingleNode("*[local-name()='showVal']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='showPercent']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='dLblPos']").GetAttribute('val') | Should -Be 'outEnd'

        $numberFormat = $dataLabels.SelectSingleNode("*[local-name()='numFmt']")
        $numberFormat | Should -Not -BeNullOrEmpty
        $numberFormat.GetAttribute('formatCode') | Should -Be '0.0%'
    }

    It 'formats Excel chart axes series and trendlines' {
        $path = Join-Path $TestDrive 'DslExcelChartAxisSeriesTrendline.xlsx'
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; Revenue = 100 }
            [PSCustomObject]@{ Month = 'Feb'; Revenue = 200 }
            [PSCustomObject]@{ Month = 'Mar'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Line -Title 'Revenue Trend' -PassThru
                { $chart | Set-OfficeExcelChartSeries -SeriesIndex 0 -LineWidthPoints 1.5 -ErrorAction Stop } |
                    Should -Throw '*LineColor is required*'
                $formattedChart = $chart |
                    Set-OfficeExcelChartAxis -CategoryTitle 'Month' -ValueTitle 'Revenue' -ValueNumberFormat '$#,##0' -SourceLinked:$false -ValueMinimum 0 -ValueMajorUnit 100 -ShowValueMinorGridlines -ValueGridlineColor '#D9EAD3' -GridlineWidthPoints 0.75 |
                    Set-OfficeExcelChartSeries -SeriesIndex 0 -LineColor '#1F4E79' -LineWidthPoints 1.5 -MarkerStyle Circle -MarkerSize 6 -MarkerFillColor '#4472C4' |
                    Set-OfficeExcelChartTrendline -SeriesIndex 0 -Type Linear -DisplayEquation -DisplayRSquared -LineColor '#C00000' -LineWidthPoints 1.25

                $formattedChart | Should -Not -BeNullOrEmpty
            }
        } | Out-Null

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $chartOuterXml = $chartXml.OuterXml

        $categoryAxis = $chartXml.SelectSingleNode("//*[local-name()='catAx']")
        $categoryTitle = $categoryAxis.SelectSingleNode("*[local-name()='title']")
        $categoryTitle | Should -Not -BeNullOrEmpty
        $categoryTitle.InnerText | Should -Be 'Month'
        $categoryAxis.SelectSingleNode("*[local-name()='majorGridlines']") | Should -BeNullOrEmpty
        $categoryAxis.SelectSingleNode("*[local-name()='minorGridlines']") | Should -BeNullOrEmpty

        $valueAxis = $chartXml.SelectSingleNode("//*[local-name()='valAx']")
        $valueAxis | Should -Not -BeNullOrEmpty
        $valueAxis.SelectSingleNode("*[local-name()='title']").InnerText | Should -Be 'Revenue'
        $valueAxis.SelectSingleNode("*[local-name()='numFmt']").GetAttribute('formatCode') | Should -Be '$#,##0'
        $valueAxis.SelectSingleNode("*[local-name()='scaling']/*[local-name()='min']").GetAttribute('val') | Should -Be '0'
        $valueAxis.SelectSingleNode("*[local-name()='majorUnit']").GetAttribute('val') | Should -Be '100'
        $valueAxis.SelectSingleNode("*[local-name()='majorGridlines']") | Should -Not -BeNullOrEmpty
        $valueAxis.SelectSingleNode("*[local-name()='minorGridlines']") | Should -Not -BeNullOrEmpty

        $chartOuterXml | Should -Match 'trendline'
        $chartOuterXml | Should -Match 'dispEq'
        $chartOuterXml | Should -Match 'dispRSqr'
        $chartOuterXml | Should -Match '1F4E79'
        $chartOuterXml | Should -Match '4472C4'
        $chartOuterXml | Should -Match 'C00000'
    }

    It 'supports url images and smart hyperlink helpers' {
        $path = Join-Path $TestDrive 'DslExcelLinksAndImages.xlsx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $imageUrl = [System.Uri]::new($imagePath).AbsoluteUri

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Reference'
                Set-OfficeExcelCell -Address 'B1' -Value 'Host'
                Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208'
                Set-OfficeExcelHostHyperlink -Address 'B2' -Url 'https://learn.microsoft.com/office/open-xml/'
                Add-OfficeExcelImageFromUrl -Address 'D2' -Url $imageUrl -WidthPixels 32 -HeightPixels 32
                Add-OfficeExcelImage -Address 'E2' -Url $imageUrl -WidthPixels 24 -HeightPixels 24
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/media/*' }).Count | Should -BeGreaterThan 0

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $smartText = $null
            $hostText = $null
            $doc['Data'].TryGetCellText(2, 1, [ref] $smartText) | Should -BeTrue
            $doc['Data'].TryGetCellText(2, 2, [ref] $hostText) | Should -BeTrue
            $smartText | Should -Be 'RFC 7208'
            $hostText | Should -Be 'learn.microsoft.com'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 2
    }

    It 'supports internal link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelInternalLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ Sheet = 'Alpha'; Target = 'Alpha' }
            [PSCustomObject]@{ Sheet = 'Beta'; Target = 'Beta' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'SummaryTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Sheet'
                Set-OfficeExcelCell -Address 'D2' -Value 'Alpha'
                Set-OfficeExcelCell -Address 'D3' -Value 'Beta'
                Set-OfficeExcelInternalLinks -Range 'D2:D3'
                Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) "Open $text" }
            }
            Add-OfficeExcelSheet -Name 'Alpha' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Alpha Home'
            }
            Add-OfficeExcelSheet -Name 'Beta' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Beta Home'
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open Alpha'
            $tableLink2 | Should -Be 'Open Beta'
            $rangeLink1 | Should -Be 'Alpha'
            $rangeLink2 | Should -Be 'Beta'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }

    It 'supports external URL link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelUrlLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ RFC = 'rfc7208'; Spec = 'rfc5321' }
            [PSCustomObject]@{ RFC = 'rfc7489'; Spec = 'rfc1035' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'LinksTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Spec'
                Set-OfficeExcelCell -Address 'D2' -Value 'rfc5321'
                Set-OfficeExcelCell -Address 'D3' -Value 'rfc1035'

                Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" }
                Set-OfficeExcelUrlLinks -Range 'D2:D3' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" }
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open rfc7208'
            $tableLink2 | Should -Be 'Open rfc7489'
            $rangeLink1 | Should -Be 'RFC 5321'
            $rangeLink2 | Should -Be 'RFC 1035'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }

    It 'styles Excel columns by header without range math' {
        $path = Join-Path $TestDrive 'DslExcelColumnStyleByHeader.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Revenue = 1200.5; Rate = 0.42; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Revenue = 800.25; Rate = 0.18; Status = 'Blocked' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'ReportRows'
                Set-OfficeExcelColumnStyleByHeader -Header Revenue -Style Currency -CultureName en-US -AutoFit
                Set-OfficeExcelColumnStyleByHeader -Header Rate -Style Percent -Decimals 1
                Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' } -BoldByText Blocked
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $revenueCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B2']")
        $rateCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='C2']")
        $statusCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='D3']")

        $revenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $rateCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $statusCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
    }

    It 'preserves case-distinct text style map entries when requested' {
        $path = Join-Path $TestDrive 'DslExcelColumnStyleByHeaderCaseSensitive.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'ready' }
        )
        $statusColors = [hashtable]::new([System.StringComparer]::Ordinal)
        $statusColors.Add('Ready', '#D4EDDA')
        $statusColors.Add('ready', '#F8D7DA')

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'ReportRows'
                Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText $statusColors -CaseSensitive
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $upperCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B2']")
        $lowerCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")

        $upperCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $lowerCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $upperCell.GetAttribute('s') | Should -Not -Be $lowerCell.GetAttribute('s')
    }

    It 'creates composed Excel report sheets from PowerShell blocks' {
        $path = Join-Path $TestDrive 'DslExcelReportSheet.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Score = 9; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Score = 4; Status = 'Blocked' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Summary' {
                Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current view'
                Add-OfficeExcelReportKpiRow -Data ([ordered] @{ Ready = 1; Blocked = 1 }) -PerRow 2
                Add-OfficeExcelReportCallout -Kind Warning -Title 'Attention' -Body 'One item needs review.'
                Add-OfficeExcelReportTable -Data $rows -Title 'Rows'
                Add-OfficeExcelReportLegend -Title 'Legend' -Headers 'Status','Meaning' -Rows @(
                    @('Ready', 'No action'),
                    @('Blocked', 'Needs owner')
                ) -FirstColumnFillByValue @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' }
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $sheet = $doc['Summary']
            $title = $null
            $subtitle = $null
            $readyLabel = $null
            $calloutTitle = $null

            $sheet.TryGetCellText(1, 1, [ref] $title) | Should -BeTrue
            $sheet.TryGetCellText(2, 1, [ref] $subtitle) | Should -BeTrue
            $sheet.TryGetCellText(4, 1, [ref] $readyLabel) | Should -BeTrue
            $sheet.TryGetCellText(7, 1, [ref] $calloutTitle) | Should -BeTrue

            $title | Should -Be 'Operational Summary'
            $subtitle | Should -Be 'Current view'
            $readyLabel | Should -Be 'Ready'
            $calloutTitle | Should -Be 'Attention'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'preserves case-distinct report legend fill entries when requested' {
        $path = Join-Path $TestDrive 'DslExcelReportLegendCaseSensitive.xlsx'
        $statusColors = [hashtable]::new([System.StringComparer]::Ordinal)
        $statusColors.Add('Ready', '#D4EDDA')
        $statusColors.Add('ready', '#F8D7DA')

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Legend' {
                Add-OfficeExcelReportLegend -Header 'Status','Meaning' -InputObject @(
                    @('Ready', 'Upper'),
                    @('ready', 'Lower')
                ) -FirstColumnFillByValue $statusColors -CaseSensitive
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $upperCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='A2']")
        $lowerCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='A3']")

        $upperCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $lowerCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $upperCell.GetAttribute('s') | Should -Not -Be $lowerCell.GetAttribute('s')
    }

    It 'uses the topmost report composer for nested report sheets' {
        $path = Join-Path $TestDrive 'DslExcelNestedReportSheets.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Outer' {
                Add-OfficeExcelReportTitle -Title 'Outer title'
                Add-OfficeExcelReportSheet -Name 'Inner' {
                    Add-OfficeExcelReportTitle -Title 'Inner title'
                }
                Add-OfficeExcelReportParagraph -Text 'Outer after inner'
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $outer = $doc['Outer']
            $inner = $doc['Inner']
            $outerTitle = $null
            $outerAfter = $null
            $innerTitle = $null

            $outer.TryGetCellText(1, 1, [ref] $outerTitle) | Should -BeTrue
            $outer.TryGetCellText(3, 1, [ref] $outerAfter) | Should -BeTrue
            $inner.TryGetCellText(1, 1, [ref] $innerTitle) | Should -BeTrue

            $outerTitle | Should -Be 'Outer title'
            $outerAfter | Should -Be 'Outer after inner'
            $innerTitle | Should -Be 'Inner title'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'uses the topmost worksheet for nested sheet blocks' {
        $path = Join-Path $TestDrive 'DslExcelNestedSheets.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Outer' {
                Set-OfficeExcelCell -Address 'A1' -Value 'Outer start'
                Add-OfficeExcelSheet -Name 'Inner' {
                    Set-OfficeExcelCell -Address 'A1' -Value 'Inner value'
                }
                Set-OfficeExcelCell -Address 'A2' -Value 'Outer after inner'
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $outer = $doc['Outer']
            $inner = $doc['Inner']
            $outerStart = $null
            $outerAfter = $null
            $innerValue = $null

            $outer.TryGetCellText(1, 1, [ref] $outerStart) | Should -BeTrue
            $outer.TryGetCellText(2, 1, [ref] $outerAfter) | Should -BeTrue
            $inner.TryGetCellText(1, 1, [ref] $innerValue) | Should -BeTrue

            $outerStart | Should -Be 'Outer start'
            $outerAfter | Should -Be 'Outer after inner'
            $innerValue | Should -Be 'Inner value'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }
}
