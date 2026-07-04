function Get-ExcelBenchmarkData {
    param(
        [string] $Profile,
        [int] $Count
    )

    switch ($Profile) {
        MixedObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkRows -Count $Count); ColumnCount = 10; WorksheetName = 'Data' }
        }
        WideObjects {
            [pscustomobject]@{ Data = @(New-ExcelBenchmarkWideRows -Count $Count); ColumnCount = 40; WorksheetName = 'Data' }
        }
        DataTable {
            [pscustomobject]@{ Data = New-ExcelBenchmarkDataTable -Count $Count; ColumnCount = 6; WorksheetName = 'Data' }
        }
        DataSet {
            [pscustomobject]@{ Data = New-ExcelBenchmarkDataSet -Count $Count; ColumnCount = 6; WorksheetName = 'Sales' }
        }
        default { throw "Unknown benchmark profile '$Profile'." }
    }
}

function Get-ExcelBenchmarkPayloadCount {
    param([object] $Value)

    if ($Value -is [Data.DataSet]) {
        $count = 0
        foreach ($table in $Value.Tables) {
            $count += $table.Rows.Count
        }
        return $count
    }

    if ($Value -is [Data.DataTable]) {
        return $Value.Rows.Count
    }

    return @($Value).Count
}

function New-ExcelBenchmarkRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id = $i
            Name = 'Server-{0:000000}' -f $i
            Department = 'Department-{0}' -f ($i % 25)
            Region = @('NA', 'EU', 'APAC', 'LATAM')[$i % 4]
            IsEnabled = ($i % 3) -ne 0
            Created = ([datetime]'2024-01-01').AddMinutes($i)
            Score = [math]::Round(($i * 1.137) % 1000, 3)
            Owner = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = $i % 17
            Notes = 'Benchmark row {0}' -f $i
        }
    }
}

function New-ExcelBenchmarkWideRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        $row = [ordered]@{
            Id = $i
            Name = 'Wide-{0:000000}' -f $i
            Created = ([datetime]'2024-01-01').AddSeconds($i)
            Enabled = ($i % 2) -eq 0
        }
        for ($column = 1; $column -le 36; $column++) {
            $row["Metric$column"] = [math]::Round((($i + $column) * 1.017) % 10000, 4)
        }
        [pscustomobject]$row
    }
}

function New-ExcelBenchmarkDataTable {
    param([int] $Count)

    $table = [Data.DataTable]::new('Data')
    $null = $table.Columns.Add('Id', [int])
    $null = $table.Columns.Add('Name', [string])
    $null = $table.Columns.Add('Created', [datetime])
    $null = $table.Columns.Add('Amount', [decimal])
    $null = $table.Columns.Add('Enabled', [bool])
    $null = $table.Columns.Add('Notes', [string])
    for ($i = 1; $i -le $Count; $i++) {
        $row = $table.NewRow()
        $row.Id = $i
        $row.Name = 'Account-{0:000000}' -f $i
        $row.Created = ([datetime]'2024-01-01').AddMinutes($i)
        $row.Amount = [decimal]([math]::Round(($i * 11.317) % 100000, 2))
        $row.Enabled = ($i % 4) -ne 0
        $row.Notes = 'DataTable row {0}' -f $i
        $table.Rows.Add($row)
    }
    , $table
}

function New-ExcelBenchmarkDataSet {
    param([int] $Count)

    $dataSet = [Data.DataSet]::new('Report')
    $sales = New-ExcelBenchmarkDataTable -Count $Count
    $sales.TableName = 'Sales'
    $inventory = [Data.DataTable]::new('Inventory')
    $null = $inventory.Columns.Add('Sku', [string])
    $null = $inventory.Columns.Add('Quantity', [int])
    $null = $inventory.Columns.Add('Updated', [datetime])
    for ($i = 1; $i -le ([math]::Max(1, [math]::Floor($Count / 4))); $i++) {
        $row = $inventory.NewRow()
        $row.Sku = 'SKU-{0:000000}' -f $i
        $row.Quantity = $i % 500
        $row.Updated = ([datetime]'2024-01-01').AddHours($i)
        $inventory.Rows.Add($row)
    }
    $dataSet.Tables.Add($sales)
    $dataSet.Tables.Add($inventory)
    , $dataSet
}

function Get-ExcelBenchmarkColumnName {
    param([int] $ColumnNumber)

    $name = ''
    $value = $ColumnNumber
    while ($value -gt 0) {
        $value--
        $name = [char][int](65 + ($value % 26)) + $name
        $value = [int][math]::Floor($value / 26)
    }
    $name
}

function Get-ExcelBenchmarkColumnCount {
    param([string] $Profile)

    switch ($Profile) {
        WideObjects { 40 }
        DataTable { 6 }
        DataSet { 6 }
        default { 10 }
    }
}

function Get-ExcelBenchmarkRange {
    param([int] $ColumnCount, [int] $Rows)

    'A1:{0}{1}' -f (Get-ExcelBenchmarkColumnName -ColumnNumber $ColumnCount), ($Rows + 1)
}

function Get-ExcelBenchmarkRangeEndCell {
    param([int] $ColumnCount, [int] $Rows)

    '{0}{1}' -f (Get-ExcelBenchmarkColumnName -ColumnNumber $ColumnCount), ($Rows + 1)
}

function Get-ExcelBenchmarkExtension {
    param([object] $Case)

    $extension = [string]$Case.FileExtension
    if ([string]::IsNullOrWhiteSpace($extension)) { return '.xlsx' }
    if ($extension.StartsWith('.')) { return $extension }
    ".$extension"
}

function Initialize-ExcelBenchmarkInput {
    param([object] $Case, [object] $Run)

    foreach ($path in @($Run.Path, $Run.SourcePath)) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force
        }
    }

    switch ([string]$Case.OperationKey) {
        ReadCsvSource {
            $Run.Payload | Export-Csv -Path $Run.SourcePath -NoTypeInformation -Encoding utf8
        }
        CsvToExcel {
            $Run.Payload | Export-Csv -Path $Run.SourcePath -NoTypeInformation -Encoding utf8
        }
        ReadFullSheet { New-ExcelBenchmarkDefaultWorkbook -Engine $Case.Engine -Run $Run }
        ReadRange { New-ExcelBenchmarkDefaultWorkbook -Engine $Case.Engine -Run $Run }
        ReadNoHeaderRange { New-ExcelBenchmarkDefaultWorkbook -Engine $Case.Engine -Run $Run }
        ReadUsedRangeDataTable { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run }
        ReadTableMetadata { New-ExcelBenchmarkDefaultWorkbook -Engine PSWriteOffice -Run $Run -TableName Data }
        ReadNamedRangeMetadata { New-ExcelBenchmarkNamedRangeWorkbook -Engine PSWriteOffice -Run $Run }
    }
}

function New-ExcelBenchmarkDefaultWorkbook {
    param(
        [string] $Engine,
        [object] $Run,
        [string] $TableName
    )

    switch ($Engine) {
        PSWriteOffice {
            if ($TableName) {
                $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName $TableName
            } else {
                $Run.Payload | Export-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName
            }
        }
        ImportExcel {
            if ($TableName) {
                $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName $TableName -AutoFilter
            } else {
                $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName
            }
        }
        ExcelFast {
            Export-Workbook -Destination $Run.Path -InputObject $Run.Payload -SheetName $Run.WorksheetName -Force
        }
        default {
            throw "Engine '$Engine' cannot create workbook input."
        }
    }
}

function New-ExcelBenchmarkNamedRangeWorkbook {
    param([string] $Engine, [object] $Run)

    switch ($Engine) {
        PSWriteOffice {
            New-OfficeExcel -Path $Run.Path {
                Add-OfficeExcelSheet -Name $Run.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Run.Payload -TableName Data
                    Set-OfficeExcelNamedRange -Name SalesData -Range $Run.Range
                }
            } | Out-Null
        }
        ImportExcel {
            $Run.Payload | Export-Excel -Path $Run.Path -WorksheetName $Run.WorksheetName -TableName Data -AutoFilter -RangeName SalesData
        }
    }
}
